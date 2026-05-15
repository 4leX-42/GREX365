# Retry module
# Centralized retry policy for Graph / EXO calls.
# - HTTP 429 (TooManyRequests): honour Retry-After header when present, else exponential backoff with jitter.
# - HTTP 5xx (server side): exponential backoff with jitter.
# - HTTP 4xx (other): no retry.
# - Network / generic exceptions classified by message pattern fallback.

function Get-RetryDelayMs {
    param(
        [int]$Attempt,
        [int]$BaseDelayMs,
        [int]$MaxDelayMs
    )
    $exp = [math]::Pow(2, [math]::Max(0, $Attempt - 1))
    $raw = [int]([math]::Min($MaxDelayMs, $BaseDelayMs * $exp))
    $jitter = Get-Random -Minimum 0 -Maximum 200
    return ($raw + $jitter)
}

function Get-ExceptionHttpStatus {
    param($ErrorRecord)

    $status = $null
    $retryAfterMs = $null

    try {
        $ex = $ErrorRecord.Exception
        if ($ex -and $ex.PSObject.Properties.Name -contains 'Response' -and $ex.Response) {
            $resp = $ex.Response
            if ($resp.StatusCode) { $status = [int]$resp.StatusCode }
            try {
                if ($resp.Headers -and $resp.Headers.RetryAfter) {
                    if ($resp.Headers.RetryAfter.Delta) {
                        $retryAfterMs = [int]$resp.Headers.RetryAfter.Delta.TotalMilliseconds
                    } elseif ($resp.Headers.RetryAfter.Date) {
                        $diff = ($resp.Headers.RetryAfter.Date - (Get-Date)).TotalMilliseconds
                        if ($diff -gt 0) { $retryAfterMs = [int]$diff }
                    }
                }
            } catch {}
        }
    } catch {}

    if (-not $status) {
        $msg = [string]$ErrorRecord.Exception.Message
        if ($msg -match '\b429\b|TooManyRequests|throttl|too\s+many\s+requests') { $status = 429 }
        elseif ($msg -match '\b50[0-4]\b|InternalServerError|BadGateway|ServiceUnavailable|GatewayTimeout|server\s+error') { $status = 503 }
    }

    return [PSCustomObject]@{ Status = $status; RetryAfterMs = $retryAfterMs }
}

function Test-IsRetryableException {
    param($ErrorRecord)

    $info = Get-ExceptionHttpStatus -ErrorRecord $ErrorRecord
    if (-not $info.Status) { return $false }
    if ($info.Status -eq 429) { return $true }
    if ($info.Status -ge 500 -and $info.Status -lt 600) { return $true }
    return $false
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 4,
        [int]$BaseDelayMs = 500,
        [int]$MaxDelayMs  = 30000,
        [string]$OperationName = 'operation',
        [switch]$Quiet
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $ScriptBlock
        } catch {
            $err  = $_
            $info = Get-ExceptionHttpStatus -ErrorRecord $err

            $retry = $false
            $waitMs = 0
            $reason = ''

            if ($info.Status -eq 429) {
                $retry = $true
                if ($info.RetryAfterMs) {
                    $waitMs = $info.RetryAfterMs + (Get-Random -Minimum 0 -Maximum 200)
                } else {
                    $waitMs = Get-RetryDelayMs -Attempt $attempt -BaseDelayMs $BaseDelayMs -MaxDelayMs $MaxDelayMs
                }
                $reason = 'HTTP 429 throttle'
            } elseif ($info.Status -ge 500 -and $info.Status -lt 600) {
                $retry = $true
                $waitMs = Get-RetryDelayMs -Attempt $attempt -BaseDelayMs $BaseDelayMs -MaxDelayMs $MaxDelayMs
                $reason = "HTTP $($info.Status) server error"
            }

            if ($retry -and $attempt -lt $MaxAttempts) {
                if (-not $Quiet -and (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
                    Write-Log ("Retry {0}/{1} '{2}' tras {3} ({4}ms)" -f $attempt, $MaxAttempts, $OperationName, $reason, $waitMs) -Level WARN -Source 'Retry'
                }
                Start-Sleep -Milliseconds $waitMs
                continue
            }
            throw
        }
    }
}
