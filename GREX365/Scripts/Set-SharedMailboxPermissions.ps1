#requires -Version 7.4
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Permisos sobre buzón compartido' } catch {}

# Bulk delegation manager for shared/user mailboxes.
# Operations supported:
#   - Grant or revoke FullAccess
#   - Grant or revoke SendAs
#   - Grant or revoke SendOnBehalf
# CSV schema:  Action;Permission;Mailbox;Principal
#   Action     = add | remove
#   Permission = FullAccess | SendAs | SendOnBehalf
#   Mailbox    = SMTP / UPN del buzón objetivo
#   Principal  = SMTP / UPN del usuario que recibe / pierde el permiso

Assert-RequiredServicesReady
Assert-Role -Required operator -Operation 'gestionar permisos de buzón'

Show-Header -Title 'GREX365' -Subtitle 'Permisos sobre buzón compartido'

$session = Start-LogSession -Name 'SharedMailboxPermissions'
$started = Get-Date

$results = New-Object System.Collections.Generic.List[object]

function Add-PermResult {
    param(
        [string]$Action,[string]$Permission,[string]$Mailbox,[string]$Principal,
        [string]$Estado,[string]$Detalle = ''
    )
    $results.Add([PSCustomObject]@{
        Action      = $Action
        Permission  = $Permission
        Mailbox     = $Mailbox
        Principal   = $Principal
        Estado      = $Estado
        Detalle     = $Detalle
    })
}

try {
    $csvPath = Read-ValidatedCsvPath -Prompt 'CSV (Action;Permission;Mailbox;Principal)'

    $rows = @(Import-Csv -Path $csvPath -Delimiter ';' -Encoding UTF8)
    if (-not $rows -or $rows.Count -eq 0) {
        throw 'CSV vacío.'
    }

    $required = @('Action','Permission','Mailbox','Principal')
    foreach ($col in $required) {
        if ($rows[0].PSObject.Properties.Name -notcontains $col) {
            throw "Columna requerida ausente: $col"
        }
    }

    Show-Section -Title 'Operación'
    Write-KeyValue -Key 'CSV'    -Value $csvPath
    Write-KeyValue -Key 'Filas'  -Value $rows.Count

    if (-not (Confirm-DestructiveAction -Operation 'Aplicar permisos del CSV' -Target ("$($rows.Count) filas"))) {
        throw 'INPUT_EMPTY'
    }

    $idx = 0
    foreach ($row in $rows) {
        $idx++
        $action = ([string]$row.Action).Trim().ToLowerInvariant()
        $perm   = ([string]$row.Permission).Trim()
        $mbx    = ([string]$row.Mailbox).Trim()
        $prn    = ([string]$row.Principal).Trim()

        $pct = [math]::Round(($idx / $rows.Count) * 100, 1)
        Write-Progress -Activity 'Aplicando permisos' -Status "Fila $idx/$($rows.Count)" -PercentComplete $pct

        if (-not $mbx -or -not $prn) {
            Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'INVALIDO' -Detalle 'Mailbox o Principal vacío'
            continue
        }
        if ($action -notin @('add','remove')) {
            Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'INVALIDO' -Detalle "Action no soportada: $action"
            continue
        }
        if ($perm -notin @('FullAccess','SendAs','SendOnBehalf')) {
            Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'INVALIDO' -Detalle "Permission no soportado: $perm"
            continue
        }

        try {
            switch ("$perm-$action") {
                'FullAccess-add' {
                    Invoke-WithRetry -OperationName 'Add-MailboxPermission FullAccess' -Quiet -ScriptBlock {
                        Add-MailboxPermission -Identity $mbx -User $prn -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'OK'
                }
                'FullAccess-remove' {
                    Invoke-WithRetry -OperationName 'Remove-MailboxPermission FullAccess' -Quiet -ScriptBlock {
                        Remove-MailboxPermission -Identity $mbx -User $prn -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'OK'
                }
                'SendAs-add' {
                    Invoke-WithRetry -OperationName 'Add-RecipientPermission SendAs' -Quiet -ScriptBlock {
                        Add-RecipientPermission -Identity $mbx -Trustee $prn -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'OK'
                }
                'SendAs-remove' {
                    Invoke-WithRetry -OperationName 'Remove-RecipientPermission SendAs' -Quiet -ScriptBlock {
                        Remove-RecipientPermission -Identity $mbx -Trustee $prn -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    }
                    Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'OK'
                }
                'SendOnBehalf-add' {
                    Invoke-WithRetry -OperationName 'Set-Mailbox GrantSendOnBehalfTo @{add}' -Quiet -ScriptBlock {
                        Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Add = $prn} -ErrorAction Stop
                    }
                    Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'OK'
                }
                'SendOnBehalf-remove' {
                    Invoke-WithRetry -OperationName 'Set-Mailbox GrantSendOnBehalfTo @{remove}' -Quiet -ScriptBlock {
                        Set-Mailbox -Identity $mbx -GrantSendOnBehalfTo @{Remove = $prn} -ErrorAction Stop
                    }
                    Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'OK'
                }
            }
        } catch {
            Add-PermResult -Action $action -Permission $perm -Mailbox $mbx -Principal $prn -Estado 'ERROR' -Detalle $_.Exception.Message
        }
    }
    Write-Progress -Activity 'Aplicando permisos' -Completed

    $ok    = @($results | Where-Object Estado -eq 'OK').Count
    $bad   = @($results | Where-Object Estado -eq 'ERROR').Count
    $inv   = @($results | Where-Object Estado -eq 'INVALIDO').Count

    Show-Section -Title 'Resumen'
    Write-KeyValue -Key 'Aplicados' -Value $ok -ValueColor Green
    Write-KeyValue -Key 'Errores'   -Value $bad -ValueColor Red
    Write-KeyValue -Key 'Inválidos' -Value $inv -ValueColor Yellow

    $outFolder = Split-Path -Path $csvPath -Parent
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $resCsv = Join-Path $outFolder ("permissions_result_$stamp.csv")
    $results | Export-Csv -Path $resCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-KeyValue -Key 'CSV resultado' -Value $resCsv

    $logFile = Stop-LogSession -Persist
    $report = Publish-OperationReport `
        -Title 'Bulk Mailbox Permissions' `
        -Operation 'Set-SharedMailboxPermissions' `
        -StartTime $started -EndTime (Get-Date) `
        -Summary ([ordered]@{ 'Total filas'=$rows.Count; Aplicados=$ok; Errores=$bad; Inválidos=$inv }) `
        -Items $results `
        -OkFields @('Aplicados') -ErrorFields @('Errores') -WarnFields @('Inválidos') `
        -LogFile $logFile
    Write-KeyValue -Key 'Informe HTML' -Value $report

} catch {
    if ($_.Exception.Message -in @('INPUT_EMPTY','INPUT_INVALID','INPUT_NOTFOUND')) {
        Show-WarningBlock -Title 'Cancelado'
    } else {
        Show-ErrorBlock -Title 'Operación fallida' -Detail $_.Exception.Message
    }
    Stop-LogSession -Persist | Out-Null
    throw
}


