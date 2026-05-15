#requires -Version 7.4
[CmdletBinding()]
param()

try { $Host.UI.RawUI.WindowTitle = 'GREX365 · Exportar miembros' } catch {}

# Helpers come from GREX365/Modules (dot-sourced by Main.ps1):
# Write-Log, Show-Header, Read-Input, Resolve-GroupByMail, Assert-RequiredServicesReady.

function Resolve-UserId {
    param([Parameter(Mandatory = $true)][string]$Email)

    $mail = $Email.Trim()
    if (-not $mail) { return $null }

    $safe = $mail.Replace("'", "''")
    try {
        $u = Get-MgUser -Filter "mail eq '$safe' or userPrincipalName eq '$safe'" -ConsistencyLevel eventual -Property Id -ErrorAction SilentlyContinue
        if ($u) { return ([string](($u | Select-Object -First 1).Id)).Trim() }
    } catch {}
    return $null
}

function Export-CleanCsv {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[object]]$Rows,
        [Parameter(Mandatory = $true)][string]$OutputCsv
    )

    $Rows |
        Select-Object @{ Name = 'Email'; Expression = { [string]$_.Email } },
                      @{ Name = 'Id';    Expression = { [string]$_.Id } } |
        Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
}

function Export-DistributionGroupMembers {
    param(
        [Parameter(Mandatory = $true)][string]$GroupIdentity,
        [Parameter(Mandatory = $true)][string]$OutputCsv
    )

    Write-Log 'Obteniendo miembros de la DL...' -Source 'Export'
    $members = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -ErrorAction Stop
    $results = New-Object System.Collections.Generic.List[object]
    $total = @($members).Count
    $i = 0

    foreach ($m in $members) {
        $i++
        $mail = if ($m.PrimarySmtpAddress) { ([string]$m.PrimarySmtpAddress).Trim() } else { $null }
        Write-Progress -Activity 'Exportando' -Status "$i / $total" -PercentComplete (($i / [math]::Max($total, 1)) * 100)
        $id = if ($mail) { Resolve-UserId -Email $mail } else { $null }
        $results.Add([PSCustomObject]@{ Email = $mail; Id = $id })
    }
    Write-Progress -Activity 'Exportando' -Completed
    Export-CleanCsv -Rows $results -OutputCsv $OutputCsv
}

function Export-UnifiedGroupMembers {
    param(
        [Parameter(Mandatory = $true)][string]$GroupIdentity,
        [Parameter(Mandatory = $true)][string]$OutputCsv
    )

    Write-Log 'Obteniendo miembros del M365 Group...' -Source 'Export'
    $members = Get-UnifiedGroupLinks -Identity $GroupIdentity -LinkType Members -ResultSize Unlimited -ErrorAction Stop
    $results = New-Object System.Collections.Generic.List[object]
    $total = @($members).Count
    $i = 0

    foreach ($m in $members) {
        $i++
        $mail = if ($m.PrimarySmtpAddress) { ([string]$m.PrimarySmtpAddress).Trim() } else { $null }
        $id   = if ($m.ExternalDirectoryObjectId) { ([string]$m.ExternalDirectoryObjectId).Trim() } else { $null }
        if (-not $id -and $mail) { $id = Resolve-UserId -Email $mail }
        Write-Progress -Activity 'Exportando' -Status "$i / $total" -PercentComplete (($i / [math]::Max($total, 1)) * 100)
        $results.Add([PSCustomObject]@{ Email = $mail; Id = $id })
    }
    Write-Progress -Activity 'Exportando' -Completed
    Export-CleanCsv -Rows $results -OutputCsv $OutputCsv
}

# --- ENTRYPOINT ---

Assert-RequiredServicesReady
Show-Header -Title 'GREX365' -Subtitle 'Exportar miembros de grupo/DL'

$session = Start-LogSession -Name 'Export-GroupMembers'

try {
    Show-Section -Title 'Origen'
    $rawSearch = Read-Input -Prompt 'Correo, nombre o alias del grupo'
    $search = Normalize-Input -Value $rawSearch
    if (-not $search) { throw 'Búsqueda vacía. Operación cancelada.' }

    $info = Resolve-GroupByMail -GroupMail $search
    if ($info.Cancelled) { Write-Log 'Selección cancelada.' -Level WARN -Source 'Export'; return }
    if (-not $info.Found) { throw ('No se encontró ningún grupo coincidente con: ' + $search) }

    Show-Section -Title 'Objeto'
    Write-KeyValue -Key 'Nombre' -Value $info.DisplayName
    Write-KeyValue -Key 'Correo' -Value $info.PrimarySmtpAddress
    Write-KeyValue -Key 'Tipo'   -Value $info.GroupType
    Write-KeyValue -Key 'ID'     -Value $info.GroupId

    $confirm = Read-Input -Prompt '¿Exportar miembros? (S/N)' -Default 'S'
    if ($confirm -notmatch '^(S|SI|Y|YES)$') { Write-Log 'Cancelado por el usuario.' -Level WARN -Source 'Export'; return }

    $outputFolder = Read-ValidatedFolder -Prompt 'Carpeta de salida CSV' -Default 'C:\Temp'

    $safeName = ($info.PrimarySmtpAddress -replace '[\\/:*?"<>|]', '_')
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outputCsv = Join-Path $outputFolder ("{0}_members_{1}.csv" -f $safeName, $stamp)

    Show-Section -Title 'Exportación'
    Write-KeyValue -Key 'Destino' -Value $outputCsv

    switch ($info.GroupType) {
        'DistributionList'         { Export-DistributionGroupMembers -GroupIdentity $info.Identity -OutputCsv $outputCsv }
        'MailEnabledSecurityGroup' { Export-DistributionGroupMembers -GroupIdentity $info.Identity -OutputCsv $outputCsv }
        'Microsoft365Group'        { Export-UnifiedGroupMembers      -GroupIdentity $info.Identity -OutputCsv $outputCsv }
        default                    { throw ('Tipo de grupo no soportado: ' + $info.GroupType) }
    }

    Write-Log ('CSV generado: ' + $outputCsv) -Level OK -Source 'Export'
    $logFile = Stop-LogSession -Persist
    Show-Section -Title 'Resumen'
    Write-KeyValue -Key 'CSV' -Value $outputCsv -ValueColor Green
    if ($logFile) { Write-KeyValue -Key 'Log' -Value $logFile }

    if ($info.GroupId) {
        Write-Host ''
        Show-AdminLink -Type $info.GroupType -Id $info.GroupId -DisplayName $info.DisplayName
    }

} catch {
    Show-ErrorBlock -Title 'Operación fallida' -Detail $_.Exception.Message
    Stop-LogSession -Persist | Out-Null
}

