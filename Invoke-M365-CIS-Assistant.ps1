#Requires -Version 5.1
<#
.SYNOPSIS
    CLI asystenta hardeningu M365 wg CIS v6.x. Cienka nakladka na modul M365CISCore.psm1.
.DESCRIPTION
    Loguje do tenanta (domena auto z Graph), wymusza wybor konta break-glass, skanuje tenant,
    pozwala wybrac kontrolki (okno Out-GridView / konsola / profil) i wdraza je,
    a na koncu generuje dokumentacje powdrozeniowa HTML.
.PARAMETER BreakGlassAccountUPN
    Opcjonalny. Jesli pusty - wybor z listy po zalogowaniu (wymuszony).
.PARAMETER TenantDomain
    Opcjonalny override; domyslnie domena pobierana automatycznie z Graph.
.PARAMETER ConditionalAccessState
    'ReportOnly' (domyslnie) lub 'Enabled'.
.PARAMETER Profile
    Sciezka do profilu JSON (zestaw kontrolek). Uzywany do wstepnego wyboru / trybu -Unattended.
.PARAMETER Unattended
    Bez interakcji: wdraza wybor wynikajacy z profilu (lub wszystkie niezgodne, gdy brak profilu).
.PARAMETER ScanOnly
    Tylko skan + raport, bez wdrozenia.
.PARAMETER SkipEntra/SkipExchange/SkipSharePoint/SkipTeams/SkipIntune
    Pomija dana usluge.
.PARAMETER RemoveLegacyPolicies
    Usuwa polityki CA o starych nazwach (CIS - ...).
.EXAMPLE
    .\Invoke-M365-CIS-Assistant.ps1 -WhatIf
.EXAMPLE
    .\Invoke-M365-CIS-Assistant.ps1 -Profile .\profiles\Baseline-L1.json -Unattended
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [string]$BreakGlassAccountUPN,
    [string]$TenantDomain,
    [ValidateSet('ReportOnly','Enabled')][string]$ConditionalAccessState='ReportOnly',
    [string]$Profile,
    [switch]$Unattended,
    [switch]$ScanOnly,
    [switch]$SkipEntra,[switch]$SkipExchange,[switch]$SkipSharePoint,[switch]$SkipTeams,[switch]$SkipIntune,
    [switch]$RemoveLegacyPolicies,
    [string]$LogPath = "$PSScriptRoot\M365-CIS-Assistant-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
)

Import-Module (Join-Path $PSScriptRoot 'M365CISCore.psm1') -Force -ErrorAction Stop
try { Start-Transcript -Path $LogPath -Force | Out-Null } catch { }
Write-CISLog "=== M365 CIS Assistant (CLI) | start $(Get-Date) ===" INFO
if ($WhatIfPreference) { Write-CISLog "Tryb -WhatIf: wdrozenie nic nie zapisze." WARN }

# 1) Polaczenia + auto-domena
Connect-CISServices -SkipEntra:$SkipEntra -SkipExchange:$SkipExchange -SkipSharePoint:$SkipSharePoint `
    -SkipTeams:$SkipTeams -SkipIntune:$SkipIntune -TenantDomain $TenantDomain `
    -ConditionalAccessState $ConditionalAccessState | Out-Null
$ctx = Get-CISContext

# 2) Stare polityki (opcjonalne sprzatanie)
if ($ctx.Connected.Graph) {
    $legacy = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like 'CIS - *' }
    if ($legacy) {
        Write-CISLog ("Wykryto {0} starych polityk CA (CIS - ...)." -f @($legacy).Count) WARN
        if ($RemoveLegacyPolicies) { if ($PSCmdlet.ShouldProcess('Stare polityki CA','Usun')) { Remove-CISLegacyPolicies | Out-Null } }
        else { Write-CISLog "Aby usunac: uruchom z -RemoveLegacyPolicies." INFO }
    }
}

# 3) WYMUSZONY break-glass (gdy Entra w grze)
if ($ctx.Connected.Graph) {
    $bgUser = $null
    if ($BreakGlassAccountUPN) {
        $bgUser = Get-MgUser -Filter "userPrincipalName eq '$BreakGlassAccountUPN'" -ErrorAction SilentlyContinue
        if (-not $bgUser) { Write-CISLog "Podane konto break-glass nie istnieje - wybierz z listy." WARN }
    }
    while (-not $bgUser) {
        $users = Get-CISUsers
        if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
            $pick = $users | Out-GridView -Title 'Wybierz konto BREAK-GLASS i kliknij OK' -OutputMode Single
        } else {
            for ($i=0;$i -lt $users.Count;$i++){ Write-Host ("{0,4}) {1} <{2}>" -f $i,$users[$i].DisplayName,$users[$i].UserPrincipalName) }
            $n = Read-Host 'Numer konta break-glass'; if ($n -match '^\d+$') { $pick=$users[[int]$n] }
        }
        if ($pick) { $bgUser = Get-MgUser -UserId $pick.Id -ErrorAction SilentlyContinue }
        if (-not $bgUser) {
            if ((Read-Host 'Nie wybrano. Sprobowac ponownie? (T/N)') -notmatch '^[TtYy]') {
                Write-CISLog "Bez break-glass nie mozna bezpiecznie tworzyc CA. Przerywam." ERROR
                Disconnect-CISServices; try { Stop-Transcript|Out-Null } catch {}; return
            }
        }
    }
    Set-CISBreakGlass -User $bgUser | Out-Null
}

# 4) Skan
Write-CISLog "--- SKAN wg CIS ---" SCAN
$scan = Invoke-CISScan
try { $scan | Select-Object Id,Obszar,Kontrolka,Status,Poziom,CIS,Aktualnie |
      Export-Csv ($LogPath -replace '\.log$','-scan.csv') -NoTypeInformation -Encoding UTF8 } catch { }

# 5) ScanOnly => raport i koniec
if ($ScanOnly) {
    $scan | Format-Table Status,Poziom,Obszar,Kontrolka -AutoSize | Out-String | Write-Host
    $html = $LogPath -replace '\.log$','-report.html'
    try { New-DeploymentReport -Scan $scan -Applied (New-Object System.Collections.Generic.List[object]) -Context $ctx -Path $html -ScanOnlyMode | Out-Null
          Write-CISLog "Dokumentacja (skan): $html" OK } catch { Write-CISLog $_.Exception.Message WARN }
    Disconnect-CISServices; try { Stop-Transcript|Out-Null } catch {}; return
}

# 6) Wybor kontrolek
$defaultIds = @($scan | Where-Object Selected).Id
if ($Profile) {
    try {
        $p = Import-CISProfile -Path $Profile
        $defaultIds = @(Get-CISProfileSelection -CisProfile $p -Scan $scan)
        Write-CISLog ("Profil '{0}': {1} kontrolek." -f $p.name, $defaultIds.Count) INFO
    } catch { Write-CISLog ("Blad profilu: {0}" -f $_.Exception.Message) ERROR }
}
if ($Unattended) {
    $selectedIds = $defaultIds
} elseif (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
    Write-CISLog "Zaznacz kontrolki do wdrozenia (Ctrl/Shift) i kliknij OK." INFO
    $picked = $scan | Select-Object Status,Poziom,Obszar,Kontrolka,CIS,Aktualnie,Id |
              Out-GridView -Title 'CIS M365 - zaznacz co wdrozyc' -PassThru
    $selectedIds = @($picked.Id)
} else {
    for ($i=0;$i -lt $scan.Count;$i++){ Write-Host ("{0,3}) [{1}] [{2}] {3}" -f $i,$scan[$i].Status,$scan[$i].Poziom,$scan[$i].Kontrolka) }
    $inp = Read-Host 'Numery (Enter = niezgodne)'
    $selectedIds = if ([string]::IsNullOrWhiteSpace($inp)) { $defaultIds } else { @($inp -split '[, ]+' | Where-Object {$_ -match '^\d+$'} | ForEach-Object { $scan[[int]$_].Id }) }
}
if (-not $selectedIds -or $selectedIds.Count -eq 0) {
    Write-CISLog "Nic nie wybrano - koncze." INFO; Disconnect-CISServices; try { Stop-Transcript|Out-Null } catch {}; return
}
Write-CISLog ("Do wdrozenia: {0}" -f ($selectedIds -join ', ')) INFO

# 7) Wdrozenie
$applied = Invoke-CISApply -Ids $selectedIds -WhatIf:$WhatIfPreference
$applied | Format-Table Status,Id,Name -AutoSize | Out-String | Write-Host
try { $applied | Export-Csv ($LogPath -replace '\.log$','-applied.csv') -NoTypeInformation -Encoding UTF8 } catch { }

# 8) Dokumentacja powdrozeniowa
$html = $LogPath -replace '\.log$','-report.html'
try { New-DeploymentReport -Scan $scan -Applied $applied -Context $ctx -Path $html | Out-Null
      Write-CISLog "Dokumentacja powdrozeniowa: $html" OK
      if (-not $WhatIfPreference) { try { Invoke-Item $html } catch { } } } catch { Write-CISLog $_.Exception.Message WARN }

Disconnect-CISServices
Write-CISLog "=== Zakonczono $(Get-Date) ===" INFO
try { Stop-Transcript | Out-Null } catch { }
