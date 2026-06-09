<#
.SYNOPSIS
    M365CISCore - silnik (engine) asystenta hardeningu M365 wg CIS Microsoft 365 Foundations v6.x.
.DESCRIPTION
    Wspólny moduł dla CLI (Invoke-M365-CIS-Assistant.ps1) i GUI (Start-M365CISApp.ps1).
    Zawiera: rejestr kontrolek, opisy, połączenia, skan, wdrożenie, profile i generator raportu.
    Logowanie można przekierować do GUI przez Set-CISLogCallback.
.NOTES
    Wymaga: Microsoft.Graph, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell, MicrosoftTeams.
#>

# ---------- LOGOWANIE ----------
$script:LogCallback = $null
function Set-CISLogCallback { param([scriptblock]$Callback) $script:LogCallback = $Callback }
function Write-CISLog {
    param([string]$Message, [ValidateSet('INFO','OK','WARN','ERROR','SKIP','SCAN')]$Level='INFO')
    if ($script:LogCallback) { & $script:LogCallback $Message $Level; return }
    $color = @{ INFO='Cyan'; OK='Green'; WARN='Yellow'; ERROR='Red'; SKIP='DarkGray'; SCAN='Magenta' }[$Level]
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message) -ForegroundColor $color
}
function Confirm-CISModule {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-CISLog "Instaluje modul $Name (CurrentUser)..." WARN
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    if (-not (Get-Module -Name $Name)) {
        Write-CISLog "Importuje modul $Name..." INFO
        Import-Module -Name $Name -Force -ErrorAction Stop
    }
}

# ---------- HELPERY ----------
function New-TestResult {
    param([bool]$Compliant, [string]$Current)
    [pscustomobject]@{ Compliant = $Compliant; Current = $Current }
}
function ConvertTo-HtmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    ($Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;')
}

# ---------- KONTEKST ----------
function Reset-CISContext {
    $script:Ctx = @{
        BgId=$null; BgUpn=$null
        TenantName=$null; TenantInitialDomain=$null; TenantId=$null; TenantDisplayName=$null
        AcceptedDomains=@()
        CaState='enabledForReportingButNotEnforced'
        Connected=@{ Graph=$false; EXO=$false; SPO=$false; Teams=$false; Intune=$false }
    }
    return $script:Ctx
}
function Get-CISContext { if (-not $script:Ctx) { Reset-CISContext | Out-Null }; $script:Ctx }

# ---------- POLACZENIA ----------
function Connect-CISServices {
    [CmdletBinding()]
    param(
        [switch]$SkipEntra, [switch]$SkipExchange, [switch]$SkipSharePoint, [switch]$SkipTeams, [switch]$SkipIntune,
        [string]$TenantDomain,
        [ValidateSet('ReportOnly','Enabled')][string]$ConditionalAccessState='ReportOnly'
    )
    Reset-CISContext | Out-Null
    $script:Ctx.CaState = if ($ConditionalAccessState -eq 'Enabled') { 'enabled' } else { 'enabledForReportingButNotEnforced' }

    if (-not $SkipEntra) {
        Confirm-CISModule 'Microsoft.Graph'
        Write-CISLog 'Lacze z Microsoft Graph...'
        Connect-MgGraph -NoWelcome -Scopes @(
            'Policy.ReadWrite.ConditionalAccess','Policy.Read.All','Application.Read.All',
            'Policy.ReadWrite.Authorization','Policy.ReadWrite.AuthenticationMethod',
            'Directory.ReadWrite.All','User.Read.All','Domain.ReadWrite.All','RoleManagement.Read.Directory',
            'Organization.Read.All',
            'DeviceManagementConfiguration.ReadWrite.All','DeviceManagementServiceConfig.ReadWrite.All',
            'DeviceManagementManagedDevices.Read.All'
        ) -ErrorAction Stop
        $script:Ctx.Connected.Graph = $true
        $script:Ctx.Connected.Intune = (-not $SkipIntune)
        try {
            $org = Get-MgOrganization -ErrorAction Stop
            $script:Ctx.TenantId = $org.Id
            $script:Ctx.TenantDisplayName = $org.DisplayName
            $init = ($org.VerifiedDomains | Where-Object { $_.IsInitial }).Name
            if (-not $init) { $init = ($org.VerifiedDomains | Where-Object { $_.Name -like '*.onmicrosoft.com' } | Select-Object -First 1).Name }
            $script:Ctx.TenantInitialDomain = $init
            $script:Ctx.TenantName = ($init -split '\.')[0]
            Write-CISLog ("Tenant: {0} ({1})" -f $org.DisplayName, $init) OK
        } catch {
            Write-CISLog 'Nie udalo sie pobrac domeny z Graph; uzyje -TenantDomain jesli podany.' WARN
            if ($TenantDomain) { $script:Ctx.TenantName=($TenantDomain -split '\.')[0]; $script:Ctx.TenantInitialDomain=$TenantDomain }
        }
    }
    if (-not $SkipExchange) {
        Confirm-CISModule 'ExchangeOnlineManagement'
        Write-CISLog 'Lacze z Exchange Online...'
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $script:Ctx.Connected.EXO = $true
        try { Enable-OrganizationCustomization -ErrorAction SilentlyContinue } catch { }
        $script:Ctx.AcceptedDomains = (Get-AcceptedDomain).Name
    }
    if (-not $SkipSharePoint) {
        $tn = $script:Ctx.TenantName
        if (-not $tn -and $TenantDomain) { $tn = ($TenantDomain -split '\.')[0] }
        if (-not $tn) { Write-CISLog 'Brak nazwy tenanta - pomijam SharePoint.' WARN }
        else {
            Confirm-CISModule 'Microsoft.Online.SharePoint.PowerShell'
            Write-CISLog ("Lacze z SharePoint Admin (https://{0}-admin.sharepoint.com)..." -f $tn)
            Connect-SPOService -Url "https://$tn-admin.sharepoint.com" -ErrorAction Stop
            $script:Ctx.Connected.SPO = $true
        }
    }
    if (-not $SkipTeams) {
        Confirm-CISModule 'MicrosoftTeams'
        Write-CISLog 'Lacze z Microsoft Teams...'
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        $script:Ctx.Connected.Teams = $true
    }
    return $script:Ctx
}
function Disconnect-CISServices {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null } catch { }
}

# ---------- UZYTKOWNICY / BREAK-GLASS ----------
function Get-CISUsers {
    Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled -ErrorAction SilentlyContinue |
        Select-Object DisplayName, UserPrincipalName, AccountEnabled, Id | Sort-Object DisplayName
}
function Set-CISBreakGlass {
    param([Parameter(Mandatory)]$User)  # obiekt z .Id i .UserPrincipalName, albo UPN (string)
    if ($User -is [string]) {
        $u = Get-MgUser -Filter "userPrincipalName eq '$User'" -ErrorAction SilentlyContinue
        if (-not $u) { throw "Nie znaleziono konta: $User" }
        $script:Ctx.BgId = $u.Id; $script:Ctx.BgUpn = $u.UserPrincipalName
    } else {
        $script:Ctx.BgId = $User.Id; $script:Ctx.BgUpn = $User.UserPrincipalName
    }
    Write-CISLog ("Konto break-glass: {0}" -f $script:Ctx.BgUpn) OK
    return $script:Ctx.BgUpn
}

# ---------- SKAN / WDROZENIE ----------
function Invoke-CISScan {
    $registry = Get-CISControlRegistry
    $scan = New-Object System.Collections.Generic.List[object]
    foreach ($c in $registry) {
        if (-not $script:Ctx.Connected[$c.Service]) { continue }
        $status='Unknown'; $current='-'
        try {
            $r = & $c.Test
            $status  = if ($r.Compliant) { 'Zgodne' } else { 'NIEZGODNE' }
            $current = $r.Current
        } catch { $status='Blad skanu'; $current=$_.Exception.Message }
        $scan.Add([pscustomobject]@{
            Selected=($status -eq 'NIEZGODNE'); Id=$c.Id; Obszar=$c.Area; Kontrolka=$c.Name
            Status=$status; Poziom=("L{0}" -f $c.Level); Level=$c.Level; CIS=$c.Cis; Aktualnie=$current
        })
        Write-CISLog ("{0,-10} L{1} {2}" -f $status,$c.Level,$c.Name) (& { if($status -eq 'Zgodne'){'OK'}elseif($status -eq 'NIEZGODNE'){'WARN'}else{'ERROR'} })
    }
    return $scan
}
function Invoke-CISApply {
    param([string[]]$Ids, [switch]$WhatIf)
    $registry = Get-CISControlRegistry
    $applied = New-Object System.Collections.Generic.List[object]
    foreach ($id in $Ids) {
        $c = $registry | Where-Object Id -eq $id
        if (-not $c) { continue }
        if ($WhatIf) {
            Write-CISLog ("WHATIF: {0}" -f $c.Name) SKIP
            $applied.Add([pscustomobject]@{ Id=$c.Id; Name=$c.Name; Status='WHATIF'; Detail='' }); continue
        }
        try {
            & $c.Apply
            Write-CISLog ("WDROZONO: {0}" -f $c.Name) OK
            $applied.Add([pscustomobject]@{ Id=$c.Id; Name=$c.Name; Status='APPLIED'; Detail='' })
        } catch {
            Write-CISLog ("BLAD [{0}]: {1}" -f $c.Id,$_.Exception.Message) ERROR
            $applied.Add([pscustomobject]@{ Id=$c.Id; Name=$c.Name; Status='ERROR'; Detail=$_.Exception.Message })
        }
    }
    return $applied
}

# ---------- POLITYKI LEGACY ----------
function Remove-CISLegacyPolicies {
    $legacy = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like 'CIS - *' }
    foreach ($lp in $legacy) {
        try { Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $lp.Id -ErrorAction Stop; Write-CISLog ("Usunieto: {0}" -f $lp.DisplayName) OK }
        catch { Write-CISLog ("Blad usuwania {0}: {1}" -f $lp.DisplayName,$_.Exception.Message) ERROR }
    }
    return @($legacy).Count
}

# ---------- PROFILE / ZESTAWY KONTROLEK ----------
function Get-CISProfileSelection {
    param([Parameter(Mandatory)]$Profile, [Parameter(Mandatory)]$Scan)
    $sel = $Profile.select
    $ids = New-Object System.Collections.Generic.List[string]
    $hasIds=@($sel.ids).Count -gt 0; $hasLvl=@($sel.levels).Count -gt 0; $hasArea=@($sel.areas).Count -gt 0
    foreach ($row in $Scan) {
        $match = $false
        if ($hasIds  -and ($sel.ids    -contains $row.Id))    { $match=$true }
        if ($hasLvl  -and ($sel.levels -contains $row.Level)) { $match=$true }
        if ($hasArea -and ($sel.areas  -contains $row.Obszar)){ $match=$true }
        if (-not $hasIds -and -not $hasLvl -and -not $hasArea){ $match=$true }   # pusty profil = wszystko
        if (@($sel.excludeIds).Count -gt 0 -and ($sel.excludeIds -contains $row.Id)) { $match=$false }
        if ($match) { $ids.Add($row.Id) }
    }
    return $ids
}
function Import-CISProfile { param([Parameter(Mandatory)][string]$Path) Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json }
function Save-CISProfile {
    param([Parameter(Mandatory)][string]$Path,[string]$Name='Custom',[string]$Description='',[string[]]$Ids=@(),[int[]]$Levels=@(),[string[]]$Areas=@(),[string[]]$ExcludeIds=@())
    $obj = [pscustomobject]@{ name=$Name; description=$Description; select=[pscustomobject]@{ ids=$Ids; levels=$Levels; areas=$Areas; excludeIds=$ExcludeIds } }
    $obj | ConvertTo-Json -Depth 6 | Out-File -FilePath $Path -Encoding UTF8
    return $Path
}


# ---------- OPISY KONTROLEK ----------
$script:ControlDocs = @{
    'ENTRA-CA-LEGACY'             = 'CAP01: polityka Conditional Access blokująca starsze (legacy) protokoły uwierzytelniania (Exchange ActiveSync, inni klienci) dla wszystkich użytkowników, z wykluczeniem konta break-glass.'
    'ENTRA-CA-MFA-ALL'            = 'CAP02: polityka CA wymagająca uwierzytelniania wieloskładnikowego (MFA) dla wszystkich użytkowników i aplikacji, z wykluczeniem konta break-glass.'
    'ENTRA-CA-MFA-ADMIN'          = 'CAP03: polityka CA wymagająca MFA dla ról uprzywilejowanych (m.in. Global, Security, Exchange, SharePoint Administrator), z wykluczeniem konta break-glass.'
    'ENTRA-CONSENT'              = 'Wyłączono samodzielne udzielanie przez użytkowników zgód na aplikacje (PermissionGrantPolicies = puste) - wymagana zgoda administratora.'
    'ENTRA-PWD-NOEXPIRE'         = 'Ustawiono polityki haseł wszystkich domen jako niewygasające (zalecane przy wymuszonym MFA i modern auth).'
    'EXO-AUDIT'                  = 'Włączono Unified Audit Log (rejestrowanie zdarzeń w całym tenancie).'
    'EXO-MBXAUDIT'              = 'Włączono domyślny audyt skrzynek pocztowych (AuditDisabled = false).'
    'EXO-MODERNAUTH'           = 'Wymuszono Modern Authentication (OAuth2ClientProfileEnabled = true).'
    'EXO-AUTOFWD'              = 'Zablokowano automatyczne przekazywanie poczty na zewnątrz (AutoForwardingMode = Off oraz AutoForwardEnabled = false w domenie zdalnej).'
    'EXO-IMAPPOP'             = 'Wyłączono protokoły IMAP i POP w planach skrzynek (nowe konta) oraz w istniejących skrzynkach.'
    'EXO-OWA-STORAGE'        = 'Wyłączono zewnętrznych dostawców pamięci masowej w Outlook Web App (AdditionalStorageProvidersAvailable = false).'
    'MDO-ANTIPHISH'         = 'Utworzono politykę anti-phishing (Defender for Office 365): ochrona użytkowników i domen, mailbox intelligence, spoof intelligence, akcja Quarantine, honorowanie DMARC.'
    'MDO-SAFEATTACH'       = 'Utworzono politykę Safe Attachments z akcją Block dla domen tenanta.'
    'MDO-SAFELINKS'       = 'Utworzono politykę Safe Links (e-mail, Teams, Office) z blokadą click-through i skanowaniem URL.'
    'MDO-SAFEDOCS'       = 'Włączono Safe Docs oraz ochronę ATP dla SharePoint, OneDrive i Teams.'
    'EXO-DKIM'          = 'Włączono podpisywanie DKIM dla domen (jeśli brak rekordów CNAME w DNS - wypisano wartości do opublikowania).'
    'DNS-DMARC'        = 'Zweryfikowano/zalecono publikację rekordu DMARC (TXT _dmarc) - rekord DNS publikowany ręcznie u rejestratora.'
    'DNS-SPF'         = 'Zweryfikowano/zalecono publikację rekordu SPF (TXT) - rekord DNS publikowany ręcznie u rejestratora.'
    'INTUNE-CA-COMPLIANT'        = 'CAP04: polityka CA wymagająca urządzenia zgodnego z Intune lub Entra hybrid-joined, z wykluczeniem konta break-glass.'
    'INTUNE-NONCOMPLIANT-DEFAULT'= 'Ustawiono, że urządzenia bez przypisanej polityki zgodności są traktowane jako Niezgodne (SecureByDefault = true).'
    'INTUNE-WIN-COMPLIANCE'      = 'Utworzono bazową politykę zgodności Windows: wymóg hasła, szyfrowanie/BitLocker, Secure Boot, Code Integrity, Defender/antywirus, akcja block po 24h.'
    'SPO-SHARING'                = 'Ograniczono udostępnianie zewnętrzne w SharePoint/OneDrive do istniejących gości, domyślny link typu Direct/View, wygasanie linków anonimowych po 30 dniach, blokada re-sharingu przez gości.'
    'SPO-LEGACYAUTH'             = 'Wyłączono starsze protokoły uwierzytelniania w SharePoint (LegacyAuthProtocolsEnabled = false).'
    'TEAMS-CONSUMER'             = 'Zablokowano komunikację z kontami konsumenckimi Teams (AllowTeamsConsumer = false).'
    'TEAMS-MEETING'              = 'Wzmocniono politykę spotkań: anonimowi nie dołączają/nie startują spotkań, automatyczne wpuszczanie tylko dla użytkowników firmy (bez gości).'
}

# ---------- RAPORT HTML ----------
function New-DeploymentReport {
    param(
        [System.Collections.Generic.List[object]]$Scan,
        [System.Collections.Generic.List[object]]$Applied,
        [hashtable]$Context,
        [string]$Path,
        [switch]$ScanOnlyMode
    )
    $now      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $operator = try { (Get-MgContext).Account } catch { $env:USERNAME }
    $tenant   = if ($Context.TenantInitialDomain) { $Context.TenantInitialDomain } else { $Context.TenantName }
    $cZ = @($Scan | Where-Object Status -eq 'Zgodne').Count
    $cN = @($Scan | Where-Object Status -eq 'NIEZGODNE').Count
    $cA = @($Applied | Where-Object Status -eq 'APPLIED').Count
    $cE = @($Applied | Where-Object Status -eq 'ERROR').Count
    $cW = @($Applied | Where-Object Status -eq 'WHATIF').Count

    $appliedRows = (@($Applied) | ForEach-Object {
        $doc = $script:ControlDocs[$_.Id]; if (-not $doc) { $doc = $_.Name }
        $cls = switch ($_.Status) { 'APPLIED' {'ok'} 'WHATIF' {'warn'} default {'err'} }
        "<tr><td><code>$(ConvertTo-HtmlText $_.Id)</code></td><td>$(ConvertTo-HtmlText $_.Name)</td><td>$(ConvertTo-HtmlText $doc)</td><td class='$cls'>$(ConvertTo-HtmlText $_.Status)</td></tr>"
    }) -join "`n"
    if (-not $appliedRows) { $appliedRows = "<tr><td colspan='4'><em>Brak wdrożonych zmian (tryb skanu lub nic nie wybrano).</em></td></tr>" }

    $scanRows = (@($Scan) | ForEach-Object {
        $cls = switch ($_.Status) { 'Zgodne' {'ok'} 'NIEZGODNE' {'warn'} default {'err'} }
        "<tr><td>$(ConvertTo-HtmlText $_.Obszar)</td><td>$(ConvertTo-HtmlText $_.Poziom)</td><td>$(ConvertTo-HtmlText $_.Kontrolka)</td><td>$(ConvertTo-HtmlText $_.CIS)</td><td class='$cls'>$(ConvertTo-HtmlText $_.Status)</td><td>$(ConvertTo-HtmlText $_.Aktualnie)</td></tr>"
    }) -join "`n"

    $title = if ($ScanOnlyMode) { 'Raport zgodności CIS (skan)' } else { 'Dokumentacja powdrożeniowa CIS M365' }
    $caNote = if ($Context.CaState -eq 'enabled') { 'Enabled (egzekwowane)' } else { 'Report-only (tylko raportowanie)' }

    $html = @"
<!doctype html><html lang="pl"><head><meta charset="utf-8">
<title>$title</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;margin:32px;color:#1b1b1b;background:#fafafa}
 h1{font-size:22px;border-bottom:3px solid #0b5cab;padding-bottom:8px}
 h2{font-size:17px;margin-top:28px;color:#0b5cab}
 table{border-collapse:collapse;width:100%;margin-top:8px;background:#fff}
 th,td{border:1px solid #ddd;padding:7px 9px;font-size:13px;text-align:left;vertical-align:top}
 th{background:#0b5cab;color:#fff}
 tr:nth-child(even){background:#f4f7fb}
 code{background:#eef;padding:1px 4px;border-radius:3px}
 .ok{color:#1a7f37;font-weight:600}.warn{color:#9a6700;font-weight:600}.err{color:#b42318;font-weight:600}
 .meta{background:#fff;border:1px solid #ddd;padding:12px 16px;border-radius:6px}
 .meta td{border:none;padding:3px 12px 3px 0}
 .pill{display:inline-block;padding:2px 10px;border-radius:12px;background:#eef;margin-right:8px;font-size:12px}
 footer{margin-top:28px;color:#888;font-size:11px}
</style></head><body>
<h1>$title</h1>
<table class="meta">
 <tr><td><b>Tenant</b></td><td>$(ConvertTo-HtmlText $tenant)</td></tr>
 <tr><td><b>Tenant ID</b></td><td>$(ConvertTo-HtmlText $Context.TenantId)</td></tr>
 <tr><td><b>Wykonał</b></td><td>$(ConvertTo-HtmlText $operator)</td></tr>
 <tr><td><b>Data</b></td><td>$now</td></tr>
 <tr><td><b>Konto break-glass</b></td><td>$(ConvertTo-HtmlText $Context.BgUpn)</td></tr>
 <tr><td><b>Tryb Conditional Access</b></td><td>$caNote</td></tr>
 <tr><td><b>Benchmark</b></td><td>CIS Microsoft 365 Foundations (v6.x), poziom L1/L2</td></tr>
</table>
<p>
 <span class="pill ok">Zgodne przed: $cZ</span>
 <span class="pill warn">Niezgodne przed: $cN</span>
 <span class="pill ok">Wdrożono: $cA</span>
 <span class="pill warn">WhatIf: $cW</span>
 <span class="pill err">Błędy: $cE</span>
</p>

<h2>1. Wykonane prace (co i jak skonfigurowano)</h2>
<table><thead><tr><th>ID</th><th>Kontrolka</th><th>Opis konfiguracji</th><th>Wynik</th></tr></thead>
<tbody>
$appliedRows
</tbody></table>

<h2>2. Pełny wynik skanu tenanta</h2>
<table><thead><tr><th>Obszar</th><th>Poziom</th><th>Kontrolka</th><th>CIS</th><th>Status</th><th>Stan zastany</th></tr></thead>
<tbody>
$scanRows
</tbody></table>

<footer>Wygenerowano przez Invoke-M365-CIS-Assistant. Dokument poglądowy - zweryfikuj ustawienia w panelach administracyjnych M365.</footer>
</body></html>
"@
    $html | Out-File -FilePath $Path -Encoding UTF8
    return $Path
}

# ---------- REJESTR KONTROLEK (DODAJ NOWE TUTAJ) ----------
function Get-CISControlRegistry {
$ControlRegistry = @(

    #----- ENTRA ID / CONDITIONAL ACCESS -----
    [pscustomobject]@{
        Id='ENTRA-CA-LEGACY'; Service='Graph'; Area='Entra ID'; Cis='5.2.2'; Level=1
        Name='Conditional Access: blokada legacy authentication'
        Test={
            $p = Get-MgIdentityConditionalAccessPolicy -All |
                 Where-Object { $_.GrantControls.BuiltInControls -contains 'block' -and
                                ($_.Conditions.ClientAppTypes -contains 'exchangeActiveSync' -or
                                 $_.Conditions.ClientAppTypes -contains 'other') }
            if ($p) { New-TestResult $true ("Polityka: " + ($p[0].DisplayName)) }
            else    { New-TestResult $false 'Brak polityki blokującej legacy auth' }
        }
        Apply={
            $body=@{ displayName='CAP01: Block Legacy Authentication'; state=$script:Ctx.CaState
                conditions=@{ clientAppTypes=@('exchangeActiveSync','other')
                    applications=@{ includeApplications=@('All') }
                    users=@{ includeUsers=@('All'); excludeUsers=@($script:Ctx.BgId) } }
                grantControls=@{ operator='OR'; builtInControls=@('block') } }
            $ex=Get-MgIdentityConditionalAccessPolicy -All | Where-Object DisplayName -eq $body.displayName
            if($ex){Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $ex.Id -BodyParameter $body|Out-Null}
            else   {New-MgIdentityConditionalAccessPolicy -BodyParameter $body|Out-Null}
        }
    },
    [pscustomobject]@{
        Id='ENTRA-CA-MFA-ALL'; Service='Graph'; Area='Entra ID'; Cis='5.2.2'; Level=1
        Name='Conditional Access: MFA dla wszystkich użytkowników'
        Test={
            $p = Get-MgIdentityConditionalAccessPolicy -All |
                 Where-Object { $_.GrantControls.BuiltInControls -contains 'mfa' -and
                                $_.Conditions.Users.IncludeUsers -contains 'All' }
            if($p){New-TestResult $true ("Polityka: "+$p[0].DisplayName)} else {New-TestResult $false 'Brak MFA dla wszystkich'}
        }
        Apply={
            $body=@{ displayName='CAP02: Require MFA for All Users'; state=$script:Ctx.CaState
                conditions=@{ clientAppTypes=@('all'); applications=@{ includeApplications=@('All') }
                    users=@{ includeUsers=@('All'); excludeUsers=@($script:Ctx.BgId) } }
                grantControls=@{ operator='OR'; builtInControls=@('mfa') } }
            $ex=Get-MgIdentityConditionalAccessPolicy -All | Where-Object DisplayName -eq $body.displayName
            if($ex){Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $ex.Id -BodyParameter $body|Out-Null}
            else   {New-MgIdentityConditionalAccessPolicy -BodyParameter $body|Out-Null}
        }
    },
    [pscustomobject]@{
        Id='ENTRA-CA-MFA-ADMIN'; Service='Graph'; Area='Entra ID'; Cis='5.2.2'; Level=1
        Name='Conditional Access: MFA dla ról administracyjnych'
        Test={
            $p = Get-MgIdentityConditionalAccessPolicy -All |
                 Where-Object { $_.GrantControls.BuiltInControls -contains 'mfa' -and
                                $_.Conditions.Users.IncludeRoles.Count -gt 0 }
            if($p){New-TestResult $true ("Polityka: "+$p[0].DisplayName)} else {New-TestResult $false 'Brak osobnej polityki MFA dla adminów'}
        }
        Apply={
            $roles=@('62e90394-69f5-4237-9190-012177145e10','194ae4cb-b126-40b2-bd5b-6091b380977d',
                'f28a1f50-f6e7-4571-818b-6a12f2af6b6c','29232cdf-9323-42fd-ade2-1d097af3e4de',
                'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9','fe930be7-5e62-47db-91af-98c3a49a38b1',
                'c4e39bd9-1100-46d3-8c65-fb160da0071f','9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3',
                '7be44c8a-adaf-4e2a-84d6-ab2649e08a13','e8611ab8-c189-46e8-94e1-60213ab1f814')
            $body=@{ displayName='CAP03: Require MFA for Admin Roles'; state=$script:Ctx.CaState
                conditions=@{ clientAppTypes=@('all'); applications=@{ includeApplications=@('All') }
                    users=@{ includeRoles=$roles; excludeUsers=@($script:Ctx.BgId) } }
                grantControls=@{ operator='OR'; builtInControls=@('mfa') } }
            $ex=Get-MgIdentityConditionalAccessPolicy -All | Where-Object DisplayName -eq $body.displayName
            if($ex){Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $ex.Id -BodyParameter $body|Out-Null}
            else   {New-MgIdentityConditionalAccessPolicy -BodyParameter $body|Out-Null}
        }
    },
    [pscustomobject]@{
        Id='ENTRA-CONSENT'; Service='Graph'; Area='Entra ID'; Cis='5.1.5'; Level=1
        Name='Ogranicz zgody użytkowników na aplikacje (wymagana zgoda admina)'
        Test={
            $a=Get-MgPolicyAuthorizationPolicy
            $ok = ($a.DefaultUserRolePermissions.PermissionGrantPoliciesAssigned.Count -eq 0)
            New-TestResult $ok ("PermissionGrantPolicies: "+($a.DefaultUserRolePermissions.PermissionGrantPoliciesAssigned -join ','))
        }
        Apply={
            $a=Get-MgPolicyAuthorizationPolicy
            Update-MgPolicyAuthorizationPolicy -AuthorizationPolicyId $a.Id `
                -DefaultUserRolePermissions @{ permissionGrantPoliciesAssigned=@() }
        }
    },
    [pscustomobject]@{
        Id='ENTRA-PWD-NOEXPIRE'; Service='Graph'; Area='Entra ID'; Cis='1.x'; Level=1
        Name='Hasła niewygasające (zalecane przy MFA)'
        Test={
            $bad = Get-MgDomain | Where-Object { $_.PasswordValidityPeriodInDays -ne 2147483647 }
            if($bad){New-TestResult $false ("Domeny z wygasaniem: "+($bad.Id -join ','))} else {New-TestResult $true 'Wszystkie domeny: bez wygasania'}
        }
        Apply={ Get-MgDomain | ForEach-Object { Update-MgDomain -DomainId $_.Id -PasswordValidityPeriodInDays 2147483647 -PasswordNotificationWindowInDays 30 -ErrorAction SilentlyContinue } }
    },

    #----- EXCHANGE ONLINE / DEFENDER FOR OFFICE 365 -----
    [pscustomobject]@{
        Id='EXO-AUDIT'; Service='EXO'; Area='Exchange'; Cis='3.1.1'; Level=1
        Name='Unified Audit Log włączony'
        Test={ $v=(Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled; New-TestResult ([bool]$v) "Enabled=$v" }
        Apply={ Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true }
    },
    [pscustomobject]@{
        Id='EXO-MBXAUDIT'; Service='EXO'; Area='Exchange'; Cis='6.1.1'; Level=1
        Name='Domyślny audyt skrzynek włączony'
        Test={ $d=(Get-OrganizationConfig).AuditDisabled; New-TestResult (-not $d) "AuditDisabled=$d" }
        Apply={ Set-OrganizationConfig -AuditDisabled $false }
    },
    [pscustomobject]@{
        Id='EXO-MODERNAUTH'; Service='EXO'; Area='Exchange'; Cis='6.5'; Level=1
        Name='Wymuszony Modern Authentication (OAuth2)'
        Test={ $v=(Get-OrganizationConfig).OAuth2ClientProfileEnabled; New-TestResult ([bool]$v) "OAuth2=$v" }
        Apply={ Set-OrganizationConfig -OAuth2ClientProfileEnabled $true }
    },
    [pscustomobject]@{
        Id='EXO-AUTOFWD'; Service='EXO'; Area='Exchange'; Cis='6.2.1'; Level=1
        Name='Blokada auto-forwardingu poczty na zewnątrz'
        Test={ $m=(Get-HostedOutboundSpamFilterPolicy -Identity Default).AutoForwardingMode; New-TestResult ($m -eq 'Off') "AutoForwardingMode=$m" }
        Apply={ Set-HostedOutboundSpamFilterPolicy -Identity Default -AutoForwardingMode Off
                Set-RemoteDomain -Identity Default -AutoForwardEnabled $false }
    },
    [pscustomobject]@{
        Id='EXO-IMAPPOP'; Service='EXO'; Area='Exchange'; Cis='6.5.2'; Level=1
        Name='Wyłącz IMAP/POP (plany + istniejące skrzynki)'
        Test={
            $plans=Get-CASMailboxPlan | Where-Object { $_.ImapEnabled -or $_.PopEnabled }
            if($plans){New-TestResult $false ("Plany z IMAP/POP: "+$plans.Count)} else {New-TestResult $true 'Plany bez IMAP/POP'}
        }
        Apply={
            Get-CASMailboxPlan | Where-Object {$_.ImapEnabled -or $_.PopEnabled} | Set-CASMailboxPlan -ImapEnabled $false -PopEnabled $false
            Get-CASMailbox -ResultSize Unlimited | Where-Object {$_.ImapEnabled -or $_.PopEnabled} | Set-CASMailbox -ImapEnabled $false -PopEnabled $false
        }
    },
    [pscustomobject]@{
        Id='EXO-OWA-STORAGE'; Service='EXO'; Area='Exchange'; Cis='6.5.3'; Level=1
        Name='Wyłącz zewnętrznych dostawców pamięci w OWA'
        Test={ $bad=Get-OwaMailboxPolicy | Where-Object { $_.AdditionalStorageProvidersAvailable }
               if($bad){New-TestResult $false ("Polityki OWA z dostawcami: "+$bad.Count)} else {New-TestResult $true 'Wyłączone'} }
        Apply={ Get-OwaMailboxPolicy | ForEach-Object { Set-OwaMailboxPolicy -Identity $_.Identity -AdditionalStorageProvidersAvailable $false } }
    },
    [pscustomobject]@{
        Id='MDO-ANTIPHISH'; Service='EXO'; Area='Defender'; Cis='2.1.4'; Level=1
        Name='Polityka anti-phishing (Defender for Office 365)'
        Test={ $p=Get-AntiPhishPolicy | Where-Object { $_.Name -eq 'CIS Anti-Phishing Policy' }
               if($p){New-TestResult $true 'Istnieje CIS Anti-Phishing Policy'} else {New-TestResult $false 'Brak dedykowanej polityki CIS'} }
        Apply={
            $n='CIS Anti-Phishing Policy'; $pp=@{ EnableTargetedUserProtection=$true; EnableOrganizationDomainsProtection=$true
                EnableMailboxIntelligence=$true; EnableMailboxIntelligenceProtection=$true; EnableSpoofIntelligence=$true
                EnableFirstContactSafetyTips=$true; AuthenticationFailAction='Quarantine'; PhishThresholdLevel=3; HonorDmarcPolicy=$true }
            if(Get-AntiPhishPolicy | Where-Object Name -eq $n){ Set-AntiPhishPolicy -Identity $n @pp }
            else { New-AntiPhishPolicy -Name $n @pp|Out-Null
                   New-AntiPhishRule -Name "$n Rule" -AntiPhishPolicy $n -RecipientDomainIs $script:Ctx.AcceptedDomains -Priority 0 -ErrorAction SilentlyContinue|Out-Null }
        }
    },
    [pscustomobject]@{
        Id='MDO-SAFEATTACH'; Service='EXO'; Area='Defender'; Cis='2.1.1'; Level=1
        Name='Safe Attachments (Block)'
        Test={ $p=Get-SafeAttachmentPolicy | Where-Object { $_.Enable }; if($p){New-TestResult $true ("Aktywne: "+$p.Count)} else {New-TestResult $false 'Brak aktywnej polityki Safe Attachments'} }
        Apply={
            $n='CIS Safe Attachments Policy'
            if(Get-SafeAttachmentPolicy | Where-Object Name -eq $n){ Set-SafeAttachmentPolicy -Identity $n -Enable $true -Action Block }
            else { New-SafeAttachmentPolicy -Name $n -Enable $true -Action Block|Out-Null
                   New-SafeAttachmentRule -Name "$n Rule" -SafeAttachmentPolicy $n -RecipientDomainIs $script:Ctx.AcceptedDomains -Priority 0 -ErrorAction SilentlyContinue|Out-Null }
        }
    },
    [pscustomobject]@{
        Id='MDO-SAFELINKS'; Service='EXO'; Area='Defender'; Cis='2.1.3'; Level=1
        Name='Safe Links (mail/Teams/Office)'
        Test={ $p=Get-SafeLinksPolicy | Where-Object { $_.EnableSafeLinksForEmail }; if($p){New-TestResult $true ("Aktywne: "+$p.Count)} else {New-TestResult $false 'Brak aktywnej polityki Safe Links'} }
        Apply={
            $n='CIS Safe Links Policy'; $pp=@{ EnableSafeLinksForEmail=$true; EnableSafeLinksForTeams=$true
                EnableSafeLinksForOffice=$true; TrackClicks=$true; AllowClickThrough=$false; ScanUrls=$true; DeliverMessageAfterScan=$true }
            if(Get-SafeLinksPolicy | Where-Object Name -eq $n){ Set-SafeLinksPolicy -Identity $n @pp }
            else { New-SafeLinksPolicy -Name $n @pp|Out-Null
                   New-SafeLinksRule -Name "$n Rule" -SafeLinksPolicy $n -RecipientDomainIs $script:Ctx.AcceptedDomains -Priority 0 -ErrorAction SilentlyContinue|Out-Null }
        }
    },
    [pscustomobject]@{
        Id='MDO-SAFEDOCS'; Service='EXO'; Area='Defender'; Cis='2.1.x'; Level=1
        Name='Safe Docs + ochrona SPO/OneDrive/Teams'
        Test={ $a=Get-AtpPolicyForO365; New-TestResult ([bool]$a.EnableSafeDocs -and [bool]$a.EnableATPForSPOTeamsODB) ("SafeDocs="+$a.EnableSafeDocs+"; SPO/ODB/Teams="+$a.EnableATPForSPOTeamsODB) }
        Apply={ Set-AtpPolicyForO365 -EnableSafeDocs $true -EnableATPForSPOTeamsODB $true -AllowSafeDocsOpen $false }
    },

    #----- SHAREPOINT ONLINE / ONEDRIVE -----
    [pscustomobject]@{
        Id='SPO-SHARING'; Service='SPO'; Area='SharePoint'; Cis='7.2'; Level=1
        Name='Ogranicz udostępnianie zewnętrzne + wygasanie linków'
        Test={ $t=Get-SPOTenant; $ok = $t.SharingCapability -in @('Disabled','ExistingExternalUserSharingOnly')
               New-TestResult $ok ("SharingCapability="+$t.SharingCapability) }
        Apply={ Set-SPOTenant -SharingCapability ExistingExternalUserSharingOnly -DefaultSharingLinkType Direct `
                    -DefaultLinkPermission View -RequireAnonymousLinksExpireInDays 30 -PreventExternalUsersFromResharing $true }
    },
    [pscustomobject]@{
        Id='SPO-LEGACYAUTH'; Service='SPO'; Area='SharePoint'; Cis='7.2.x'; Level=1
        Name='Wyłącz legacy auth w SharePoint'
        Test={ $t=Get-SPOTenant; New-TestResult (-not $t.LegacyAuthProtocolsEnabled) ("LegacyAuth="+$t.LegacyAuthProtocolsEnabled) }
        Apply={ Set-SPOTenant -LegacyAuthProtocolsEnabled $false }
    },

    #----- MICROSOFT TEAMS -----
    [pscustomobject]@{
        Id='TEAMS-CONSUMER'; Service='Teams'; Area='Teams'; Cis='8.2'; Level=1
        Name='Zablokuj czat z kontami konsumenckimi Teams'
        Test={ $f=Get-CsTenantFederationConfiguration; New-TestResult (-not $f.AllowTeamsConsumer) ("AllowTeamsConsumer="+$f.AllowTeamsConsumer) }
        Apply={ Set-CsTenantFederationConfiguration -AllowTeamsConsumer $false }
    },
    [pscustomobject]@{
        Id='TEAMS-MEETING'; Service='Teams'; Area='Teams'; Cis='8.1'; Level=1
        Name='Wzmocnij lobby/anonimowych w spotkaniach'
        Test={ $m=Get-CsTeamsMeetingPolicy -Identity Global; New-TestResult (-not $m.AllowAnonymousUsersToJoinMeeting) ("AnonJoin="+$m.AllowAnonymousUsersToJoinMeeting) }
        Apply={ Set-CsTeamsMeetingPolicy -Identity Global -AllowAnonymousUsersToJoinMeeting $false `
                    -AllowAnonymousUsersToStartMeeting $false -AutoAdmittedUsers EveryoneInCompanyExcludingGuests }
    },

    #----- UWIERZYTELNIANIE POCZTY DOMEN (DKIM / DMARC / SPF) -----
    [pscustomobject]@{
        Id='EXO-DKIM'; Service='EXO'; Area='Email-Auth'; Cis='2.1.x'; Level=1
        Name='DKIM włączony dla domen'
        Test={
            $bad=@()
            foreach($d in $script:Ctx.AcceptedDomains){
                $c=Get-DkimSigningConfig -Identity $d -ErrorAction SilentlyContinue
                if(-not $c -or -not $c.Enabled){ $bad+=$d }
            }
            if($bad.Count){ New-TestResult $false ("Bez DKIM: "+($bad -join ',')) } else { New-TestResult $true 'DKIM aktywny na wszystkich domenach' }
        }
        Apply={
            foreach($d in $script:Ctx.AcceptedDomains){
                $c=Get-DkimSigningConfig -Identity $d -ErrorAction SilentlyContinue
                if(-not $c){
                    # Tworzy konfigurację (generuje selektory/CNAME). Włączenie powiedzie się tylko gdy CNAME są w DNS.
                    try { New-DkimSigningConfig -DomainName $d -Enabled $true -ErrorAction Stop | Out-Null }
                    catch {
                        New-DkimSigningConfig -DomainName $d -Enabled $false -ErrorAction SilentlyContinue | Out-Null
                        $c2=Get-DkimSigningConfig -Identity $d -ErrorAction SilentlyContinue
                        Write-Host ("  [DKIM] {0}: opublikuj CNAME w DNS, potem uruchom ponownie:" -f $d) -ForegroundColor Yellow
                        Write-Host ("        selector1._domainkey -> {0}" -f $c2.Selector1CNAME) -ForegroundColor Yellow
                        Write-Host ("        selector2._domainkey -> {0}" -f $c2.Selector2CNAME) -ForegroundColor Yellow
                    }
                } elseif(-not $c.Enabled){
                    try { Set-DkimSigningConfig -Identity $d -Enabled $true -ErrorAction Stop }
                    catch {
                        Write-Host ("  [DKIM] {0}: brak rekordów CNAME w DNS - opublikuj i uruchom ponownie." -f $d) -ForegroundColor Yellow
                        Write-Host ("        selector1._domainkey -> {0}" -f $c.Selector1CNAME) -ForegroundColor Yellow
                        Write-Host ("        selector2._domainkey -> {0}" -f $c.Selector2CNAME) -ForegroundColor Yellow
                    }
                }
            }
        }
    },
    [pscustomobject]@{
        Id='DNS-DMARC'; Service='EXO'; Area='Email-Auth'; Cis='2.1.x'; Level=1
        Name='DMARC opublikowany [rekord DNS - ręcznie]'
        Test={
            if(-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)){ return New-TestResult $false 'Brak Resolve-DnsName (uruchom na Windows)' }
            $bad=@()
            foreach($d in $script:Ctx.AcceptedDomains){
                $txt=(Resolve-DnsName -Name "_dmarc.$d" -Type TXT -ErrorAction SilentlyContinue).Strings
                if(-not ($txt -match 'v=DMARC1')){ $bad+=$d }
            }
            if($bad.Count){ New-TestResult $false ("Bez DMARC: "+($bad -join ',')) } else { New-TestResult $true 'DMARC obecny na wszystkich domenach' }
        }
        Apply={
            Write-Host "  [DMARC] DNS jest poza M365 - opublikuj u rejestratora rekord TXT dla każdej domeny:" -ForegroundColor Yellow
            foreach($d in $script:Ctx.AcceptedDomains){
                Write-Host ("        _dmarc.{0}  TXT  ""v=DMARC1; p=quarantine; rua=mailto:dmarc@{0}; ruf=mailto:dmarc@{0}; fo=1""" -f $d) -ForegroundColor Yellow
            }
        }
    },
    [pscustomobject]@{
        Id='DNS-SPF'; Service='EXO'; Area='Email-Auth'; Cis='2.1.x'; Level=1
        Name='SPF opublikowany [rekord DNS - ręcznie]'
        Test={
            if(-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)){ return New-TestResult $false 'Brak Resolve-DnsName (uruchom na Windows)' }
            $bad=@()
            foreach($d in $script:Ctx.AcceptedDomains){
                $txt=(Resolve-DnsName -Name $d -Type TXT -ErrorAction SilentlyContinue).Strings
                if(-not ($txt -match 'v=spf1')){ $bad+=$d }
            }
            if($bad.Count){ New-TestResult $false ("Bez SPF: "+($bad -join ',')) } else { New-TestResult $true 'SPF obecny na wszystkich domenach' }
        }
        Apply={
            Write-Host "  [SPF] DNS jest poza M365 - opublikuj rekord TXT dla każdej domeny:" -ForegroundColor Yellow
            foreach($d in $script:Ctx.AcceptedDomains){
                Write-Host ("        {0}  TXT  ""v=spf1 include:spf.protection.outlook.com -all""" -f $d) -ForegroundColor Yellow
            }
        }
    },

    #----- MICROSOFT INTUNE (device compliance) -----
    [pscustomobject]@{
        Id='INTUNE-CA-COMPLIANT'; Service='Intune'; Area='Intune'; Cis='5.2.x'; Level=2
        Name='Conditional Access: wymagaj urządzenia zgodnego / hybrid-joined'
        Test={
            $p=Get-MgIdentityConditionalAccessPolicy -All |
               Where-Object { $_.GrantControls.BuiltInControls -contains 'compliantDevice' }
            if($p){ New-TestResult $true ("Polityka: "+$p[0].DisplayName) } else { New-TestResult $false 'Brak CA na zgodne urządzenie' }
        }
        Apply={
            $body=@{ displayName='CAP04: Require Compliant or Hybrid Joined Device'; state=$script:Ctx.CaState
                conditions=@{ clientAppTypes=@('all'); applications=@{ includeApplications=@('All') }
                    users=@{ includeUsers=@('All'); excludeUsers=@($script:Ctx.BgId) } }
                grantControls=@{ operator='OR'; builtInControls=@('compliantDevice','domainJoinedDevice') } }
            $ex=Get-MgIdentityConditionalAccessPolicy -All | Where-Object DisplayName -eq $body.displayName
            if($ex){Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $ex.Id -BodyParameter $body|Out-Null}
            else   {New-MgIdentityConditionalAccessPolicy -BodyParameter $body|Out-Null}
        }
    },
    [pscustomobject]@{
        Id='INTUNE-NONCOMPLIANT-DEFAULT'; Service='Intune'; Area='Intune'; Cis='6.x'; Level=1
        Name='Urządzenia bez polityki zgodności = oznaczone jako Niezgodne'
        Test={
            $s=(Get-MgDeviceManagement -Property settings -ErrorAction SilentlyContinue).Settings
            New-TestResult ([bool]$s.SecureByDefault) ("SecureByDefault="+$s.SecureByDefault)
        }
        Apply={ Update-MgDeviceManagement -BodyParameter @{ settings=@{ secureByDefault=$true } } }
    },
    [pscustomobject]@{
        Id='INTUNE-WIN-COMPLIANCE'; Service='Intune'; Area='Intune'; Cis='6.x'; Level=1
        Name='Bazowa polityka zgodności Windows (BitLocker, Secure Boot, AV, hasło)'
        Test={
            $pol=Get-MgDeviceManagementDeviceCompliancePolicy -All -ErrorAction SilentlyContinue
            if($pol -and $pol.Count -gt 0){ New-TestResult $true ("Polityk zgodności: "+$pol.Count) } else { New-TestResult $false 'Brak polityk zgodności urządzeń' }
        }
        Apply={
            $body=@{
                '@odata.type'='#microsoft.graph.windows10CompliancePolicy'
                displayName='CIS - Windows Baseline Compliance'
                passwordRequired=$true; passwordMinimumLength=8; passwordRequiredType='alphanumeric'
                storageRequireEncryption=$true; bitLockerEnabled=$true; secureBootEnabled=$true
                codeIntegrityEnabled=$true; defenderEnabled=$true; rtpEnabled=$true
                antivirusRequired=$true; antiSpywareRequired=$true
                scheduledActionsForRule=@(
                    @{ ruleName='PasswordRequired'
                       scheduledActionConfigurations=@( @{ actionType='block'; gracePeriodHours=24 } ) }
                )
            }
            New-MgDeviceManagementDeviceCompliancePolicy -BodyParameter $body | Out-Null
        }
    }
)
    return $ControlRegistry
}


# ---------- EKSPORT ----------
function Get-CISControlDocs { $script:ControlDocs }

Export-ModuleMember -Function `
    Set-CISLogCallback, Write-CISLog, Confirm-CISModule, New-TestResult, ConvertTo-HtmlText, `
    Reset-CISContext, Get-CISContext, Connect-CISServices, Disconnect-CISServices, `
    Get-CISUsers, Set-CISBreakGlass, Invoke-CISScan, Invoke-CISApply, Remove-CISLegacyPolicies, `
    Get-CISProfileSelection, Import-CISProfile, Save-CISProfile, `
    New-DeploymentReport, Get-CISControlRegistry, Get-CISControlDocs
