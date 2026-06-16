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
$script:LogCallback        = $null
$script:DeviceCodeCallback = $null   # scriptblock: param($Url, $Code) -> zwraca scriptblock do zamkniecia okna
$script:ExoModFile         = $null   # pelna sciezka do zaladowanego ExchangeOnlineManagement.psd1
$script:MsalApp            = $null   # PublicClientApplication - cache miedzy polaczeniami
function Set-CISLogCallback        { param([scriptblock]$Callback) $script:LogCallback      = $Callback }
function Set-CISDeviceCodeCallback { param([scriptblock]$Callback) $script:DeviceCodeCallback = $Callback }
function Write-CISLog {
    param([string]$Message, [ValidateSet('INFO','OK','WARN','ERROR','SKIP','SCAN')]$Level='INFO')
    if ($script:LogCallback) { & $script:LogCallback $Message $Level; return }
    $color = @{ INFO='Cyan'; OK='Green'; WARN='Yellow'; ERROR='Red'; SKIP='DarkGray'; SCAN='Magenta' }[$Level]
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message) -ForegroundColor $color
}
function Confirm-CISModule {
    param([string]$Name, [switch]$OnlyInstall, [version]$MinVersion)
    $installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    $needsInstall = -not $installed -or ($MinVersion -and [version]$installed.Version -lt $MinVersion)
    if ($needsInstall) {
        $label = if ($MinVersion) { "$Name (min v$MinVersion)" } else { $Name }
        Write-CISLog "Instaluje $label..." WARN
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-CISLog "Zainstalowano $Name." OK
    }
    if (-not $OnlyInstall) {
        Write-CISLog "Importuje modul $Name..." INFO
        Import-Module -Name $Name -Force -ErrorAction Stop
        Write-CISLog "Zaladowano $Name." OK
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
        Connected=@{ Graph=$false; EXO=$false; SPO=$false; Teams=$false; Intune=$false; Purview=$false; PowerPlatform=$false }
    }
    return $script:Ctx
}
function Get-CISContext { if (-not $script:Ctx) { Reset-CISContext | Out-Null }; $script:Ctx }
function Set-CISContext { param([hashtable]$Ctx) $script:Ctx = $Ctx }

# Uruchamia polecenie uwierzytelnienia w tle (runspace) z przechwytem Write-Host/Write-Warning,
# wydobywa kod urzadzenia i pokazuje go przez DeviceCodeCallback lub log.
# $ConnectInvoke - string z poleceniem do Invoke-Expression w runspace.
function Invoke-CISDeviceConnect {
    param([string]$ServiceName, [string]$ConnectInvoke)
    Write-CISLog "$ServiceName - logowanie kodem urzadzenia..." WARN

    $authSync = [hashtable]::Synchronized(@{
        Done=$false; Error=$null; Code=$null
        Url='https://microsoft.com/devicelogin'
        CodeShown=$false; DismissCallback=$null; FallbackWin=$null
    })

    $authRs = [runspacefactory]::CreateRunspace()
    $authRs.Open()
    $authRs.SessionStateProxy.SetVariable('authSync',       $authSync)
    $authRs.SessionStateProxy.SetVariable('_connectInvoke', $ConnectInvoke)
    $authRs.SessionStateProxy.SetVariable('_psModulePath',  $env:PSModulePath)

    $authPs = [powershell]::Create()
    $authPs.Runspace = $authRs
    [void]$authPs.AddScript({
        $env:PSModulePath = $_psModulePath
        function Write-Host {
            param([object]$Object,[switch]$NoNewline,
                  [System.ConsoleColor]$ForegroundColor,[System.ConsoleColor]$BackgroundColor)
            $msg = [string]$Object
            if ($msg -match '(https://[^\s]+devicelogin[^\s]*)') { $authSync.Url  = $Matches[1].TrimEnd('.') }
            if ($msg -match '\bcode\s+([A-Z0-9]{7,12})\b')       { $authSync.Code = $Matches[1] }
        }
        function Write-Warning {
            param([object]$Message)
            $msg = [string]$Message
            if ($msg -match '(https://[^\s]+devicelogin[^\s]*)') { $authSync.Url  = $Matches[1].TrimEnd('.') }
            if ($msg -match '\bcode\s+([A-Z0-9]{7,12})\b')       { $authSync.Code = $Matches[1] }
        }
        try   { Invoke-Expression $_connectInvoke }
        catch { $authSync.Error = $_.Exception.Message }
        finally { $authSync.Done = $true }
    })
    $authHandle = $authPs.BeginInvoke()

    # Pomocnik skanujacy Information i Warning streamy w poszukiwaniu kodu urzadzenia
    $scanSources = {
        foreach ($item in @($authPs.Streams.Information)) {
            $msg = [string]$item.MessageData
            if ($msg -match '(https://[^\s]+devicelogin[^\s]*)') { $authSync.Url  = $Matches[1].TrimEnd('.') }
            if ($msg -match '\bcode\s+([A-Z0-9]{7,12})\b')       { $authSync.Code = $Matches[1] }
        }
        foreach ($item in @($authPs.Streams.Warning)) {
            $msg = [string]$item.Message
            if ($msg -match '(https://[^\s]+devicelogin[^\s]*)') { $authSync.Url  = $Matches[1].TrimEnd('.') }
            if ($msg -match '\bcode\s+([A-Z0-9]{7,12})\b')       { $authSync.Code = $Matches[1] }
        }
    }

    if ([System.Windows.Application]::Current) {
        $dcFrame   = [System.Windows.Threading.DispatcherFrame]::new()
        $startTime = [DateTime]::UtcNow
        $deadline  = $startTime.AddMinutes(15)
        $_svcName  = $ServiceName
        $_dcCb     = $script:DeviceCodeCallback

        $dcTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $dcTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $dcTimer.Add_Tick({
            if (-not $authSync.Code) { & $scanSources }

            # Kod znaleziony - pokaz okno z kodem
            if ($authSync.Code -and -not $authSync.CodeShown) {
                $authSync.CodeShown = $true
                if ($authSync.FallbackWin) { $authSync.FallbackWin.Close(); $authSync.FallbackWin = $null }
                $codeDisp = ($authSync.Code.ToCharArray() -join ' ')
                if ($_dcCb) {
                    $authSync.DismissCallback = & $_dcCb $authSync.Url $codeDisp
                } else {
                    Write-CISLog "=== $_svcName ===" INFO
                    Write-CISLog "Otworz: $($authSync.Url)  Kod: $codeDisp" INFO
                }
            }

            # Po 5s bez kodu - pokaz okno z URL i instrukcja (zamknie sie samo po auth)
            if (-not $authSync.CodeShown -and -not $authSync.Done -and
                ([DateTime]::UtcNow - $startTime).TotalSeconds -gt 5 -and
                -not $authSync.FallbackWin) {
                $fbXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Logowanie: $_svcName" Width="480" Height="210"
        WindowStartupLocation="CenterScreen" Topmost="True" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="13" Background="#f4f6f9" ResizeMode="NoResize">
  <Grid Margin="24,18,24,16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="12"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="16"/>
      <RowDefinition Height="6"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" TextWrapping="Wrap" Foreground="#333">
      Otworz w przegladarce i wpisz kod z okna <Bold>konsoli PowerShell</Bold> (za tym oknem):
    </TextBlock>
    <Border Grid.Row="2" Background="White" CornerRadius="6" Padding="10,8"
            BorderBrush="#b3d0ee" BorderThickness="1">
      <DockPanel>
        <Button x:Name="btnFbOpen" DockPanel.Dock="Right" Content="Otworz" Width="72" Height="28"
                Margin="8,0,0,0" Background="#0b5cab" Foreground="White" BorderThickness="0" FontWeight="Bold"/>
        <TextBlock FontFamily="Consolas" FontSize="12" Foreground="#0b5cab" VerticalAlignment="Center"
                   Text="https://microsoft.com/devicelogin"/>
      </DockPanel>
    </Border>
    <ProgressBar Grid.Row="4" Height="6" IsIndeterminate="True" Foreground="#0b5cab" Background="#e0e0e0"/>
    <TextBlock Grid.Row="5" Foreground="#999" FontSize="11"
               Text="Okno zamknie sie automatycznie po zalogowaniu."/>
  </Grid>
</Window>
"@
                try {
                    [xml]$fbXml2 = $fbXaml
                    $fbWin = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $fbXml2))
                    $fbWin.FindName('btnFbOpen').Add_Click({ Start-Process 'https://microsoft.com/devicelogin' })
                    $authSync.FallbackWin = $fbWin
                    $fbWin.Show()
                } catch { Write-CISLog "Blad okna fallback: $($_.Exception.Message)" WARN }
            }

            if ($authSync.Done -or [DateTime]::UtcNow -gt $deadline) {
                $dcTimer.Stop()
                if ($authSync.FallbackWin) { $authSync.FallbackWin.Close(); $authSync.FallbackWin = $null }
                $dcFrame.Continue = $false
            }
        }.GetNewClosure())
        $dcTimer.Start()
        [System.Windows.Threading.Dispatcher]::PushFrame($dcFrame)
    } else {
        # Tryb CLI - proste czekanie w petli
        $wLimit = 0
        while (-not $authSync.Done -and $wLimit -lt 900) {
            if (-not $authSync.Code) { & $scanSources }
            if ($authSync.Code -and -not $authSync.CodeShown) {
                $authSync.CodeShown = $true
                Write-CISLog "Otworz: $($authSync.Url)  Kod: $(($authSync.Code.ToCharArray() -join ' '))" INFO
            }
            [System.Threading.Thread]::Sleep(1000); $wLimit++
        }
    }

    if ($authSync.DismissCallback) { & $authSync.DismissCallback }
    try { $authPs.EndInvoke($authHandle) } catch {}
    $authPs.Dispose(); $authRs.Close()
    if ($authSync.Error) { throw $authSync.Error }
}

# Logowanie do Graph przez MSAL.NET bezposrednio - omija Connect-MgGraph auth (problemy z WAM/runspace).
# Uzywamy Microsoft Graph Command Line Tools (publiczny klient MS, delegated, kazdy tenant).
# AcquireTokenInteractive otwiera przegladarke systemowa (nie WAM, nie embedded browser).
# ---------- MSAL HELPERS ----------

function Initialize-CISMsal {
    # Upewnij sie ze MSAL.NET jest zaladowany (moze byc lazy-loaded przez Graph modul)
    $msalAsm = [System.AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' } |
        Select-Object -First 1
    if (-not $msalAsm) {
        $msalDll = Get-Module -Name 'Microsoft.Graph.*' |
            Select-Object -ExpandProperty ModuleBase -ErrorAction SilentlyContinue |
            ForEach-Object { Get-ChildItem $_ -Recurse -Filter 'Microsoft.Identity.Client.dll' -ErrorAction SilentlyContinue } |
            Sort-Object { [version]$_.VersionInfo.FileVersion } -Descending |
            Select-Object -First 1
        if (-not $msalDll) {
            $msalDll = ($env:PSModulePath -split ';') |
                ForEach-Object { Get-ChildItem $_ -Recurse -Filter 'Microsoft.Identity.Client.dll' -ErrorAction SilentlyContinue } |
                Sort-Object { [version]$_.VersionInfo.FileVersion } -Descending |
                Select-Object -First 1
        }
        if ($msalDll) {
            Add-Type -Path $msalDll.FullName -ErrorAction SilentlyContinue
        } else {
            throw 'Nie znaleziono Microsoft.Identity.Client.dll. Zainstaluj Microsoft.Graph: Install-Module Microsoft.Graph -Scope CurrentUser'
        }
    }
    if (-not $script:MsalApp) {
        $script:MsalApp = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create('14d82eec-204b-4c2f-b7e8-296a70dab67e').
            WithAuthority('https://login.microsoftonline.com/organizations').
            WithRedirectUri('http://localhost').
            Build()
    }
}

# Pobierz access token przez MSAL (silent z cache lub interaktywna przeglądarka)
function Get-CISMsalToken {
    param([string[]]$Scopes, [string]$Label = 'Microsoft 365')

    Initialize-CISMsal

    $accounts = $script:MsalApp.GetAccountsAsync().GetAwaiter().GetResult()
    $task = $null
    if ($accounts) {
        try {
            $task = $script:MsalApp.AcquireTokenSilent($Scopes, ($accounts | Select-Object -First 1)).ExecuteAsync()
            $task.Wait(8000) | Out-Null
            if ($task.IsFaulted) { $task = $null }
        } catch { $task = $null }
    }
    if (-not $task -or -not $task.IsCompleted -or $task.IsFaulted) {
        Write-CISLog "Otwieram przegladarke do logowania ($Label)..." INFO
        $task = $script:MsalApp.AcquireTokenInteractive($Scopes).WithUseEmbeddedWebView($false).ExecuteAsync()
        if ([System.Windows.Application]::Current) {
            $dcFrame  = [System.Windows.Threading.DispatcherFrame]::new()
            $deadline = [DateTime]::UtcNow.AddMinutes(10)
            $dcTimer  = [System.Windows.Threading.DispatcherTimer]::new()
            $dcTimer.Interval = [TimeSpan]::FromMilliseconds(300)
            $dcTimer.Add_Tick({
                if ($task.IsCompleted -or $task.IsFaulted -or $task.IsCanceled -or [DateTime]::UtcNow -gt $deadline) {
                    $dcTimer.Stop(); $dcFrame.Continue = $false
                }
            }.GetNewClosure())
            $dcTimer.Start()
            [System.Windows.Threading.Dispatcher]::PushFrame($dcFrame)
        } else {
            $task.Wait([TimeSpan]::FromMinutes(10)) | Out-Null
        }
    }

    if ($task.IsFaulted)       { throw $task.Exception.GetBaseException() }
    if ($task.IsCanceled)      { throw "Logowanie $Label anulowane." }
    if (-not $task.IsCompleted){ throw "Timeout logowania $Label (10 min)." }
    return $task.Result.AccessToken
}

function Connect-CISGraphInteractive {
    param([string[]]$Scopes)
    Write-CISLog 'Logowanie do Microsoft Graph (przeglądarka)...' INFO
    $token = Get-CISMsalToken -Scopes $Scopes -Label 'Microsoft Graph'
    $secToken = ConvertTo-SecureString $token -AsPlainText -Force
    Connect-MgGraph -AccessToken $secToken -NoWelcome -ErrorAction Stop
    Write-CISLog 'Microsoft Graph - polaczono.' OK
}

# ---------- POLACZENIA ----------
function Connect-CISServices {
    [CmdletBinding()]
    param(
        [switch]$SkipEntra, [switch]$SkipExchange, [switch]$SkipSharePoint,
        [switch]$SkipTeams, [switch]$SkipIntune, [switch]$SkipPurview,
        [switch]$SkipPowerPlatform,
        [string]$TenantDomain,
        [ValidateSet('ReportOnly','Enabled')][string]$ConditionalAccessState='ReportOnly'
    )
    Reset-CISContext | Out-Null
    $script:Ctx.CaState = if ($ConditionalAccessState -eq 'Enabled') { 'enabled' } else { 'enabledForReportingButNotEnforced' }

    if (-not $SkipEntra) {
        # PS5.1 ma limit 4096 funkcji - import calego Microsoft.Graph (meta) go przekracza.
        # Instalujemy meta-modul tylko po to by upewnic sie ze podmoduly sa dostepne,
        # a importujemy wylacznie te ktore sa rzeczywiscie potrzebne.
        Confirm-CISModule 'Microsoft.Graph' -OnlyInstall
        $graphSubModules = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Identity.SignIns',
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.DeviceManagement',
            'Microsoft.Graph.DeviceManagement.Administration'
        )
        foreach ($m in $graphSubModules) {
            if (-not (Get-Module -ListAvailable -Name $m)) {
                Write-CISLog "Instaluje $m..." WARN
                Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            }
            Write-CISLog "Importuje $m..." INFO
            Import-Module -Name $m -Force -ErrorAction Stop
        }
        if (-not (Get-Command 'Connect-MgGraph' -ErrorAction SilentlyContinue)) {
            throw "Connect-MgGraph niedostepne. Zainstaluj recznie:`n  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force"
        }
        $mgScopes = @(
            'Policy.ReadWrite.ConditionalAccess','Policy.Read.All','Application.Read.All',
            'Policy.ReadWrite.Authorization','Policy.ReadWrite.AuthenticationMethod',
            'Directory.ReadWrite.All','User.ReadWrite.All','Domain.ReadWrite.All',
            'RoleManagement.ReadWrite.Directory','Organization.Read.All',
            'DeviceManagementConfiguration.ReadWrite.All','DeviceManagementServiceConfig.ReadWrite.All',
            'DeviceManagementManagedDevices.Read.All'
        )
        Connect-CISGraphInteractive -Scopes $mgScopes
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
        Confirm-CISModule 'ExchangeOnlineManagement' -MinVersion '3.2.0'
        $exoMod = Get-Module ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
        $script:ExoModFile = $exoMod.Path
        Write-CISLog "Lacze z Exchange Online (EXO v$($exoMod.Version))..." INFO
        # EXO ma wlasny MSAL. Podajemy UPN z Graph context - SSO powinno przejsc cicho (bez przegladarki).
        $mgCtx = Get-MgContext
        $exoUpn = if ($mgCtx.Account) { $mgCtx.Account } else { $null }
        if ($exoUpn) {
            Connect-ExchangeOnline -UserPrincipalName $exoUpn -ShowBanner:$false -ErrorAction Stop
        } else {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
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
            $spoAdminUrl = "https://$tn-admin.sharepoint.com"
            Write-CISLog ("Lacze z SharePoint Admin ({0})..." -f $spoAdminUrl) INFO
            # Connect-SPOService uzywa nowoczesnego auth (MSAL) - z SSO po zalogowaniu do Graph powinno przejsc cicho
            Connect-SPOService -Url $spoAdminUrl -ErrorAction Stop
            $script:Ctx.Connected.SPO = $true
            Write-CISLog 'SharePoint Online - polaczono.' OK
        }
    }
    if (-not $SkipTeams) {
        Confirm-CISModule 'MicrosoftTeams'
        $teamsMod = Get-Module MicrosoftTeams | Sort-Object Version -Descending | Select-Object -First 1
        $teamsModFile = $teamsMod.Path
        $teamsLoad = if ($teamsModFile) { "Import-Module '$teamsModFile' -Force -ErrorAction Stop" } else { "Import-Module MicrosoftTeams -Force -ErrorAction Stop" }
        Write-CISLog 'Lacze z Microsoft Teams (kod urzadzenia)...'
        Invoke-CISDeviceConnect -ServiceName 'Microsoft Teams' -ConnectInvoke "$teamsLoad`nConnect-MicrosoftTeams -UseDeviceAuthentication -ErrorAction Stop | Out-Null"
        $script:Ctx.Connected.Teams = $true
    }
    if (-not $SkipPurview) {
        # ExchangeOnlineManagement dostarcza Connect-IPPSSession (Security & Compliance / Purview)
        if (-not $script:ExoModFile) {
            Confirm-CISModule 'ExchangeOnlineManagement' -MinVersion '3.2.0'
        }
        Write-CISLog 'Lacze z Microsoft Purview / Security & Compliance...' INFO
        $mgCtxP = Get-MgContext
        $purviewUpn = if ($mgCtxP.Account) { $mgCtxP.Account } else { $null }
        if ($purviewUpn) {
            Connect-IPPSSession -UserPrincipalName $purviewUpn -ShowBanner:$false -ErrorAction Stop
        } else {
            Connect-IPPSSession -ShowBanner:$false -ErrorAction Stop
        }
        $script:Ctx.Connected.Purview = $true
    }
    if (-not $SkipPowerPlatform) {
        try {
            Confirm-CISModule 'Microsoft.PowerApps.Administration.PowerShell' -MinVersion '2.0.0' -ErrorAction SilentlyContinue
            Write-CISLog 'Lacze z Power Platform...' INFO
            Add-PowerAppsAccount -ErrorAction Stop
            $script:Ctx.Connected.PowerPlatform = $true
            Write-CISLog 'Power Platform - polaczono.' OK
        } catch {
            Write-CISLog ("Power Platform pominiety: {0}" -f $_.Exception.Message) WARN
            $script:Ctx.Connected.PowerPlatform = $false
        }
    } else { $script:Ctx.Connected.PowerPlatform = $false }
    return $script:Ctx
}
function Disconnect-CISServices {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-IPPSSession -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch { }
    try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null } catch { }
}

# ---------- UZYTKOWNICY / BREAK-GLASS ----------
function Get-CISUsers {
    Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled -ErrorAction SilentlyContinue |
        Select-Object DisplayName, UserPrincipalName, AccountEnabled, Id | Sort-Object DisplayName
}
function Get-CISGlobalAdmins {
    $gaTemplateId = '62e90394-69f5-4237-9190-012177145e10'   # Global Administrator role template
    try {
        $role = Get-MgDirectoryRole -All -ErrorAction Stop | Where-Object { $_.RoleTemplateId -eq $gaTemplateId }
        if (-not $role) {
            # Aktywuj role jesli jeszcze nie aktywowana w tenancie
            Enable-MgDirectoryRole -RoleTemplateId $gaTemplateId -ErrorAction Stop | Out-Null
            $role = Get-MgDirectoryRole -All | Where-Object { $_.RoleTemplateId -eq $gaTemplateId }
        }
        if ($role) {
            return Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop |
                ForEach-Object { Get-MgUser -UserId $_.Id -Property Id,DisplayName,UserPrincipalName,AccountEnabled -ErrorAction SilentlyContinue } |
                Where-Object { $_ } |
                Select-Object DisplayName, UserPrincipalName, AccountEnabled, Id |
                Sort-Object DisplayName
        }
    } catch {
        Write-CISLog ("Nie mozna pobrac Global Adminow ({0}) - zwracam wszystkich uzytkownikow." -f $_.Exception.Message) WARN
    }
    Get-CISUsers
}
function New-CISBreakGlassAccount {
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [string]$DisplayName = 'Break Glass Account'
    )
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    $bgaPwd = try { [System.Web.Security.Membership]::GeneratePassword(24, 6) }
           catch {
               $c = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()'
               $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
               $b = New-Object byte[] 32; $rng.GetBytes($b)
               -join ($b | ForEach-Object { $c[$_ % $c.Length] })
           }
    $nick = ($UserPrincipalName -split '@')[0] -replace '[^a-zA-Z0-9]',''
    $params = @{
        DisplayName           = $DisplayName
        UserPrincipalName     = $UserPrincipalName
        MailNickname          = $nick
        AccountEnabled        = $true
        PasswordProfile       = @{ Password=$bgaPwd; ForceChangePasswordNextSignIn=$false }
        PasswordPolicies      = 'DisablePasswordExpiration'
        UsageLocation         = 'PL'
    }
    $user = New-MgUser @params -ErrorAction Stop
    Write-CISLog ("Utworzono konto BGA: {0}" -f $UserPrincipalName) OK
    # Przypisz role Global Administrator
    $gaTemplateId = '62e90394-69f5-4237-9190-012177145e10'
    try {
        $role = Get-MgDirectoryRole -All | Where-Object { $_.RoleTemplateId -eq $gaTemplateId }
        if (-not $role) {
            Enable-MgDirectoryRole -RoleTemplateId $gaTemplateId -ErrorAction Stop | Out-Null
            $role = Get-MgDirectoryRole -All | Where-Object { $_.RoleTemplateId -eq $gaTemplateId }
        }
        New-MgDirectoryRoleMember -DirectoryRoleId $role.Id `
            -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" } `
            -ErrorAction Stop | Out-Null
        Write-CISLog ("Przypisano role Global Administrator do {0}" -f $UserPrincipalName) OK
    } catch {
        Write-CISLog ("Nie udalo sie przypisac roli GA ({0}) - przypisz recznie w Entra." -f $_.Exception.Message) WARN
    }
    return [pscustomobject]@{ User=$user; Password=$bgaPwd }
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
    param([Parameter(Mandatory)]$CisProfile, [Parameter(Mandatory)]$Scan)
    $sel = $CisProfile.select
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
    'ENTRA-APPREG'               = 'Zablokowano rejestrowanie aplikacji Azure AD przez zwykłych użytkowników (AllowedToCreateApps = false). Aplikacje może rejestrować wyłącznie administrator.'
    'ENTRA-TENANT-CREATE'        = 'Zablokowano tworzenie nowych tenantów Microsoft Entra przez zwykłych użytkowników (AllowedToCreateTenants = false).'
    'ENTRA-SECGROUP'             = 'Zablokowano tworzenie grup zabezpieczeń przez zwykłych użytkowników (AllowedToCreateSecurityGroups = false). Grupy tworzy wyłącznie administrator.'
    'ENTRA-GUEST-PERMS'          = 'Ustawiono najbardziej restrykcyjne uprawnienia gości (GuestUserRoleId = Restricted Guest User). Goście widzą tylko swój własny profil i nie mogą enumerować katalogu.'
    'ENTRA-GUEST-INVITE'         = 'Zapraszanie gości do tenanta ograniczono do administratorów i użytkowników z rolą Guest Inviter (AllowInvitesFrom = adminsAndGuestInviters).'
    'ENTRA-PORTAL'               = 'Ograniczono dostęp do portalu Microsoft Entra tylko do administratorów (EnableAdminPanelRestriction = true). Zwykli użytkownicy nie mogą przeglądać katalogu przez portal.'
    'ENTRA-M365GROUP'            = 'Zablokowano tworzenie grup Microsoft 365 przez zwykłych użytkowników (EnableGroupCreation = false w Group.Unified). Grupy tworzy wyłącznie administrator.'
    'MDO-MALWARE'                = 'Włączono filtr typów plików w polityce anti-malware (EnableFileFilter = true). Blokuje wykonywalne i potencjalnie niebezpieczne typy załączników.'
    'MDO-ANTISPAM-IN'            = 'Skonfigurowano politykę antyspam: wiadomości HC spam, phishing i HC phishing kierowane do kwarantanny. Safety Tips włączone.'
    'MDO-ANTISPAM-OUT'           = 'Polityka antyspam wychodzący: konta przekraczające limit wysyłki są blokowane (ActionWhenThresholdReached = BlockUser), podejrzane wiadomości BCC do admina.'
    'TEAMS-EXTERNAL'             = 'Zablokowano połączenia z publicznymi użytkownikami Skype w Teams (AllowPublicUsers = false).'
    'TEAMS-GUESTCALL'            = 'Wyłączono prywatne połączenia głosowe dla gości w Teams (AllowPrivateCalling = false).'
    'TEAMS-RECORDING'            = 'Wyłączono nagrywanie spotkań Teams (AllowCloudRecording = false). Zalecenie CIS L2 - rozważ, czy organizacja nie wymaga nagrywania.'
    'TEAMS-EXTCONTROL'           = 'Zablokowano zewnętrznym uczestnikom możliwość przejmowania/oddawania kontroli nad ekranem (AllowExternalParticipantGiveRequestControl = false).'
    'TEAMS-MEETINGCHAT'          = 'Czat podczas spotkań Teams dostępny tylko dla uwierzytelnionych uczestników - anonimowi użytkownicy wykluczeni (AllowMeetingChat = EnabledExceptAnonymous).'
    'ENTRA-BREAKGLASS'           = 'Audyt konta break-glass: weryfikacja czy ma zarejestrowane silne metody MFA (FIDO2/Passkey). Kontrolka tylko do odczytu - wymaga ręcznej konfiguracji.'
    'EXO-SMTPAUTH'               = 'Wyłączono SMTP AUTH globalnie (SmtpClientAuthenticationDisabled = true). Blokuje starsze klienty używające Basic Auth do wysyłania poczty - wymagany Modern Auth / OAuth.'
    'EXO-CALENDAR'               = 'Ograniczono zewnętrzne udostępnianie kalendarza do podstawowych informacji o dostępności (FreeBusySimple). Szczegóły spotkań nie są widoczne dla zewnętrznych.'
    'EXO-MAILTIPS'               = 'Włączono MailTips - ostrzeżenia wyświetlane w Outlooku przy wysyłaniu do dużych grup, zewnętrznych odbiorców i adresatów bez dostępu.'
    'SPO-UNMANAGED'              = 'Dostęp do SharePoint/OneDrive z urządzeń niezarządzanych ograniczony do trybu tylko podgląd w przeglądarce (ConditionalAccessPolicy = AllowLimitedAccess).'
    'SPO-GUESTEXPIRY'            = 'Dostęp gości do SharePoint wygasa po 60 dniach (ExternalUserExpirationRequired = true, ExternalUserExpireInDays = 60). Goście muszą być regularnie ponownie zapraszani.'
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
    'PUR-DLP-EXCHANGE'           = 'Polityka DLP obejmująca Exchange Online - wykrywanie danych wrażliwych w wiadomościach e-mail i blokada ich wysyłki poza organizację.'
    'PUR-DLP-CLOUD'              = 'Polityka DLP obejmująca SharePoint Online, OneDrive i Teams - wykrywanie i ochrona danych wrażliwych w chmurze.'
    'PUR-SENSITIVITY-LABELS'     = 'Etykiety wrażliwości (Sensitivity Labels) są opublikowane i dostępne dla użytkowników - umożliwiają klasyfikację i ochronę dokumentów oraz e-maili.'
    'PUR-RETENTION'              = 'Polityka retencji obejmuje Exchange, SharePoint i Teams - zapewnia przechowywanie danych przez wymagany okres zgodnie z przepisami.'
    'MDO-ZAP'                    = 'Zero-hour Auto Purge (ZAP) włączony dla spam, phishing i malware - retro-usuwa złośliwe wiadomości już po dostarczeniu, gdy nowa sygnatura zostanie wykryta.'
    'EXO-FORWARDINGRULES'        = 'Audyt zewnętrznych reguł przekazywania poczty - wykrywa skrzynki z ForwardingSmtpAddress oraz reguły Inbox Rules kierujące pocztę na zewnętrzne adresy. Wymaga ręcznego przeglądu.'
    'PP-ENVCONFIG'               = 'Power Platform: tylko administratorzy mogą tworzyć nowe środowiska produkcyjne. Zapobiega niekontrolowanemu tworzeniu środowisk przez użytkowników.'
    'PP-TRIALENV'                = 'Power Platform: blokada tworzenia środowisk trial (próbnych) przez zwykłych użytkowników. CIS 9.1.2 L1.'
    'PP-SHAREWITHTENANT'         = 'Power Platform: użytkownicy nie mogą udostępniać canvas apps dla całego tenanta. Ogranicza masowe udostępnianie danych.'
    'COPILOT-PLUGINS'            = 'Copilot M365: zarządzanie wtyczkami od zewnętrznych wydawców. Niezweryfikowane pluginy mogą eksfiltrować dane lub naruszać polityki DLP. CIS 1.3.1 L1.'
    'COPILOT-M365GROUPS'         = 'Copilot M365: audyt przypisania licencji — tylko autoryzowani użytkownicy powinni mieć dostęp do Copilot. Weryfikuje licencje COPILOT/M365_COPILOT.'
    'INTUNE-ENCRYPT-WIN'         = 'Intune: szyfrowanie BitLocker wymagane w polityce zgodności Windows. Chroni dane na zagubionych/skradzionych urządzeniach. CIS 5.2.1 L1.'
    'INTUNE-AV-WIN'              = 'Intune: aktywny program antywirusowy wymagany przez politykę zgodności Windows. Urządzenia bez AV są blokowane. CIS 5.2.2 L1.'
    'INTUNE-FIREWALL-WIN'        = 'Intune: zapora sieciowa (Windows Firewall) wymagana przez politykę zgodności. CIS 5.2.3 L1.'
    'INTUNE-JAILBREAK'           = 'Intune: blokada urządzeń mobilnych z jailbreak (iOS) lub root (Android). Takie urządzenia omijają kontrole bezpieczeństwa systemu. CIS 5.1.1 L1.'
    'PUR-DLP-TEAMS'              = 'Purview DLP: polityka zapobiegania utracie danych obejmująca czat i kanały Teams. Zapobiega udostępnianiu poufnych informacji przez Teams. CIS 3.3.3 L1.'
    'PUR-INSIDER-RISK'           = 'Purview Insider Risk Management: aktywna polityka wykrywająca podejrzane działania wewnętrzne (wycieki danych, sabotaż). Wymaga licencji E5 Compliance. CIS 3.5 L2.'
    'ENTRA-GA-COUNT'             = 'Liczba Global Administratorów w zakresie 2-4 kont. Mniej niż 2 = brak zapasu przy awarii. Więcej niż 4 = zbyt duża powierzchnia ataku. Maester: CIS.M365.1.1.1.'
    'ENTRA-CA-RISK-USER'         = 'Polityka Conditional Access blokująca użytkowników oznaczonych jako wysokie ryzyko przez Entra ID Identity Protection. Wymaga licencji Entra ID P2. CIS 1.2.4 L2.'
    'ENTRA-CA-RISK-SIGNIN'       = 'Polityka Conditional Access wymagająca MFA dla logowań o średnim lub wysokim ryzyku wg Identity Protection. Zatrzymuje ataki credential stuffing. CIS 1.2.3 L2.'
    'ENTRA-AUTHFIDO2'            = 'Metoda uwierzytelniania FIDO2 (passkeys, klucze sprzętowe YubiKey) włączona w Authentication Methods Policy. Silniejsza niż TOTP, odporna na phishing.'
    'MDO-PHISH-THRESHOLD'        = 'PhishThresholdLevel=3 (Aggressive) w domyślnej polityce anti-phishing. Wyższy próg = więcej wiadomości traktowanych jako phishing. ORCA 101, CIS 2.1.4.'
    'MDO-DMARC-POLICY'           = 'HonorDmarcPolicy=true: MDO respektuje politykę DMARC p=reject/quarantine. Blokuje e-maile podszywające się pod domeny z opublikowanym DMARC. CIS 2.1.x.'
    'MDO-BULK-THRESHOLD'         = 'BulkThreshold≤6: wiadomości masowe (BCL≥6) trafiają do spamu lub kwarantanny. Ogranicza komercyjne maile masowe. ORCA, CIS 2.1.5.'
    'MDO-QUARANTINE-NOTIFY'      = 'ESNEnabled=true: użytkownicy otrzymują codzienne powiadomienia o wiadomościach w kwarantannie. Mogą samodzielnie zgłosić false positive. ORCA, CIS 2.1.x.'
    'MDO-MAILBOX-INTEL'          = 'Mailbox Intelligence + IntelligenceProtection włączone: MDO analizuje historię korespondencji i blokuje próby personifikacji nadawcy. Akcja: Quarantine. ORCA.'
    'EXO-BYPASS-RULES'           = 'Audyt: brak aktywnych reguł transportu z SetSCL=-1. Takie reguły pomijają filtrowanie spamu i mogą być nadużywane przez atakujących. Wymaga ręcznego przeglądu w EAC.'
    'EXO-CONNECTORS-TLS'         = 'Wszystkie aktywne connectory przychodzące (inbound connectors) wymagają szyfrowania TLS. Zapobiega przechwyceniu poczty w transmisji.'
    'SPO-DEFAULT-LINK'           = 'DefaultSharingLinkType=Direct: domyślny link udostępniania tworzy link "tylko konkretne osoby" zamiast "każdy z linkiem". Bezpieczniejszy domyślnie. CIS 7.2.3.'
    'SPO-LINK-PERMISSION'        = 'DefaultLinkPermission=View: domyślne uprawnienie linku = podgląd (nie edycja). Użytkownik musi świadomie wybrać "Edytuj" przy udostępnianiu. CIS 7.2.4.'
    'TEAMS-LOBBY'                = 'AutoAdmittedUsers=EveryoneInCompanyExcludingGuests: goście i zewnętrzni są zatrzymywani w lobby. Organizator musi ich wpuścić. CIS 8.1.1.'
    'TEAMS-PSTN-LOBBY'           = 'AllowPSTNUsersToBypassLobby=false: osoby dzwoniące przez telefon (PSTN) muszą przejść przez lobby — nie dołączają automatycznie do spotkania. CIS 8.1.2.'
    'TEAMS-PRESENTER'            = 'DesignatedPresenterRoleMode=OrganizerOnlyUserOverride: domyślnie tylko organizator i współorganizatorzy mogą prezentować (udostępniać ekran, pulpit). Uczestnicy mogą być podwyższeni do prezentera przez organizatora. CIS 8.5.1 L1.'
    'ENTRA-SMARTLOCKOUT'         = 'Smart Lockout: LockoutThreshold≤10 (blokada po 10 nieudanych próbach), LockoutDurationInSeconds≥60 (blokada min. 60s). Broni przed atakami brute-force i password spray. Domyślne wartości Microsoft są zgodne; niebezpieczne gdy tenant zmienił na zbyt liberalne. CIS 1.2.5 L1.'
    'ENTRA-CA-DEVICE'            = 'CAP04: Conditional Access wymaga urządzenia zgodnego z Intune (compliantDevice) LUB hybrid Azure AD Joined (domainJoinedDevice). Blokuje dostęp z niezarządzanych urządzeń. Wymaga licencji Entra ID P1+. CIS 5.2.2 L2.'
    'MDO-SAFEATTACH-SPO'         = 'EnableATPForSPOTeamsODB=true: Microsoft Defender for Office 365 skanuje pliki przesyłane do SharePoint Online, OneDrive i Microsoft Teams. Wykrywa złośliwe pliki zanim użytkownicy je pobiorą. CIS 2.1.2 L2.'
    'MDO-SAFELINKS-TRACK'        = 'TrackClicks=true w politykach Safe Links: MDO rejestruje kliknięcia w przepisane linki przez użytkowników. Umożliwia audyt i wykrywanie kliknięć w linki złośliwe po incydencie. CIS 2.1.1 L1.'
    'EXO-REPORT-JUNK'            = 'ReportJunkEnabled + ReportNotJunkEnabled + ReportPhishEnabled=true w ReportSubmissionPolicy: użytkownicy mogą zgłaszać spam, phishing i prawidłowe wiadomości bezpośrednio do Microsoft. Poprawia filtry uczenia maszynowego. CIS 2.6.1 L1.'
    'SPO-EMAIL-ATTEST'           = 'EmailAttestationRequired=true: osoby, które otrzymały link gościa do SharePoint, muszą ponownie potwierdzić adres e-mail po upłynięciu EmailAttestationReAuthDays (zalecane 30 dni). Zapobiega dostępowi po zmianie właściciela konta. CIS 7.2.5 L2.'
    'ENTRA-ADMIN-NOMAILBOX'      = 'Global Administratorzy nie powinni mieć licencji Exchange/M365 (brak skrzynki pocztowej). Dedykowane konta admin bez dostępu do poczty zmniejszają powierzchnię ataku — phishing/kompromitacja konta admin nie daje dostępu do danych. CIS 1.1.2 L1.'
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
            $body = @{ defaultUserRolePermissions = @{ permissionGrantPoliciesAssigned = @() } }
            Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/v1.0/policies/authorizationPolicy' -Body $body -ContentType 'application/json' | Out-Null
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

    #----- ENTRA ID - USTAWIENIA UZYTKOWNIKOW -----
    [pscustomobject]@{
        Id='ENTRA-APPREG'; Service='Graph'; Area='Entra ID'; Cis='1.1.3'; Level=1
        Name='Zablokuj rejestrowanie aplikacji przez zwykłych użytkowników'
        Test={
            $ok = -not (Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions.AllowedToCreateApps
            New-TestResult $ok ("AllowedToCreateApps="+((Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions.AllowedToCreateApps))
        }
        Apply={ Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ allowedToCreateApps=$false } }
    },
    [pscustomobject]@{
        Id='ENTRA-TENANT-CREATE'; Service='Graph'; Area='Entra ID'; Cis='1.1.4'; Level=1
        Name='Zablokuj tworzenie nowych tenantów przez zwykłych użytkowników'
        Test={
            $ok = -not (Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions.AllowedToCreateTenants
            New-TestResult $ok ("AllowedToCreateTenants="+((Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions.AllowedToCreateTenants))
        }
        Apply={ Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ allowedToCreateTenants=$false } }
    },
    [pscustomobject]@{
        Id='ENTRA-SECGROUP'; Service='Graph'; Area='Entra ID'; Cis='1.1.6'; Level=1
        Name='Zablokuj tworzenie grup zabezpieczeń przez zwykłych użytkowników'
        Test={
            $ok = -not (Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions.AllowedToCreateSecurityGroups
            New-TestResult $ok ("AllowedToCreateSecurityGroups="+((Get-MgPolicyAuthorizationPolicy).DefaultUserRolePermissions.AllowedToCreateSecurityGroups))
        }
        Apply={ Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ allowedToCreateSecurityGroups=$false } }
    },
    [pscustomobject]@{
        Id='ENTRA-GUEST-PERMS'; Service='Graph'; Area='Entra ID'; Cis='1.1.7'; Level=1
        Name='Uprawnienia gości - najbardziej restrykcyjne'
        # GuestUserRoleId: 2af84b1e = Restricted Guest User (najrestrykcyjniejszy)
        Test={
            $roleId = (Get-MgPolicyAuthorizationPolicy).GuestUserRoleId
            $ok = ($roleId -eq '2af84b1e-32c8-42b7-82bc-daa82404023b')
            New-TestResult $ok ("GuestUserRoleId=$roleId")
        }
        Apply={ Update-MgPolicyAuthorizationPolicy -GuestUserRoleId '2af84b1e-32c8-42b7-82bc-daa82404023b' }
    },
    [pscustomobject]@{
        Id='ENTRA-GUEST-INVITE'; Service='Graph'; Area='Entra ID'; Cis='1.1.5'; Level=1
        Name='Zapraszanie gości - tylko administratorzy i Guest Inviters'
        Test={
            $v = (Get-MgPolicyAuthorizationPolicy).AllowInvitesFrom
            $ok = $v -in @('adminsAndGuestInviters','adminsOnly','none')
            New-TestResult $ok ("AllowInvitesFrom=$v")
        }
        Apply={ Update-MgPolicyAuthorizationPolicy -AllowInvitesFrom 'adminsAndGuestInviters' }
    },
    [pscustomobject]@{
        Id='ENTRA-PORTAL'; Service='Graph'; Area='Entra ID'; Cis='1.1.2'; Level=1
        Name='Ogranicz dostęp do portalu Entra tylko do administratorów'
        Test={
            try {
                $tmplId = ((Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/directorySettingTemplates').value |
                    Where-Object displayName -eq 'Authorization Policy' | Select-Object -First 1).id
                if (-not $tmplId) { return New-TestResult $false 'Brak szablonu Authorization Policy' }
                $setting = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/directorySettings').value |
                    Where-Object templateId -eq $tmplId | Select-Object -First 1
                if (-not $setting) { return New-TestResult $false 'Brak ustawienia (domyslnie: dostep dla wszystkich)' }
                $val = ($setting.values | Where-Object name -eq 'EnableAdminPanelRestriction').value
                New-TestResult ($val -eq 'true') ("EnableAdminPanelRestriction=$val")
            } catch { New-TestResult $false "Blad odczytu: $_" }
        }
        Apply={
            $templates = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/directorySettingTemplates').value
            $tmpl = $templates | Where-Object displayName -eq 'Authorization Policy' | Select-Object -First 1
            if (-not $tmpl) { throw 'Brak szablonu Authorization Policy w directorySettingTemplates' }
            $settings = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/directorySettings').value
            $setting = $settings | Where-Object templateId -eq $tmpl.id | Select-Object -First 1
            $newVals = @(@{ name='EnableAdminPanelRestriction'; value='true' })
            if ($setting) {
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/directorySettings/$($setting.id)" `
                    -Body @{ values = $newVals } -ContentType 'application/json' | Out-Null
            } else {
                $allVals = $tmpl.values | ForEach-Object {
                    @{ name=$_.name; value=if($_.name -eq 'EnableAdminPanelRestriction'){'true'}else{$_.defaultValue} }
                }
                Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/directorySettings' `
                    -Body @{ templateId=$tmpl.id; values=$allVals } -ContentType 'application/json' | Out-Null
            }
        }
    },

    [pscustomobject]@{
        Id='ENTRA-M365GROUP'; Service='Graph'; Area='Entra ID'; Cis='1.1.8'; Level=1
        Name='Zablokuj tworzenie grup Microsoft 365 przez zwykłych użytkowników'
        Test={
            try {
                $tmpl = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/directorySettingTemplates').value |
                    Where-Object displayName -eq 'Group.Unified' | Select-Object -First 1
                if (-not $tmpl) { return New-TestResult $false 'Brak szablonu Group.Unified' }
                $setting = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/directorySettings').value |
                    Where-Object templateId -eq $tmpl.id | Select-Object -First 1
                if (-not $setting) { return New-TestResult $false 'Brak ustawienia - domyslnie tworzenie wlaczone dla wszystkich' }
                $val = ($setting.values | Where-Object name -eq 'EnableGroupCreation').value
                New-TestResult ($val -eq 'false') ("EnableGroupCreation=$val")
            } catch { New-TestResult $false "Blad odczytu: $_" }
        }
        Apply={
            $templates = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/directorySettingTemplates').value
            $tmpl = $templates | Where-Object displayName -eq 'Group.Unified' | Select-Object -First 1
            if (-not $tmpl) { throw 'Brak szablonu Group.Unified' }
            $settings = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/directorySettings').value
            $setting = $settings | Where-Object templateId -eq $tmpl.id | Select-Object -First 1
            if ($setting) {
                $newVals = $setting.values | ForEach-Object {
                    @{ name=$_.name; value=if($_.name -eq 'EnableGroupCreation'){'false'}else{$_.value} }
                }
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/directorySettings/$($setting.id)" `
                    -Body @{ values=$newVals } -ContentType 'application/json' | Out-Null
            } else {
                $allVals = $tmpl.values | ForEach-Object {
                    @{ name=$_.name; value=if($_.name -eq 'EnableGroupCreation'){'false'}else{$_.defaultValue} }
                }
                Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/directorySettings' `
                    -Body @{ templateId=$tmpl.id; values=$allVals } -ContentType 'application/json' | Out-Null
            }
        }
    },

    #----- EXCHANGE ONLINE / DEFENDER FOR OFFICE 365 -----
    [pscustomobject]@{
        Id='EXO-AUDIT'; Service='EXO'; Area='Exchange'; Cis='3.1.1'; Level=1
        Name='Unified Audit Log włączony'
        Test={ $v=(Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled; New-TestResult ([bool]$v) "Enabled=$v" }
        Apply={ Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true -WarningAction SilentlyContinue | Out-Null
                Write-CISLog 'EXO-AUDIT: jesli UAL nadal nie aktywny, wlacz recznie w Purview > Audit.' WARN }
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
            Get-CASMailbox -ResultSize Unlimited | Where-Object {$_.ImapEnabled -or $_.PopEnabled} | ForEach-Object {
                Set-CASMailbox -Identity $_.ExchangeGuid -ImapEnabled $false -PopEnabled $false -ErrorAction SilentlyContinue
            }
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
    [pscustomobject]@{
        Id='MDO-MALWARE'; Service='EXO'; Area='Defender'; Cis='2.1.7'; Level=1
        Name='Anti-malware: filtr niebezpiecznych typów plików'
        Test={
            $p = Get-MalwareFilterPolicy -Identity Default -ErrorAction SilentlyContinue
            if (-not $p) { return New-TestResult $false 'Brak polityki Default' }
            New-TestResult ([bool]$p.EnableFileFilter) ("EnableFileFilter="+$p.EnableFileFilter)
        }
        Apply={ Set-MalwareFilterPolicy -Identity Default -EnableFileFilter $true }
    },
    [pscustomobject]@{
        Id='MDO-ANTISPAM-IN'; Service='EXO'; Area='Defender'; Cis='2.1.5'; Level=1
        Name='Anti-spam przychodzący: phishing i HC spam do kwarantanny'
        Test={
            $p = Get-HostedContentFilterPolicy -Identity Default -ErrorAction SilentlyContinue
            if (-not $p) { return New-TestResult $false 'Brak polityki Default' }
            $ok = ($p.HighConfidenceSpamAction -eq 'Quarantine') -and
                  ($p.PhishSpamAction -eq 'Quarantine') -and
                  ($p.HighConfidencePhishAction -eq 'Quarantine')
            New-TestResult $ok ("HCSpam="+$p.HighConfidenceSpamAction+"; Phish="+$p.PhishSpamAction+"; HCPhish="+$p.HighConfidencePhishAction)
        }
        Apply={
            Set-HostedContentFilterPolicy -Identity Default `
                -HighConfidenceSpamAction Quarantine `
                -PhishSpamAction Quarantine `
                -HighConfidencePhishAction Quarantine `
                -SpamAction MoveToJmf `
                -BulkSpamAction MoveToJmf
        }
    },
    [pscustomobject]@{
        Id='MDO-ANTISPAM-OUT'; Service='EXO'; Area='Defender'; Cis='2.1.9'; Level=1
        Name='Anti-spam wychodzący: blokuj konta rozsyłające spam'
        Test={
            $p = Get-HostedOutboundSpamFilterPolicy -Identity Default -ErrorAction SilentlyContinue
            if (-not $p) { return New-TestResult $false 'Brak polityki Default' }
            $ok = $p.ActionWhenThresholdReached -in @('BlockUser','BlockUserAndSendBounce')
            New-TestResult $ok ("Action="+$p.ActionWhenThresholdReached+"; BccSuspicious="+$p.BccSuspiciousOutboundMail)
        }
        Apply={
            Set-HostedOutboundSpamFilterPolicy -Identity Default -ActionWhenThresholdReached BlockUser
            Write-CISLog 'MDO-ANTISPAM-OUT: ActionWhenThresholdReached=BlockUser. BccSuspiciousOutboundMail pominiety - wymaga adresow odbiorcow. Skonfiguruj recznie w EAC > Anti-spam > Outbound > Notifications.' WARN
        }
    },

    [pscustomobject]@{
        Id='EXO-SMTPAUTH'; Service='EXO'; Area='Exchange'; Cis='6.5.1'; Level=1
        Name='Wyłącz SMTP AUTH globalnie (legacy protocol)'
        Test={ $v=(Get-TransportConfig).SmtpClientAuthenticationDisabled; New-TestResult ([bool]$v) "SmtpClientAuthenticationDisabled=$v" }
        Apply={ Set-TransportConfig -SmtpClientAuthenticationDisabled $true }
    },
    [pscustomobject]@{
        Id='EXO-CALENDAR'; Service='EXO'; Area='Exchange'; Cis='6.3'; Level=1
        Name='Ogranicz zewnętrzne udostępnianie kalendarza (max: podstawowe info o dostępności)'
        Test={
            $p = Get-SharingPolicy | Where-Object { $_.Default } | Select-Object -First 1
            if (-not $p) { return New-TestResult $true 'Brak domyslnej polityki udostepniania' }
            $detailed = $p.Domains | Where-Object { $_ -match 'Detail|Reviewer' }
            New-TestResult ($detailed.Count -eq 0) ("Domeny z detalami: "+($detailed -join '; '))
        }
        Apply={
            $p = Get-SharingPolicy | Where-Object { $_.Default } | Select-Object -First 1
            if ($p) { Set-SharingPolicy -Identity $p.Identity -Domains @('*:CalendarSharingFreeBusySimple') -ErrorAction Stop }
        }
    },
    [pscustomobject]@{
        Id='EXO-MAILTIPS'; Service='EXO'; Area='Exchange'; Cis='6.x'; Level=1
        Name='Włącz MailTips (ostrzeżenia przy wysyłce zewnętrznej)'
        Test={
            $c = Get-OrganizationConfig
            $ok = [bool]$c.MailTipsAllTipsEnabled -and [bool]$c.MailTipsExternalRecipientsTipsEnabled
            New-TestResult $ok ("AllTips="+$c.MailTipsAllTipsEnabled+"; External="+$c.MailTipsExternalRecipientsTipsEnabled)
        }
        Apply={ Set-OrganizationConfig -MailTipsAllTipsEnabled $true -MailTipsExternalRecipientsTipsEnabled $true -MailTipsGroupMetricsEnabled $true -MailTipsLargeAudienceThreshold 25 }
    },

    #----- SHAREPOINT ONLINE / ONEDRIVE -----
    [pscustomobject]@{
        Id='SPO-SHARING'; Service='SPO'; Area='SharePoint'; Cis='7.2'; Level=1
        Name='Ogranicz udostępnianie zewnętrzne + wygasanie linków'
        Test={ $t=Get-SPOTenant; $ok = $t.SharingCapability -in @('Disabled','ExistingExternalUserSharingOnly')
               New-TestResult $ok ("SharingCapability="+$t.SharingCapability) }
        Apply={
            # ExistingExternalUserSharingOnly wylacza anonimowe linki - RequireAnonymousLinksExpireInDays jest wtedy nieprawidlowy
            Set-SPOTenant -SharingCapability ExistingExternalUserSharingOnly -DefaultSharingLinkType Direct `
                -DefaultLinkPermission View -PreventExternalUsersFromResharing $true -WarningAction SilentlyContinue
        }
    },
    [pscustomobject]@{
        Id='SPO-LEGACYAUTH'; Service='SPO'; Area='SharePoint'; Cis='7.2.x'; Level=1
        Name='Wyłącz legacy auth w SharePoint'
        Test={ $t=Get-SPOTenant; New-TestResult (-not $t.LegacyAuthProtocolsEnabled) ("LegacyAuth="+$t.LegacyAuthProtocolsEnabled) }
        Apply={ Set-SPOTenant -LegacyAuthProtocolsEnabled $false }
    },
    [pscustomobject]@{
        Id='SPO-UNMANAGED'; Service='SPO'; Area='SharePoint'; Cis='7.1'; Level=1
        Name='Ogranicz dostęp z urządzeń niezarządzanych (tylko podgląd w przeglądarce)'
        Test={
            $t = Get-SPOTenant
            $ok = $t.ConditionalAccessPolicy -in @('AllowLimitedAccess','BlockAccess')
            New-TestResult $ok ("ConditionalAccessPolicy="+$t.ConditionalAccessPolicy)
        }
        Apply={ Set-SPOTenant -ConditionalAccessPolicy AllowLimitedAccess }
    },
    [pscustomobject]@{
        Id='SPO-GUESTEXPIRY'; Service='SPO'; Area='SharePoint'; Cis='7.4'; Level=1
        Name='Wygasanie dostępu gości SharePoint (max 60 dni)'
        Test={
            $t = Get-SPOTenant
            $ok = [bool]$t.ExternalUserExpirationRequired -and ($t.ExternalUserExpireInDays -le 60)
            New-TestResult $ok ("Required="+$t.ExternalUserExpirationRequired+"; Days="+$t.ExternalUserExpireInDays)
        }
        Apply={ Set-SPOTenant -ExternalUserExpirationRequired $true -ExternalUserExpireInDays 60 }
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
    [pscustomobject]@{
        Id='TEAMS-EXTERNAL'; Service='Teams'; Area='Teams'; Cis='8.2.1'; Level=1
        Name='Zablokuj połączenia z użytkownikami publicznymi Skype'
        Test={
            $f = Get-CsTenantFederationConfiguration
            New-TestResult (-not $f.AllowPublicUsers) ("AllowPublicUsers="+$f.AllowPublicUsers)
        }
        Apply={ Set-CsTenantFederationConfiguration -AllowPublicUsers $false }
    },
    [pscustomobject]@{
        Id='TEAMS-GUESTCALL'; Service='Teams'; Area='Teams'; Cis='8.3'; Level=1
        Name='Goście nie mogą dzwonić w Teams'
        Test={
            $g = Get-CsTeamsGuestCallingConfiguration -ErrorAction SilentlyContinue
            if (-not $g) { return New-TestResult $true 'Brak konfiguracji (domyslnie: wyłączone)' }
            New-TestResult (-not $g.AllowPrivateCalling) ("AllowPrivateCalling="+$g.AllowPrivateCalling)
        }
        Apply={ Set-CsTeamsGuestCallingConfiguration -AllowPrivateCalling $false }
    },
    [pscustomobject]@{
        Id='TEAMS-RECORDING'; Service='Teams'; Area='Teams'; Cis='8.5'; Level=2
        Name='Nagrywanie spotkań Teams: wyłączone (CIS L2)'
        Test={
            $p = Get-CsTeamsMeetingPolicy -Identity Global
            New-TestResult (-not $p.AllowCloudRecording) ("AllowCloudRecording="+$p.AllowCloudRecording)
        }
        Apply={ Set-CsTeamsMeetingPolicy -Identity Global -AllowCloudRecording $false }
    },
    [pscustomobject]@{
        Id='TEAMS-EXTCONTROL'; Service='Teams'; Area='Teams'; Cis='8.4'; Level=1
        Name='Zablokuj zewnętrznym przejmowanie kontroli nad ekranem'
        Test={
            $p = Get-CsTeamsMeetingPolicy -Identity Global
            New-TestResult (-not $p.AllowExternalParticipantGiveRequestControl) ("AllowExternalParticipantGiveRequestControl="+$p.AllowExternalParticipantGiveRequestControl)
        }
        Apply={ Set-CsTeamsMeetingPolicy -Identity Global -AllowExternalParticipantGiveRequestControl $false }
    },
    [pscustomobject]@{
        Id='TEAMS-MEETINGCHAT'; Service='Teams'; Area='Teams'; Cis='8.1.x'; Level=1
        Name='Czat na spotkaniach: anonimowi użytkownicy bez dostępu'
        Test={
            $p = Get-CsTeamsMeetingPolicy -Identity Global
            # Teams PS 6.x: MeetingChatEnabledType; starsze: AllowMeetingChat
            $val = if ($p.PSObject.Properties['MeetingChatEnabledType']) { $p.MeetingChatEnabledType } else { $p.AllowMeetingChat }
            $ok = $val -in @('EnabledExceptAnonymous','Disabled')
            New-TestResult $ok ("MeetingChat=$val")
        }
        Apply={
            $p = Get-CsTeamsMeetingPolicy -Identity Global
            if ($p.PSObject.Properties['MeetingChatEnabledType']) {
                Set-CsTeamsMeetingPolicy -Identity Global -MeetingChatEnabledType EnabledExceptAnonymous
            } else {
                Set-CsTeamsMeetingPolicy -Identity Global -AllowMeetingChat EnabledExceptAnonymous
            }
        }
    },
    [pscustomobject]@{
        Id='ENTRA-BREAKGLASS'; Service='Graph'; Area='Entra ID'; Cis='1.4'; Level=1
        Name='Konto break-glass - weryfikacja metod MFA (audyt)'
        Test={
            $bgUpn = $script:Ctx.BreakGlassUpn
            if (-not $bgUpn) { return New-TestResult $false 'Brak skonfigurowanego konta break-glass w tym sesji' }
            $bgUser = Get-MgUser -Filter "UserPrincipalName eq '$bgUpn'" -Property Id,DisplayName -ErrorAction SilentlyContinue
            if (-not $bgUser) { return New-TestResult $false "Konto '$bgUpn' nie istnieje w katalogu" }
            $methods = Get-MgUserAuthenticationMethod -UserId $bgUser.Id -ErrorAction SilentlyContinue
            $mfaMethods = $methods | Where-Object { $_.AdditionalProperties['@odata.type'] -notmatch '#microsoft.graph.passwordAuthenticationMethod' }
            New-TestResult ($mfaMethods.Count -gt 0) ("Konto=$bgUpn; MethodsMFA="+$mfaMethods.Count+"; MethodsTotal="+$methods.Count)
        }
        Apply={ Write-CISLog 'ENTRA-BREAKGLASS: tylko audyt - skonfiguruj recznie silne MFA (FIDO2/Passkey) dla konta break-glass.' WARN }
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


    [pscustomobject]@{
        Id='MDO-ZAP'; Service='EXO'; Area='Defender'; Cis='2.1.6'; Level=1
        Name='Zero-hour Auto Purge (ZAP) - retro-usuwanie malware i spam'
        Test={
            $p = Get-HostedContentFilterPolicy -Identity Default -ErrorAction SilentlyContinue
            $m = Get-MalwareFilterPolicy -Identity Default -ErrorAction SilentlyContinue
            $zapSpam   = if ($p)  { [bool]$p.SpamZapEnabled   } else { $false }
            $zapPhish  = if ($p)  { [bool]$p.PhishZapEnabled  } else { $false }
            $zapMalw   = if ($m)  { [bool]$m.ZapEnabled       } else { $false }
            $ok = $zapSpam -and $zapPhish -and $zapMalw
            New-TestResult $ok ("SpamZAP=$zapSpam; PhishZAP=$zapPhish; MalwareZAP=$zapMalw")
        }
        Apply={
            Set-HostedContentFilterPolicy -Identity Default -SpamZapEnabled $true -PhishZapEnabled $true -ErrorAction SilentlyContinue
            Set-MalwareFilterPolicy -Identity Default -ZapEnabled $true -ErrorAction SilentlyContinue
        }
    },
    [pscustomobject]@{
        Id='EXO-FORWARDINGRULES'; Service='EXO'; Area='Exchange'; Cis='6.2.2'; Level=1
        Name='Audyt: zewnętrzne reguły przekazywania poczty w skrzynkach'
        Test={
            $fwdMbx   = @(Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
                Where-Object { $_.ForwardingSmtpAddress -or ($_.ForwardingAddress -and $_.DeliverToMailboxAndForward) })
            $fwdRules = @(Get-InboxRule -Mailbox * -ErrorAction SilentlyContinue 2>$null |
                Where-Object { $_.ForwardTo -or $_.RedirectTo } |
                Where-Object { ($_.ForwardTo + $_.RedirectTo) -match '@' -and ($_.ForwardTo + $_.RedirectTo) -notmatch 'EX:' })
            $total = $fwdMbx.Count + $fwdRules.Count
            if ($total -eq 0) {
                New-TestResult $true 'Brak zewnetrznych regul przekazywania'
            } else {
                $detail = ($fwdMbx | ForEach-Object { $_.UserPrincipalName }) -join ', '
                New-TestResult $false ("Skrzynki z forwarding: $($fwdMbx.Count); Reguly: $($fwdRules.Count). $detail")
            }
        }
        Apply={ Write-CISLog 'EXO-FORWARDINGRULES: tylko audyt - przejrzyj wyniki i ręcznie usuń niepotrzebne reguły.' WARN }
    },

    #----- MICROSOFT PURVIEW (Security & Compliance) -----
    [pscustomobject]@{
        Id='PUR-DLP-EXCHANGE'; Service='Purview'; Area='Purview'; Cis='3.3.x'; Level=1
        Name='Polityka DLP dla Exchange (ochrona danych wrazliwych w poczcie)'
        Test={
            $p = Get-DlpCompliancePolicy -ErrorAction SilentlyContinue |
                 Where-Object { $_.Workload -match 'Exchange' -and $_.Mode -eq 'Enable' }
            if ($p) { New-TestResult $true ("Aktywne polityki DLP dla Exchange: " + @($p).Count) }
            else    { New-TestResult $false 'Brak aktywnej polityki DLP obejmujacej Exchange' }
        }
        Apply={
            $name = 'CIS - DLP Exchange Baseline'
            if (-not (Get-DlpCompliancePolicy -Identity $name -ErrorAction SilentlyContinue)) {
                New-DlpCompliancePolicy -Name $name -Mode Enable -ExchangeLocation All -ErrorAction Stop | Out-Null
                New-DlpComplianceRule -Name "$name - Credit Cards" -Policy $name `
                    -ContentContainsSensitiveInformation @(@{Name='Credit Card Number';minCount='1'}) `
                    -BlockAccess $true -NotifyUser 'LastModifiedBy' -ErrorAction SilentlyContinue | Out-Null
                New-DlpComplianceRule -Name "$name - PII" -Policy $name `
                    -ContentContainsSensitiveInformation @(@{Name='U.S. Individual Taxpayer Identification Number (ITIN)';minCount='1'}) `
                    -BlockAccess $false -NotifyUser 'LastModifiedBy' -GenerateAlert $true -ErrorAction SilentlyContinue | Out-Null
            } else { Set-DlpCompliancePolicy -Identity $name -Mode Enable | Out-Null }
        }
    },
    [pscustomobject]@{
        Id='PUR-DLP-CLOUD'; Service='Purview'; Area='Purview'; Cis='3.3.x'; Level=1
        Name='Polityka DLP dla SharePoint / OneDrive / Teams'
        Test={
            $p = Get-DlpCompliancePolicy -ErrorAction SilentlyContinue |
                 Where-Object { ($_.Workload -match 'SharePoint|OneDrive|Teams') -and $_.Mode -eq 'Enable' }
            if ($p) { New-TestResult $true ("Aktywne polityki DLP dla chmury: " + @($p).Count) }
            else    { New-TestResult $false 'Brak aktywnej polityki DLP dla SharePoint/OneDrive/Teams' }
        }
        Apply={
            $name = 'CIS - DLP Cloud Baseline'
            if (-not (Get-DlpCompliancePolicy -Identity $name -ErrorAction SilentlyContinue)) {
                New-DlpCompliancePolicy -Name $name -Mode Enable `
                    -SharePointLocation All -OneDriveLocation All -TeamsLocation All `
                    -ErrorAction Stop | Out-Null
                New-DlpComplianceRule -Name "$name - Sensitive Data" -Policy $name `
                    -ContentContainsSensitiveInformation @(@{Name='Credit Card Number';minCount='1'}) `
                    -BlockAccess $true -NotifyUser 'LastModifiedBy' -ErrorAction SilentlyContinue | Out-Null
            } else { Set-DlpCompliancePolicy -Identity $name -Mode Enable | Out-Null }
        }
    },
    [pscustomobject]@{
        Id='PUR-SENSITIVITY-LABELS'; Service='Purview'; Area='Purview'; Cis='3.2.x'; Level=1
        Name='Etykiety wrazliwosci (Sensitivity Labels) opublikowane'
        Test={
            $labels = Get-Label -ErrorAction SilentlyContinue
            if ($labels -and @($labels).Count -gt 0) { New-TestResult $true ("Etykiety: " + @($labels).Count) }
            else { New-TestResult $false 'Brak opublikowanych etykiet wrazliwosci' }
        }
        Apply={
            Write-CISLog 'Etykiety wrazliwosci konfiguruje sie w Purview > Information Protection > Sensitivity labels.' WARN
            Write-CISLog 'Minimalny zestaw: Public, Internal, Confidential, Highly Confidential.' INFO
        }
    },
    [pscustomobject]@{
        Id='PUR-RETENTION'; Service='Purview'; Area='Purview'; Cis='3.4.x'; Level=1
        Name='Polityka retencji obejmujaca Exchange / SharePoint / Teams'
        Test={
            $p = Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue | Where-Object { -not $_.Disabled }
            if ($p -and @($p).Count -gt 0) { New-TestResult $true ("Aktywne polityki retencji: " + @($p).Count) }
            else { New-TestResult $false 'Brak aktywnej polityki retencji' }
        }
        Apply={
            $name = 'CIS - Retention Baseline (1 rok)'
            if (-not (Get-RetentionCompliancePolicy -Identity $name -ErrorAction SilentlyContinue)) {
                New-RetentionCompliancePolicy -Name $name `
                    -ExchangeLocation All -SharePointLocation All -OneDriveLocation All `
                    -TeamsChannelLocation All -TeamsChatsLocation All -Enabled $true `
                    -ErrorAction Stop | Out-Null
                New-RetentionComplianceRule -Name "$name - Zasada" -Policy $name `
                    -RetentionDuration 365 -RetentionComplianceAction Keep `
                    -ErrorAction SilentlyContinue | Out-Null
            }
        }
    },
    #----- POWER PLATFORM -----
    [pscustomobject]@{
        Id='PP-ENVCONFIG'; Service='PowerPlatform'; Area='Power Platform'; Cis='9.1.1'; Level=1
        Name='Power Platform: tworzenie srodowisk tylko przez adminow'
        Test={
            try {
                $s = Get-TenantSettings -ErrorAction Stop
                New-TestResult ([bool]$s.disableEnvironmentCreationByNonAdminUsers) ("disableEnvCreation=$($s.disableEnvironmentCreationByNonAdminUsers)")
            } catch { New-TestResult $false "Blad odczytu ustawien PP: $_" }
        }
        Apply={
            $s = Get-TenantSettings
            $s.disableEnvironmentCreationByNonAdminUsers = $true
            Set-TenantSettings $s | Out-Null
        }
    },
    [pscustomobject]@{
        Id='PP-TRIALENV'; Service='PowerPlatform'; Area='Power Platform'; Cis='9.1.2'; Level=1
        Name='Power Platform: blokada trial environment przez uzytkownikow'
        Test={
            try {
                $s = Get-TenantSettings -ErrorAction Stop
                New-TestResult ([bool]$s.disableTrialEnvironmentCreationByNonAdminUsers) ("disableTrialEnv=$($s.disableTrialEnvironmentCreationByNonAdminUsers)")
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            $s = Get-TenantSettings
            $s.disableTrialEnvironmentCreationByNonAdminUsers = $true
            Set-TenantSettings $s | Out-Null
        }
    },
    [pscustomobject]@{
        Id='PP-SHAREWITHTENANT'; Service='PowerPlatform'; Area='Power Platform'; Cis='9.1.3'; Level=2
        Name='Power Platform: blokada udostepniania canvas app dla calego tenanta'
        Test={
            try {
                $s = Get-TenantSettings -ErrorAction Stop
                $val = $s.disableShareWithEveryoneByNonAdminUsers
                if ($null -eq $val) {
                    New-TestResult $false "Wlasciwos disableShareWithEveryoneByNonAdminUsers niedostepna w tej wersji modulu - sprawdz recznie w PPAC"
                } else {
                    New-TestResult ([bool]$val) ("disableShareAll=$val")
                }
            } catch { New-TestResult $false "Blad odczytu ustawien PP: $_" }
        }
        Apply={
            try {
                $s = Get-TenantSettings -ErrorAction Stop
                if ($null -ne $s.disableShareWithEveryoneByNonAdminUsers) {
                    $s.disableShareWithEveryoneByNonAdminUsers = $true
                    Set-TenantSettings $s | Out-Null
                } else {
                    Write-CISLog 'PP-SHAREWITHTENANT: wlasciwos niedostepna w tej wersji modulu. Skonfiguruj recznie: PPAC > Settings > Tenant settings > Canvas apps.' WARN
                }
            } catch { Write-CISLog "PP-SHAREWITHTENANT blad: $_" ERROR }
        }
    },
    #----- MICROSOFT COPILOT (M365 AI) -----
    [pscustomobject]@{
        Id='COPILOT-PLUGINS'; Service='Graph'; Area='Copilot'; Cis='1.3.1'; Level=1
        Name='Copilot: zarzadzanie wtyczkami (pluginami) od zewnetrznych wydawcow'
        Test={
            try {
                $policy = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/admin/microsoft365Apps/installation/policy' -ErrorAction Stop
                $blocked = $policy.isThirdPartyAppsBlocked
                New-TestResult ([bool]$blocked) ("isThirdPartyAppsBlocked=$blocked")
            } catch {
                New-TestResult $false "Nie mozna odczytac polityki Copilot plugins: sprawdz w M365 Admin Center > Copilot"
            }
        }
        Apply={ Write-CISLog 'COPILOT-PLUGINS: Skonfiguruj w M365 Admin Center > Copilot > Manage Plugins.' WARN }
    },
    [pscustomobject]@{
        Id='COPILOT-M365GROUPS'; Service='Graph'; Area='Copilot'; Cis='1.3.2'; Level=1
        Name='Copilot: audyt uprawnien - tylko licencjonowani uzytkownicy maja dostep'
        Test={
            try {
                $skus = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -ErrorAction Stop
                $copilotSku = $skus.value | Where-Object { $_.skuPartNumber -match 'COPILOT|M365_COPILOT' }
                if ($copilotSku) {
                    $total = ($copilotSku | Measure-Object -Property consumedUnits -Sum).Sum
                    New-TestResult $true ("Licencje Copilot: $total aktywnych. Zweryfikuj przypisanie tylko uprawnionym uzytkownikom.")
                } else {
                    New-TestResult $true 'Brak licencji Copilot M365 w tenancie (nie dotyczy).'
                }
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={ Write-CISLog 'COPILOT-M365GROUPS: Audyt reczny - sprawdz przypisanie licencji Copilot w M365 Admin Center.' WARN }
    },
    #----- INTUNE ROZSZERZONE -----
    [pscustomobject]@{
        Id='INTUNE-ENCRYPT-WIN'; Service='Intune'; Area='Intune'; Cis='5.2.1'; Level=1
        Name='Intune: wymaganie szyfrowania BitLocker (Windows)'
        Test={
            try {
                $policies = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies' -ErrorAction Stop
                $winPolicies = $policies.value | Where-Object { $_.'@odata.type' -match 'windows10' }
                if (-not $winPolicies) { return New-TestResult $false 'Brak polityk zgodnosci Windows w Intune' }
                $withEncrypt = $winPolicies | ForEach-Object {
                    $p = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($_.id)" -ErrorAction SilentlyContinue
                    $p
                } | Where-Object { $_.bitLockerEnabled -eq $true }
                New-TestResult ($withEncrypt.Count -gt 0) ("Polityki z BitLocker: $($withEncrypt.Count)/$($winPolicies.Count)")
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={ Write-CISLog 'INTUNE-ENCRYPT-WIN: Wlacz bitLockerEnabled w polityce zgodnosci Windows w Intune.' WARN }
    },
    [pscustomobject]@{
        Id='INTUNE-AV-WIN'; Service='Intune'; Area='Intune'; Cis='5.2.2'; Level=1
        Name='Intune: wymaganie aktywnego antywirusa (Windows)'
        Test={
            try {
                $policies = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies' -ErrorAction Stop
                $winPolicies = $policies.value | Where-Object { $_.'@odata.type' -match 'windows10' }
                if (-not $winPolicies) { return New-TestResult $false 'Brak polityk zgodnosci Windows' }
                $withAV = $winPolicies | ForEach-Object {
                    Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($_.id)" -ErrorAction SilentlyContinue
                } | Where-Object { $_.antivirusRequired -eq $true }
                New-TestResult ($withAV.Count -gt 0) ("Polityki z antywirus: $($withAV.Count)/$($winPolicies.Count)")
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={ Write-CISLog 'INTUNE-AV-WIN: Wlacz antivirusRequired w polityce zgodnosci Windows.' WARN }
    },
    [pscustomobject]@{
        Id='INTUNE-FIREWALL-WIN'; Service='Intune'; Area='Intune'; Cis='5.2.3'; Level=1
        Name='Intune: wymaganie aktywnej zapory sieciowej (Windows)'
        Test={
            try {
                $policies = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies' -ErrorAction Stop
                $winPolicies = $policies.value | Where-Object { $_.'@odata.type' -match 'windows10' }
                if (-not $winPolicies) { return New-TestResult $false 'Brak polityk zgodnosci Windows' }
                $withFW = $winPolicies | ForEach-Object {
                    Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($_.id)" -ErrorAction SilentlyContinue
                } | Where-Object { $_.firewallEnabled -eq $true }
                New-TestResult ($withFW.Count -gt 0) ("Polityki z firewall: $($withFW.Count)/$($winPolicies.Count)")
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={ Write-CISLog 'INTUNE-FIREWALL-WIN: Wlacz firewallEnabled w polityce zgodnosci Windows.' WARN }
    },
    [pscustomobject]@{
        Id='INTUNE-JAILBREAK'; Service='Intune'; Area='Intune'; Cis='5.1.1'; Level=1
        Name='Intune: blokada urzadzen z jailbreak/root (iOS/Android)'
        Test={
            try {
                $policies = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies' -ErrorAction Stop
                $mobilePolicies = $policies.value | Where-Object { $_.'@odata.type' -match 'ios|android' }
                if (-not $mobilePolicies) { return New-TestResult $false 'Brak polityk zgodnosci iOS/Android w Intune' }
                $withJB = $mobilePolicies | ForEach-Object {
                    Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($_.id)" -ErrorAction SilentlyContinue
                } | Where-Object { $_.jailBroken -eq 'blocked' -or $_.securityBlockJailbrokenDevices -eq $true }
                New-TestResult ($withJB.Count -gt 0) ("Polityki blokujace jailbreak: $($withJB.Count)/$($mobilePolicies.Count)")
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={ Write-CISLog 'INTUNE-JAILBREAK: Wlacz blokade jailbreak w politykach zgodnosci iOS i Android.' WARN }
    },
    #----- PURVIEW ROZSZERZONE -----
    [pscustomobject]@{
        Id='PUR-DLP-TEAMS'; Service='Purview'; Area='Purview'; Cis='3.3.3'; Level=1
        Name='Purview DLP: polityka obejmujaca czat Teams'
        Test={
            try {
                $dlp = Get-DlpCompliancePolicy -ErrorAction Stop | Where-Object {
                    $_.TeamsChannelLocation -or $_.TeamsChatsLocation
                }
                New-TestResult ($dlp.Count -gt 0) ("Polityki DLP z Teams: $($dlp.Count)")
            } catch { New-TestResult $false "Blad odczytu DLP Teams: $_" }
        }
        Apply={
            $name = 'CIS DLP Teams Policy'
            if (-not (Get-DlpCompliancePolicy -Identity $name -ErrorAction SilentlyContinue)) {
                New-DlpCompliancePolicy -Name $name -TeamsChannelLocation All -TeamsChatsLocation All -Enabled $true -ErrorAction Stop | Out-Null
            }
        }
    },
    [pscustomobject]@{
        Id='PUR-INSIDER-RISK'; Service='Purview'; Area='Purview'; Cis='3.5.x'; Level=2
        Name='Purview: Insider Risk Management - polityka aktywna'
        Test={
            try {
                $irm = Get-InsiderRiskPolicy -ErrorAction Stop
                New-TestResult ($irm.Count -gt 0) ("Polityki Insider Risk: $($irm.Count)")
            } catch {
                New-TestResult $false "Insider Risk niedostepny lub brak licencji E5/E5 Compliance: $_"
            }
        }
        Apply={ Write-CISLog 'PUR-INSIDER-RISK: Skonfiguruj Insider Risk Management w Purview Compliance Portal.' WARN }
    }

    # ---- ENTRA: Global Admin count (Maester CIS.M365.1.1.1) ----
    @{
        Id='ENTRA-GA-COUNT'; Service='Graph'; Area='Entra ID'; Cis='1.1.1'; Level=1
        Name='Entra: liczba Global Adminow w zakresie 2-4 (CIS 1.1.1)'
        Test={
            try {
                $gaId = '62e90394-69f5-4237-9190-012177145e10'
                $role = Get-MgDirectoryRole -All -ErrorAction Stop | Where-Object RoleTemplateId -eq $gaId
                $cnt  = if ($role) { @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue).Count } else { 0 }
                New-TestResult ($cnt -ge 2 -and $cnt -le 4) "Global Adminow: $cnt (zalecane 2-4)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={ Write-CISLog 'ENTRA-GA-COUNT: Dostosuj recznie liczbe Global Adminow (zalecane 2-4) w Entra ID > Roles > Global Administrator.' WARN }
    }

    # ---- ENTRA: CA policy — block high-risk users ----
    @{
        Id='ENTRA-CA-RISK-USER'; Service='Graph'; Area='Entra ID'; Cis='1.2.4'; Level=2
        Name='CA: blokuj uzytkownikow wysokiego ryzyka Identity Protection (CIS 1.2.4)'
        Test={
            try {
                $caps = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies').value
                $pol  = $caps | Where-Object {
                    ($_.state -eq 'enabled' -or $_.state -eq 'enabledForReportingButNotEnforced') -and
                    ($_.conditions.userRiskLevels -contains 'high')
                } | Select-Object -First 1
                New-TestResult ($null -ne $pol) $(if($pol){"Polityka: $($pol.displayName)"}else{"Brak polityki CA dla ryzyka uzytkownika"})
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            $st  = (Get-CISContext).CaState
            $bgId= (Get-CISContext).BgId
            $body= @{
                displayName  = 'CAP05 - Block High Risk Users'
                state        = $st
                conditions   = @{
                    users           = @{ includeUsers=@('All'); excludeUsers=@($bgId | Where-Object {$_}); excludeGroups=@() }
                    applications    = @{ includeApplications=@('All') }
                    userRiskLevels  = @('high')
                }
                grantControls = @{ operator='OR'; builtInControls=@('block') }
            }
            Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies' -Body $body -ContentType 'application/json' | Out-Null
            Write-CISLog "ENTRA-CA-RISK-USER: CAP05 Block High Risk Users wdrozono ($st)" OK
        }
    }

    # ---- ENTRA: CA policy — MFA for risky sign-ins ----
    @{
        Id='ENTRA-CA-RISK-SIGNIN'; Service='Graph'; Area='Entra ID'; Cis='1.2.3'; Level=2
        Name='CA: wymagaj MFA dla ryzykownych logowan Identity Protection (CIS 1.2.3)'
        Test={
            try {
                $caps = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies').value
                $pol  = $caps | Where-Object {
                    ($_.state -eq 'enabled' -or $_.state -eq 'enabledForReportingButNotEnforced') -and
                    $_.conditions.signInRiskLevels -and
                    ($_.conditions.signInRiskLevels -contains 'high' -or $_.conditions.signInRiskLevels -contains 'medium')
                } | Select-Object -First 1
                New-TestResult ($null -ne $pol) $(if($pol){"Polityka: $($pol.displayName)"}else{"Brak polityki CA dla ryzyka logowania"})
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            $st  = (Get-CISContext).CaState
            $bgId= (Get-CISContext).BgId
            $body= @{
                displayName  = 'CAP06 - MFA for Risky Sign-ins'
                state        = $st
                conditions   = @{
                    users             = @{ includeUsers=@('All'); excludeUsers=@($bgId | Where-Object {$_}); excludeGroups=@() }
                    applications      = @{ includeApplications=@('All') }
                    signInRiskLevels  = @('high','medium')
                }
                grantControls = @{ operator='OR'; builtInControls=@('mfa') }
            }
            Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies' -Body $body -ContentType 'application/json' | Out-Null
            Write-CISLog "ENTRA-CA-RISK-SIGNIN: CAP06 MFA for Risky Sign-ins wdrozono ($st)" OK
        }
    }

    # ---- ENTRA: FIDO2 / Passkeys auth method ----
    @{
        Id='ENTRA-AUTHFIDO2'; Service='Graph'; Area='Entra ID'; Cis='1.x'; Level=1
        Name='Entra: metoda uwierzytelniania FIDO2 (passkeys) wlaczona'
        Test={
            try {
                $cfg = Invoke-MgGraphRequest -Method GET `
                    -Uri 'https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/FIDO2' -ErrorAction Stop
                New-TestResult ($cfg.state -eq 'enabled') "FIDO2 state=$($cfg.state)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Invoke-MgGraphRequest -Method PATCH `
                -Uri 'https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/FIDO2' `
                -Body @{ '@odata.type'='#microsoft.graph.fido2AuthenticationMethodConfiguration'; state='enabled'; isSelfServiceRegistrationAllowed=$true } `
                -ContentType 'application/json' | Out-Null
            Write-CISLog 'ENTRA-AUTHFIDO2: FIDO2 (passkeys) wlaczone jako metoda uwierzytelniania' OK
        }
    }

    # ---- MDO: Phishing threshold >= 3 (Maester ORCA.101) ----
    @{
        Id='MDO-PHISH-THRESHOLD'; Service='EXO'; Area='Defender'; Cis='2.1.4'; Level=1
        Name='MDO: poziom agresywnosci anty-phishing >= 3 Aggressive (ORCA 101)'
        Test={
            try {
                $pol = Get-AntiPhishPolicy -IsDefault $true -ErrorAction Stop
                $lvl = $pol.PhishThresholdLevel
                New-TestResult ($lvl -ge 3) "PhishThresholdLevel=$lvl (zalecane >=3 Aggressive)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-AntiPhishPolicy -Identity 'Office365 AntiPhish Default' -PhishThresholdLevel 3
            Write-CISLog 'MDO-PHISH-THRESHOLD: PhishThresholdLevel=3 (Aggressive)' OK
        }
    }

    # ---- MDO: Honor DMARC policy in anti-phish ----
    @{
        Id='MDO-DMARC-POLICY'; Service='EXO'; Area='Defender'; Cis='2.1.x'; Level=1
        Name='MDO: respektowanie DMARC p=reject w filtrze anty-phishingowym'
        Test={
            try {
                $pol = Get-AntiPhishPolicy -IsDefault $true -ErrorAction Stop
                New-TestResult ($pol.HonorDmarcPolicy -eq $true) "HonorDmarcPolicy=$($pol.HonorDmarcPolicy)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-AntiPhishPolicy -Identity 'Office365 AntiPhish Default' -HonorDmarcPolicy $true
            Write-CISLog 'MDO-DMARC-POLICY: HonorDmarcPolicy=$true - polityka DMARC p=reject bedzie honorowana' OK
        }
    }

    # ---- MDO: Bulk threshold BCL <= 6 (Maester ORCA) ----
    @{
        Id='MDO-BULK-THRESHOLD'; Service='EXO'; Area='Defender'; Cis='2.1.5'; Level=1
        Name='MDO: prog BCL (bulk complaint level) <= 6 (ORCA)'
        Test={
            try {
                $pol = Get-HostedContentFilterPolicy -Identity Default -ErrorAction Stop
                $bcl = $pol.BulkThreshold
                New-TestResult ($bcl -le 6) "BulkThreshold=$bcl (zalecane <=6)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-HostedContentFilterPolicy -Identity Default -BulkThreshold 6
            Write-CISLog 'MDO-BULK-THRESHOLD: BulkThreshold=6' OK
        }
    }

    # ---- MDO: End-user quarantine notifications ----
    @{
        Id='MDO-QUARANTINE-NOTIFY'; Service='EXO'; Area='Defender'; Cis='2.1.x'; Level=1
        Name='MDO: powiadomienia kwarantanny dla uzytkownikow koncowych wlaczone'
        Test={
            try {
                $pols     = Get-QuarantinePolicy -ErrorAction Stop
                $esnCount = @($pols | Where-Object ESNEnabled -eq $true).Count
                New-TestResult ($esnCount -gt 0) "Polityki z ESN wlaczonym: $esnCount / $($pols.Count)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            try {
                Set-QuarantinePolicy -Identity DefaultGlobalTag -ESNEnabled $true -EndUserQuarantinePermissionsValue 23 -ErrorAction Stop
                Write-CISLog 'MDO-QUARANTINE-NOTIFY: ESNEnabled=$true dla DefaultGlobalTag' OK
            } catch {
                Write-CISLog "MDO-QUARANTINE-NOTIFY: $($_.Exception.Message) — skonfiguruj w Microsoft Defender portal > Policies > Quarantine policies." WARN
            }
        }
    }

    # ---- MDO: Mailbox intelligence (Maester ORCA) ----
    @{
        Id='MDO-MAILBOX-INTEL'; Service='EXO'; Area='Defender'; Cis='2.1.4'; Level=1
        Name='MDO: Mailbox Intelligence i ochrona przed personifikacja wlaczone (ORCA)'
        Test={
            try {
                $pol = Get-AntiPhishPolicy -IsDefault $true -ErrorAction Stop
                $ok  = $pol.EnableMailboxIntelligence -and $pol.EnableMailboxIntelligenceProtection
                New-TestResult $ok "MailboxIntelligence=$($pol.EnableMailboxIntelligence), IntelligenceProtection=$($pol.EnableMailboxIntelligenceProtection)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-AntiPhishPolicy -Identity 'Office365 AntiPhish Default' `
                -EnableMailboxIntelligence $true `
                -EnableMailboxIntelligenceProtection $true `
                -MailboxIntelligenceProtectionAction Quarantine
            Write-CISLog 'MDO-MAILBOX-INTEL: MailboxIntelligence+IntelligenceProtection=true, akcja=Quarantine' OK
        }
    }

    # ---- EXO: No transport rules bypassing spam (SCL=-1) ----
    @{
        Id='EXO-BYPASS-RULES'; Service='EXO'; Area='Exchange'; Cis='6.x'; Level=1
        Name='EXO: brak aktywnych regul transportu omijajacych filtrowanie spamu (SCL=-1)'
        Test={
            try {
                $bypass = @(Get-TransportRule -ErrorAction Stop | Where-Object { $_.SetSCL -eq -1 -and $_.State -eq 'Enabled' })
                New-TestResult ($bypass.Count -eq 0) ("Reguly SCL=-1: {0}" -f $(if($bypass.Count-eq 0){'brak (OK)'}else{($bypass.Name -join ', ')}))
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={ Write-CISLog 'EXO-BYPASS-RULES: Przejrzyj w EAC > Mail flow > Rules reguly z SCL=-1 i wylacz te nieuzasadnione.' WARN }
    }

    # ---- EXO: Inbound connectors require TLS ----
    @{
        Id='EXO-CONNECTORS-TLS'; Service='EXO'; Area='Exchange'; Cis='6.x'; Level=1
        Name='EXO: aktywne connectory przychodzace wymagaja TLS'
        Test={
            try {
                $conns = @(Get-InboundConnector -ErrorAction Stop | Where-Object Enabled -eq $true)
                if ($conns.Count -eq 0) { return New-TestResult $true 'Brak aktywnych connectorow przychodzacych' }
                $noTls = @($conns | Where-Object { $_.RequireTls -ne $true })
                New-TestResult ($noTls.Count -eq 0) ("Bez TLS: {0}" -f $(if($noTls.Count-eq 0){'brak (OK)'}else{($noTls.Name-join', ')}))
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            $conns = Get-InboundConnector -ErrorAction Stop | Where-Object { $_.Enabled -eq $true -and $_.RequireTls -ne $true }
            foreach ($c in $conns) {
                Set-InboundConnector -Identity $c.Identity -RequireTls $true -ErrorAction SilentlyContinue
                Write-CISLog "EXO-CONNECTORS-TLS: wlaczono RequireTls dla '$($c.Name)'" OK
            }
            if (-not $conns) { Write-CISLog 'EXO-CONNECTORS-TLS: wszystkie connectory juz wymagaja TLS' OK }
        }
    }

    # ---- SPO: Default sharing link type = Direct (specific people) ----
    @{
        Id='SPO-DEFAULT-LINK'; Service='SPO'; Area='SharePoint'; Cis='7.2.3'; Level=1
        Name='SPO: domyslny link udostepniania = tylko konkretne osoby (Direct)'
        Test={
            try {
                $spo  = Get-SPOTenant -ErrorAction Stop
                $link = $spo.DefaultSharingLinkType
                New-TestResult ($link -eq 'Direct' -or $link -eq 'Internal') "DefaultSharingLinkType=$link (zalecane Direct)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-SPOTenant -DefaultSharingLinkType Direct
            Write-CISLog 'SPO-DEFAULT-LINK: DefaultSharingLinkType=Direct (tylko konkretne osoby)' OK
        }
    }

    # ---- SPO: Default link permission = View ----
    @{
        Id='SPO-LINK-PERMISSION'; Service='SPO'; Area='SharePoint'; Cis='7.2.4'; Level=1
        Name='SPO: domyslne uprawnienie linku = tylko podglad (View)'
        Test={
            try {
                $spo  = Get-SPOTenant -ErrorAction Stop
                $perm = $spo.DefaultLinkPermission
                New-TestResult ($perm -eq 'View') "DefaultLinkPermission=$perm (zalecane View)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-SPOTenant -DefaultLinkPermission View
            Write-CISLog 'SPO-LINK-PERMISSION: DefaultLinkPermission=View' OK
        }
    }

    # ---- Teams: Lobby bypass — EveryoneInCompanyExcludingGuests ----
    @{
        Id='TEAMS-LOBBY'; Service='Teams'; Area='Teams'; Cis='8.1.1'; Level=1
        Name='Teams: lobby omijaja tylko uzytkownicy organizacji (bez gosci)'
        Test={
            try {
                $pol    = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction Stop
                $bypass = $pol.AutoAdmittedUsers
                New-TestResult ($bypass -eq 'EveryoneInCompanyExcludingGuests' -or $bypass -eq 'OrganizerOnly') "AutoAdmittedUsers=$bypass"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-CsTeamsMeetingPolicy -Identity Global -AutoAdmittedUsers 'EveryoneInCompanyExcludingGuests'
            Write-CISLog 'TEAMS-LOBBY: AutoAdmittedUsers=EveryoneInCompanyExcludingGuests' OK
        }
    }

    # ---- Teams: PSTN callers must go through lobby ----
    @{
        Id='TEAMS-PSTN-LOBBY'; Service='Teams'; Area='Teams'; Cis='8.1.2'; Level=1
        Name='Teams: dzwoniacy przez PSTN nie omijaja lobby'
        Test={
            try {
                $pol    = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction Stop
                $bypass = $pol.AllowPSTNUsersToBypassLobby
                New-TestResult ($bypass -eq $false) "AllowPSTNUsersToBypassLobby=$bypass (zalecane False)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-CsTeamsMeetingPolicy -Identity Global -AllowPSTNUsersToBypassLobby $false
            Write-CISLog 'TEAMS-PSTN-LOBBY: AllowPSTNUsersToBypassLobby=$false' OK
        }
    },

    # ---- Teams: Only organizer/co-organizer can present (CIS 8.5.1 L1) ----
    [pscustomobject]@{
        Id='TEAMS-PRESENTER'; Service='Teams'; Area='Teams'; Cis='8.5.1'; Level=1
        Name='Teams: tylko organizator/wspol-organizator moze prezentowac'
        Test={
            try {
                $pol  = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction Stop
                $mode = $pol.DesignatedPresenterRoleMode
                New-TestResult ($mode -in @('OrganizerOnlyUserOverride','OrganizerOnly')) "DesignatedPresenterRoleMode=$mode (zalecane OrganizerOnlyUserOverride)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-CsTeamsMeetingPolicy -Identity Global -DesignatedPresenterRoleMode OrganizerOnlyUserOverride
            Write-CISLog 'TEAMS-PRESENTER: DesignatedPresenterRoleMode=OrganizerOnlyUserOverride' OK
        }
    },

    # ---- Entra ID: Smart Lockout (CIS 1.2.5 L1) ----
    [pscustomobject]@{
        Id='ENTRA-SMARTLOCKOUT'; Service='Graph'; Area='Entra ID'; Cis='1.2.5'; Level=1
        Name='Entra ID: Smart Lockout — prog blokady ≤10, czas trwania ≥60s'
        Test={
            try {
                $resp       = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/settings' -ErrorAction Stop
                $pwdSetting = $resp.value | Where-Object { $_.displayName -eq 'Password Rule Settings' }
                if (-not $pwdSetting) {
                    return New-TestResult $true 'Uzywa domyslnych wartosci MS (threshold=10, duration=60s) — zgodne'
                }
                $thresh = [int](($pwdSetting.values | Where-Object name -eq 'LockoutThreshold').value)
                $dur    = [int](($pwdSetting.values | Where-Object name -eq 'LockoutDurationInSeconds').value)
                New-TestResult ($thresh -le 10 -and $dur -ge 60) "LockoutThreshold=$thresh (max 10), LockoutDurationInSeconds=$dur (min 60s)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            try {
                $resp       = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/settings' -ErrorAction Stop
                $pwdSetting = $resp.value | Where-Object { $_.displayName -eq 'Password Rule Settings' }
                $vals = @(
                    @{ name='LockoutThreshold';         value='10' }
                    @{ name='LockoutDurationInSeconds'; value='60' }
                    @{ name='BannedPasswordCheckOnPremisesMode'; value='Audit' }
                    @{ name='EnableBannedPasswordCheckOnPremises'; value='True' }
                    @{ name='EnableBannedPasswordCheck'; value='True' }
                    @{ name='LockoutObservationWindow'; value='60' }
                )
                if ($pwdSetting) {
                    foreach ($v in $vals) {
                        $existing = $pwdSetting.values | Where-Object name -eq $v.name
                        if ($existing) { $existing.value = $v.value }
                    }
                    Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/beta/settings/$($pwdSetting.id)" `
                        -Body @{ values = $pwdSetting.values } -ContentType 'application/json' | Out-Null
                } else {
                    $templateId = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/directorySettingTemplates' -ErrorAction Stop).value |
                        Where-Object displayName -eq 'Password Rule Settings' | Select-Object -First 1 -ExpandProperty id
                    if ($templateId) {
                        Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/beta/settings' `
                            -Body @{ templateId=$templateId; values=$vals } -ContentType 'application/json' | Out-Null
                    }
                }
                Write-CISLog 'ENTRA-SMARTLOCKOUT: LockoutThreshold=10, LockoutDurationInSeconds=60' OK
            } catch { throw }
        }
    },

    # ---- Entra ID: CA04 — wymagaj zgodnego lub hybrid-joined urzadzenia (L2) ----
    [pscustomobject]@{
        Id='ENTRA-CA-DEVICE'; Service='Graph'; Area='Entra ID'; Cis='5.2.2'; Level=2
        Name='Conditional Access: wymagaj urzadzenia zgodnego (Intune) lub hybrid-joined (CAP04)'
        Test={
            $p = Get-MgIdentityConditionalAccessPolicy -All | Where-Object {
                $_.State -ne 'disabled' -and
                ($_.GrantControls.BuiltInControls -contains 'compliantDevice' -or
                 $_.GrantControls.BuiltInControls -contains 'domainJoinedDevice')
            }
            if ($p) { New-TestResult $true ("Polityka: " + $p[0].DisplayName) }
            else    { New-TestResult $false 'Brak polityki CA wymajacej zgodnego/hybrid-joined urzadzenia' }
        }
        Apply={
            $body = @{
                displayName   = 'CAP04: Require Compliant or Hybrid Joined Device'
                state         = $script:Ctx.CaState
                conditions    = @{
                    clientAppTypes = @('all')
                    applications   = @{ includeApplications = @('All') }
                    users          = @{ includeUsers = @('All'); excludeUsers = @($script:Ctx.BgId) }
                }
                grantControls = @{ operator = 'OR'; builtInControls = @('compliantDevice','domainJoinedDevice') }
            }
            $ex = Get-MgIdentityConditionalAccessPolicy -All | Where-Object DisplayName -eq $body.displayName
            if ($ex) { Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $ex.Id -BodyParameter $body | Out-Null }
            else     { New-MgIdentityConditionalAccessPolicy -BodyParameter $body | Out-Null }
            Write-CISLog "ENTRA-CA-DEVICE: polityka CAP04 utworzona/zaktualizowana (stan: $($script:Ctx.CaState))" OK
        }
    },

    # ---- MDO: Safe Attachments dla SPO/OneDrive/Teams (CIS 2.1.2 L2) ----
    [pscustomobject]@{
        Id='MDO-SAFEATTACH-SPO'; Service='EXO'; Area='Defender'; Cis='2.1.2'; Level=2
        Name='MDO: Safe Attachments wlaczone dla SharePoint, OneDrive i Teams'
        Test={
            try {
                $atp = Get-AtpPolicyForO365 -ErrorAction Stop
                New-TestResult ($atp.EnableATPForSPOTeamsODB -eq $true) "EnableATPForSPOTeamsODB=$($atp.EnableATPForSPOTeamsODB)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-AtpPolicyForO365 -EnableATPForSPOTeamsODB $true
            Write-CISLog 'MDO-SAFEATTACH-SPO: EnableATPForSPOTeamsODB=true' OK
        }
    },

    # ---- MDO: Safe Links — sledz klikniecia (CIS 2.1.1 L1) ----
    [pscustomobject]@{
        Id='MDO-SAFELINKS-TRACK'; Service='EXO'; Area='Defender'; Cis='2.1.1'; Level=1
        Name='MDO: Safe Links — sledz klikniecia uzytkownikow (TrackClicks=true)'
        Test={
            try {
                $pols = @(Get-SafeLinksPolicy -ErrorAction Stop | Where-Object { $_.IsEnabled })
                if ($pols.Count -eq 0) { return New-TestResult $false 'Brak aktywnych polityk Safe Links' }
                $noTrack = @($pols | Where-Object { $_.TrackClicks -ne $true })
                New-TestResult ($noTrack.Count -eq 0) ("TrackClicks=false w: {0}" -f $(if($noTrack.Count -eq 0){'brak (OK)'}else{($noTrack.Name -join ', ')}))
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            $fixed = 0
            Get-SafeLinksPolicy | Where-Object { $_.IsEnabled -and $_.TrackClicks -ne $true } | ForEach-Object {
                Set-SafeLinksPolicy -Identity $_.Identity -TrackClicks $true -ErrorAction SilentlyContinue
                Write-CISLog "MDO-SAFELINKS-TRACK: TrackClicks=true dla '$($_.Name)'" OK
                $fixed++
            }
            if ($fixed -eq 0) { Write-CISLog 'MDO-SAFELINKS-TRACK: wszystkie polityki juz maja TrackClicks=true' OK }
        }
    },

    # ---- EXO: Report Junk/Phish wlaczony (CIS 2.6.1 L1) ----
    [pscustomobject]@{
        Id='EXO-REPORT-JUNK'; Service='EXO'; Area='Exchange'; Cis='2.6.1'; Level=1
        Name='EXO: uzytkownicy moga zglaszac spam i phishing do Microsoft'
        Test={
            try {
                $pol = Get-ReportSubmissionPolicy -ErrorAction Stop
                $ok  = $pol.ReportJunkEnabled -and $pol.ReportNotJunkEnabled -and $pol.ReportPhishEnabled
                New-TestResult $ok "ReportJunk=$($pol.ReportJunkEnabled), ReportNotJunk=$($pol.ReportNotJunkEnabled), ReportPhish=$($pol.ReportPhishEnabled)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            $params = @{ ReportJunkEnabled=$true; ReportNotJunkEnabled=$true; ReportPhishEnabled=$true }
            $pol = Get-ReportSubmissionPolicy -ErrorAction SilentlyContinue
            if ($pol) { Set-ReportSubmissionPolicy @params }
            else      { New-ReportSubmissionPolicy @params | Out-Null }
            Write-CISLog 'EXO-REPORT-JUNK: ReportJunk/NotJunk/Phish=true' OK
        }
    },

    # ---- SPO: Email attestation dla udostepnianych linkow (CIS 7.2.5 L2) ----
    [pscustomobject]@{
        Id='SPO-EMAIL-ATTEST'; Service='SPO'; Area='SharePoint'; Cis='7.2.5'; Level=2
        Name='SPO: wymagaj ponownej weryfikacji e-mail dla linkow udostepniania (co 30 dni)'
        Test={
            try {
                $spo = Get-SPOTenant -ErrorAction Stop
                New-TestResult ($spo.EmailAttestationRequired -eq $true) "EmailAttestationRequired=$($spo.EmailAttestationRequired), ReAuthDays=$($spo.EmailAttestationReAuthDays)"
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={
            Set-SPOTenant -EmailAttestationRequired $true -EmailAttestationReAuthDays 30
            Write-CISLog 'SPO-EMAIL-ATTEST: EmailAttestationRequired=true, ReAuthDays=30' OK
        }
    },

    # ---- Entra ID: Global Admins bez licencji Exchange (CIS 1.1.2 L1) ----
    [pscustomobject]@{
        Id='ENTRA-ADMIN-NOMAILBOX'; Service='Graph'; Area='Entra ID'; Cis='1.1.2'; Level=1
        Name='Entra ID: konta Global Admin nie maja licencji Exchange (dedykowane konta admin)'
        Test={
            try {
                $role   = Get-MgDirectoryRole -Filter "displayName eq 'Global Administrator'" -ErrorAction Stop | Select-Object -First 1
                if (-not $role) { return New-TestResult $false 'Nie znaleziono roli Global Administrator' }
                $admins = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction Stop
                # Exchange/M365 plans zawierające mailbox: EXCHANGESTANDARD, EXCHANGEENTERPRISE, ENTERPRISEPREMIUM, SPE_E3, SPE_E5 itp.
                $exchSkus = @('EXCHANGESTANDARD','EXCHANGEENTERPRISE','ENTERPRISEPREMIUM','STANDARDPACK','ENTERPRISEPACK',
                              'SPE_E3','SPE_E5','MICROSOFT_BUSINESS_CENTER','O365_BUSINESS_PREMIUM','O365_BUSINESS_ESSENTIALS',
                              'EXCHANGEARCHIVE','EXCHANGEDESKLESS','M365EDU_A3_FACULTY','M365EDU_A5_FACULTY')
                $withMailbox = @()
                foreach ($m in $admins) {
                    try {
                        $u = Get-MgUser -UserId $m.Id -Property 'userPrincipalName,assignedLicenses' -ErrorAction Stop
                        if ($u.AssignedLicenses.Count -gt 0) {
                            # Sprawdz czy ma SKU z Exchange
                            $skus = Get-MgSubscribedSku -ErrorAction SilentlyContinue
                            $hasExch = $u.AssignedLicenses | ForEach-Object {
                                $skuId = $_.SkuId
                                $sku = $skus | Where-Object SkuId -eq $skuId | Select-Object -First 1
                                if ($sku -and ($exchSkus -contains $sku.SkuPartNumber -or $sku.SkuPartNumber -match 'EXCHANGE|ENTERPRISE|M365|O365|SPE_E')) {
                                    $u.UserPrincipalName
                                }
                            } | Select-Object -First 1
                            if ($hasExch) { $withMailbox += $u.UserPrincipalName }
                        }
                    } catch {}
                }
                $ok = ($withMailbox.Count -eq 0)
                New-TestResult $ok ("Global Admins z licencja Exchange: {0}" -f $(if($ok){'brak (OK)'}else{($withMailbox -join ', ')}))
            } catch { New-TestResult $false "Blad: $_" }
        }
        Apply={ Write-CISLog 'ENTRA-ADMIN-NOMAILBOX: Usun licencje Exchange z kont Global Admin lub uzyj dedykowanych kont admin bez licencji uzytkownika. Wymagana akcja reczna.' WARN }
    }
)
    return $ControlRegistry
}


# ---------- EKSPORT ----------
function Export-CISScanToCsv {
    param(
        [Parameter(Mandatory)][object[]]$Scan,
        [Parameter(Mandatory)][string]$Path
    )
    $rows = $Scan | ForEach-Object {
        $doc = $script:ControlDocs[$_.Id]; if (-not $doc) { $doc = '' }
        [pscustomobject]@{
            ID          = $_.Id
            Obszar      = $_.Obszar
            CIS         = $_.CIS
            Poziom      = $_.Poziom
            Kontrolka   = $_.Kontrolka
            Status      = $_.Status
            Stan        = $_.Aktualnie
            Opis        = $doc
        }
    }
    $rows | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation -Delimiter ';'
    Write-CISLog "Eksport CSV: $Path ($(($rows).Count) wierszy)" OK
}

function Export-CISReportToWord {
    param(
        [Parameter(Mandatory)][object[]]$Scan,
        [object[]]$Applied,
        [hashtable]$Context,
        [Parameter(Mandatory)][string]$Path,
        [switch]$SaveAsPdf
    )
    try { $word = New-Object -ComObject Word.Application -ErrorAction Stop }
    catch { throw 'Microsoft Word nie jest zainstalowany. Zainstaluj Word lub użyj eksportu HTML.' }

    $word.Visible = $false
    $wordDoc = $word.Documents.Add()
    $sel = $word.Selection

    # --- Styl nagłówka dokumentu ---
    $wordDoc.BuiltInDocumentProperties("Title")  = 'Raport zgodności CIS M365'
    $wordDoc.BuiltInDocumentProperties("Author") = try { (Get-MgContext).Account } catch { $env:USERNAME }

    $tenant  = if ($Context.TenantDisplayName) { $Context.TenantDisplayName } else { $Context.TenantInitialDomain }
    $domain  = $Context.TenantInitialDomain
    $now     = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $operator = try { (Get-MgContext).Account } catch { $env:USERNAME }

    $total = $Scan.Count
    $ok    = @($Scan | Where-Object Status -eq 'Zgodne').Count
    $nok   = @($Scan | Where-Object Status -eq 'NIEZGODNE').Count
    $warn  = @($Scan | Where-Object Status -eq 'WARN').Count
    $pct   = if ($total -gt 0) { [int]([math]::Round($ok / $total * 100)) } else { 0 }

    # Kolory Word (RGB jako long int)
    $colorBlue   = 0x0B5CAB   # niebieski nagłówek
    $colorGreen  = 0x1A7F37   # zielony OK
    $colorRed    = 0xB42318   # czerwony NIEZGODNE
    $colorOrange = 0x9A6700   # pomarańczowy WARN
    $colorGray   = 0x555555

    # === STRONA TYTUŁOWA ===
    $sel.ParagraphFormat.Alignment = 1  # wyśrodkuj
    $sel.Font.Size  = 28; $sel.Font.Bold = $true; $sel.Font.Color = $colorBlue
    $sel.TypeText("Raport zgodności M365")
    $sel.TypeParagraph()
    $sel.Font.Size  = 20; $sel.Font.Color = $colorBlue
    $sel.TypeText("CIS Microsoft 365 Foundations Benchmark v6.x")
    $sel.TypeParagraph(); $sel.TypeParagraph()
    $sel.Font.Size  = 14; $sel.Font.Bold = $false; $sel.Font.Color = $colorGray
    $sel.TypeText("Klient: $tenant")
    $sel.TypeParagraph()
    $sel.TypeText("Domena: $domain")
    $sel.TypeParagraph()
    $sel.TypeText("Data: $now")
    $sel.TypeParagraph()
    $sel.TypeText("Wykonał: $operator")
    $sel.TypeParagraph(); $sel.TypeParagraph()

    # Wynik procentowy na stronie tytułowej
    $scoreColor = if ($pct -ge 80) { $colorGreen } elseif ($pct -ge 50) { $colorOrange } else { $colorRed }
    $sel.Font.Size = 48; $sel.Font.Bold = $true; $sel.Font.Color = $scoreColor
    $sel.TypeText("$pct%")
    $sel.TypeParagraph()
    $sel.Font.Size = 14; $sel.Font.Bold = $false; $sel.Font.Color = $colorGray
    $sel.TypeText("Zgodnosc CIS: $ok / $total kontrolek OK")
    $sel.TypeParagraph()

    # Nowa strona
    $sel.InsertBreak(7)  # wdPageBreak
    $sel.ParagraphFormat.Alignment = 0  # do lewej

    # === PODSUMOWANIE ===
    $sel.Font.Size = 16; $sel.Font.Bold = $true; $sel.Font.Color = $colorBlue
    $sel.TypeText("1. Podsumowanie wykonawcze")
    $sel.TypeParagraph()
    $sel.Font.Size = 11; $sel.Font.Bold = $false; $sel.Font.Color = 0x000000
    $sel.TypeText("Tenant: $tenant ($domain)")
    $sel.TypeParagraph()
    $sel.TypeText("Data skanu: $now | Wykonał: $operator")
    $sel.TypeParagraph()
    $sel.TypeText("Benchmark: CIS Microsoft 365 Foundations Benchmark v6.x (L1/L2)")
    $sel.TypeParagraph(); $sel.TypeParagraph()

    $sel.Font.Color = $colorGreen;  $sel.TypeText("OK (zgodne):      $ok")
    $sel.TypeParagraph()
    $sel.Font.Color = $colorRed;    $sel.TypeText("NIEZGODNE:        $nok")
    $sel.TypeParagraph()
    $sel.Font.Color = $colorOrange; $sel.TypeText("WARN (ostrzezenia): $warn")
    $sel.TypeParagraph()
    $sel.Font.Color = 0x000000;     $sel.TypeText("Lacznie kontrolek: $total")
    $sel.TypeParagraph()
    $sel.Font.Color = $scoreColor;  $sel.Font.Bold = $true
    $sel.TypeText("Wynik zgodnosci: $pct%")
    $sel.Font.Bold = $false; $sel.Font.Color = 0x000000
    $sel.TypeParagraph(); $sel.TypeParagraph()

    if ($Applied -and $Applied.Count -gt 0) {
        $appliedOk = @($Applied | Where-Object Status -eq 'APPLIED').Count
        $appliedErr = @($Applied | Where-Object Status -eq 'ERROR').Count
        $sel.TypeText("W tej sesji wdrozono $($Applied.Count) kontrolek (OK: $appliedOk, Bledy: $appliedErr).")
        $sel.TypeParagraph(); $sel.TypeParagraph()
    }

    # Nowa strona
    $sel.InsertBreak(7)

    # === TABELA WYNIKOW ===
    $sel.Font.Size = 16; $sel.Font.Bold = $true; $sel.Font.Color = $colorBlue
    $sel.TypeText("2. Szczegolowe wyniki kontrolek")
    $sel.TypeParagraph(); $sel.Font.Bold = $false; $sel.Font.Size = 11; $sel.Font.Color = 0x000000

    $tbl = $wordDoc.Tables.Add($sel.Range, $total + 1, 5)
    $tbl.Style = "Table Grid"
    $tbl.Borders.Enable = $true

    # Nagłówek tabeli
    $headers = @('Obszar', 'CIS', 'Kontrolka', 'Status', 'Stan zastany')
    for ($c = 1; $c -le 5; $c++) {
        $cell = $tbl.Cell(1, $c)
        $cell.Range.Text = $headers[$c - 1]
        $cell.Range.Font.Bold = $true
        $cell.Range.Font.Color = 0xFFFFFF
        $cell.Shading.BackgroundPatternColor = $colorBlue
    }
    # Szerokości kolumn (cm → points: *28.35)
    $tbl.Columns(1).Width = 2.8 * 28.35  # Obszar
    $tbl.Columns(2).Width = 1.6 * 28.35  # CIS
    $tbl.Columns(3).Width = 7.0 * 28.35  # Kontrolka
    $tbl.Columns(4).Width = 2.2 * 28.35  # Status
    $tbl.Columns(5).Width = 5.0 * 28.35  # Stan

    for ($i = 0; $i -lt $Scan.Count; $i++) {
        $row  = $Scan[$i]
        $r    = $i + 2
        $ctrlDoc = $script:ControlDocs[$row.Id]
        $name = if ($ctrlDoc) { "$($row.Kontrolka)`n$ctrlDoc" } else { $row.Kontrolka }

        $tbl.Cell($r, 1).Range.Text = $row.Obszar
        $tbl.Cell($r, 2).Range.Text = $row.CIS
        $tbl.Cell($r, 3).Range.Text = $name
        $tbl.Cell($r, 4).Range.Text = $row.Status
        $tbl.Cell($r, 5).Range.Text = $row.Aktualnie

        $statColor = switch ($row.Status) {
            'Zgodne'    { $colorGreen }
            'NIEZGODNE' { $colorRed }
            'WARN'      { $colorOrange }
            default     { 0x000000 }
        }
        $tbl.Cell($r, 4).Range.Font.Color = $statColor
        $tbl.Cell($r, 4).Range.Font.Bold  = $true
        $tbl.Cell($r, 3).Range.Font.Size  = 9
        $tbl.Cell($r, 5).Range.Font.Size  = 9

        # Pastelowe tło wiersza
        $bgColor = switch ($row.Status) {
            'Zgodne'    { 0xE8F5E9 }
            'NIEZGODNE' { 0xFFEBEE }
            'WARN'      { 0xFFF8E1 }
            default     { 0xFFFFFF }
        }
        for ($c = 1; $c -le 5; $c++) {
            $tbl.Cell($r, $c).Shading.BackgroundPatternColor = $bgColor
        }
    }

    # Przesuń kursor za tabelę
    $sel.MoveDown(5, 1) | Out-Null  # wdLine

    # Zapisz
    if ($SaveAsPdf) {
        $pdfPath = [IO.Path]::ChangeExtension($Path, '.pdf')
        $wordDoc.SaveAs2([ref]$pdfPath, [ref]17)  # 17 = wdFormatPDF
        Write-CISLog "Raport PDF: $pdfPath" OK
    }
    $wordDoc.SaveAs2([ref]$Path, [ref]16)  # 16 = wdFormatDocx
    $wordDoc.Close()
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-CISLog "Raport Word: $Path" OK
    return $Path
}

function Send-CISReportByEmail {
    param(
        [Parameter(Mandatory)][string]$To,
        [Parameter(Mandatory)][object[]]$Scan,
        [object[]]$Applied,
        [hashtable]$Context,
        [string]$HtmlReportPath
    )
    $from     = (Get-MgContext).Account
    $tenant   = if ($Context.TenantDisplayName) { $Context.TenantDisplayName } else { $Context.TenantInitialDomain }
    $now      = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $total    = $Scan.Count
    $ok       = @($Scan | Where-Object Status -eq 'Zgodne').Count
    $nok      = @($Scan | Where-Object Status -eq 'NIEZGODNE').Count
    $pct      = if ($total -gt 0) { [int]([math]::Round($ok / $total * 100)) } else { 0 }
    $subject  = "Raport CIS M365 — $tenant — $now — Zgodnosc: $pct%"

    # Treść e-mail (HTML)
    $body = if ($HtmlReportPath -and (Test-Path $HtmlReportPath)) {
        Get-Content $HtmlReportPath -Raw -Encoding UTF8
    } else {
        $nok2 = $nok
        @"
<h2>Raport zgodności CIS M365 — $tenant</h2>
<p>Data: $now | Wykonał: $from</p>
<p>Wynik: <strong>$pct% ($ok/$total OK)</strong> | Niezgodne: $nok2</p>
<p>Szczegóły: wygeneruj raport HTML lub Word z aplikacji M365 CIS Assistant.</p>
"@
    }

    $msgBody = @{
        message = @{
            subject = $subject
            body    = @{ contentType = 'HTML'; content = $body }
            toRecipients = @(@{ emailAddress = @{ address = $To } })
        }
        saveToSentItems = $true
    }
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$from/sendMail" -Body $msgBody -ContentType 'application/json' | Out-Null
    Write-CISLog "E-mail z raportem wyslany do: $To" OK
}

function Get-CISControlDocs { $script:ControlDocs }

Export-ModuleMember -Function `
    Set-CISLogCallback, Set-CISDeviceCodeCallback, Write-CISLog, Confirm-CISModule, New-TestResult, ConvertTo-HtmlText, `
    Reset-CISContext, Get-CISContext, Set-CISContext, Connect-CISServices, Disconnect-CISServices, `
    Get-CISUsers, Get-CISGlobalAdmins, New-CISBreakGlassAccount, Set-CISBreakGlass, `
    Invoke-CISScan, Invoke-CISApply, Remove-CISLegacyPolicies, `
    Get-CISProfileSelection, Import-CISProfile, Save-CISProfile, `
    New-DeploymentReport, Get-CISControlRegistry, Get-CISControlDocs, `
    Export-CISScanToCsv, Export-CISReportToWord, Send-CISReportByEmail
