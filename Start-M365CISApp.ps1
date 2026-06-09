#Requires -Version 5.1
<#
.SYNOPSIS
    M365 CIS Assistant - graficzny interfejs (WPF) na bazie modulu M365CISCore.psm1.
.DESCRIPTION
    Pelny przeplyw w oknie: Polacz -> (wybor break-glass) -> Skanuj -> filtruj/wybierz/profil ->
    Wdroz zaznaczone -> Raport HTML. Dziala w sesji STA (Windows PowerShell 5.1 zalecane).
.NOTES
    Uruchom: powershell -STA -ExecutionPolicy Bypass -File .\Start-M365CISApp.ps1
#>
[CmdletBinding()]
param()

# --- Wymuszenie STA (WPF) ---
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "Przelaczam na sesje STA (wymagane przez WPF)..." -ForegroundColor Yellow
    $file = $PSCommandPath; if (-not $file) { $file = $MyInvocation.MyCommand.Path }
    Start-Process powershell -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$file`""
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Katalog aplikacji - dziala dla .ps1 i dla skompilowanego .exe (ps2exe)
$script:AppRoot = if ($PSScriptRoot) { $PSScriptRoot }
                  elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
                  else { Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) }

Import-Module (Join-Path $script:AppRoot 'M365CISCore.psm1') -Force -ErrorAction Stop

# --- Okno logowania (device code) ---
function Show-DeviceCodeDialog {
    param([string]$Url, [string]$Code)

    $dcXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Logowanie - M365 CIS Assistant" Width="520" Height="320"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="13" Background="#f4f6f9" Topmost="True" ShowInTaskbar="False">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="56"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="44"/>
    </Grid.RowDefinitions>

    <!-- Pasek tytulowy -->
    <Border Grid.Row="0" Background="#0b5cab">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0">
        <TextBlock Text="🔑" FontSize="20" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/>
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="Logowanie do Microsoft 365" Foreground="White" FontSize="14" FontWeight="Bold"/>
          <TextBlock Text="Uwierzytelnienie kodem urządzenia" Foreground="#c5dcf0" FontSize="11"/>
        </StackPanel>
      </StackPanel>
    </Border>

    <!-- Srodek -->
    <Grid Grid.Row="1" Margin="24,18,24,8">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="14"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Krok 1 -->
      <TextBlock Grid.Row="0" Foreground="#444" Margin="0,0,0,6">
        <Run FontWeight="SemiBold">1.</Run>
        <Run>Otwórz w przeglądarce:</Run>
      </TextBlock>
      <Border Grid.Row="1" Background="White" CornerRadius="6" Padding="12,9"
              BorderBrush="#b3d0ee" BorderThickness="1">
        <DockPanel>
          <Button x:Name="btnOpenUrl" DockPanel.Dock="Right" Content="Otwórz ↗"
                  Width="84" Height="30" Margin="10,0,0,0" BorderThickness="0"
                  Background="#0b5cab" Foreground="White" FontWeight="Bold" FontSize="12"/>
          <TextBlock x:Name="lblUrl" FontFamily="Consolas" FontSize="12" FontWeight="SemiBold"
                     Foreground="#0b5cab" VerticalAlignment="Center" TextWrapping="Wrap"/>
        </DockPanel>
      </Border>

      <!-- Krok 2 -->
      <TextBlock Grid.Row="3" Foreground="#444" Margin="0,0,0,6">
        <Run FontWeight="SemiBold">2.</Run>
        <Run>Wpisz ten kod:</Run>
      </TextBlock>
      <Border Grid.Row="4" Background="White" CornerRadius="6" Padding="14,10"
              BorderBrush="#86c995" BorderThickness="2">
        <DockPanel>
          <Button x:Name="btnCopy" DockPanel.Dock="Right" Content="Kopiuj"
                  Width="68" Height="36" Margin="12,0,0,0" BorderThickness="0"
                  Background="#1a7f37" Foreground="White" FontWeight="Bold" FontSize="12"/>
          <TextBlock x:Name="lblCode" FontFamily="Consolas" FontSize="32" FontWeight="Bold"
                     Foreground="#1a7f37" HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </DockPanel>
      </Border>
    </Grid>

    <!-- Pasek postępu -->
    <Border Grid.Row="2" Background="White" BorderBrush="#e0e0e0" BorderThickness="0,1,0,0">
      <DockPanel Margin="20,0">
        <ProgressBar DockPanel.Dock="Right" Width="120" Height="4" IsIndeterminate="True"
                     VerticalAlignment="Center" Margin="10,0,0,0" Foreground="#0b5cab" Background="#e0e0e0"/>
        <TextBlock Text="Oczekiwanie na logowanie… okno zamknie się automatycznie"
                   Foreground="#888" FontSize="11" VerticalAlignment="Center"/>
      </DockPanel>
    </Border>
  </Grid>
</Window>
'@

    [xml]$dcXml = $dcXaml
    $dcWin = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $dcXml))

    $dcWin.FindName('lblUrl').Text  = $Url
    $dcWin.FindName('lblCode').Text = $Code

    # Przyciski - przechowaj wartosci w Tag zeby uniknac problemow z domknieciem
    $btnOpen = $dcWin.FindName('btnOpenUrl')
    $btnOpen.Tag = $Url
    $btnOpen.Add_Click({ Start-Process $this.Tag })

    $btnCp = $dcWin.FindName('btnCopy')
    $btnCp.Tag = $Code -replace ' ',''
    $btnCp.Add_Click({
        [System.Windows.Clipboard]::SetText($this.Tag)
        $this.Content = '✓ Skopiowano'
        $this.Background = '#1557a0'
    })

    $dcWin.Show()

    # Zwroc callback zamykajacy okno (wywolywany z watku UI po zakonczeniu auth)
    return [scriptblock]::Create('$dcWin.Close()').GetNewClosure()
}

# --- Stan aplikacji ---
$script:AllRows    = New-Object System.Collections.ObjectModel.ObservableCollection[object]   # pelna lista
$script:View       = New-Object System.Collections.ObjectModel.ObservableCollection[object]   # widoczne (po filtrze)
$script:LastScan   = $null
$script:LastApplied= New-Object System.Collections.Generic.List[object]
$script:BgUser     = $null

# --- XAML ---
$xamlText = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 CIS Assistant" Height="780" Width="1180" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="12" Background="#f4f6f9">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="160"/>
    </Grid.RowDefinitions>

    <!-- Naglowek -->
    <Border Grid.Row="0" Background="#0b5cab" CornerRadius="6" Padding="12,8">
      <StackPanel Orientation="Horizontal">
        <TextBlock Text="M365 CIS Assistant" Foreground="White" FontSize="18" FontWeight="Bold" VerticalAlignment="Center"/>
        <TextBlock Text="  |  CIS Microsoft 365 Foundations v6.x" Foreground="#cfe2f5" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblTenant" Text="   Tenant: (niepolaczony)" Foreground="White" Margin="20,0,0,0" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblBg" Text="   Break-glass: (brak)" Foreground="#ffe08a" Margin="20,0,0,0" VerticalAlignment="Center"/>
      </StackPanel>
    </Border>

    <!-- Opcje polaczenia -->
    <Border Grid.Row="1" Background="White" CornerRadius="6" Padding="10" Margin="0,8,0,0" BorderBrush="#ddd" BorderThickness="1">
      <WrapPanel>
        <TextBlock Text="Conditional Access:" VerticalAlignment="Center" Margin="0,0,6,0"/>
        <ComboBox x:Name="cmbCaState" Width="130" Margin="0,0,16,0">
          <ComboBoxItem Content="ReportOnly" IsSelected="True"/>
          <ComboBoxItem Content="Enabled"/>
        </ComboBox>
        <CheckBox x:Name="chkSkipEntra"  Content="Skip Entra"  Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipExo"    Content="Skip Exchange" Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipSpo"    Content="Skip SharePoint" Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipTeams"  Content="Skip Teams" Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipIntune" Content="Skip Intune" Margin="0,0,16,0" VerticalAlignment="Center"/>
        <Button x:Name="btnConnect" Content="1. Polacz z tenantem" Width="170" Height="28" Background="#0b5cab" Foreground="White" FontWeight="Bold"/>
      </WrapPanel>
    </Border>

    <!-- Filtry + profile -->
    <Border Grid.Row="2" Background="White" CornerRadius="6" Padding="10" Margin="0,8,0,0" BorderBrush="#ddd" BorderThickness="1">
      <WrapPanel>
        <TextBlock Text="Poziom:" VerticalAlignment="Center" Margin="0,0,4,0"/>
        <ComboBox x:Name="cmbLevel" Width="110" Margin="0,0,12,0">
          <ComboBoxItem Content="Wszystkie" IsSelected="True"/>
          <ComboBoxItem Content="L1"/>
          <ComboBoxItem Content="L2"/>
        </ComboBox>
        <TextBlock Text="Status:" VerticalAlignment="Center" Margin="0,0,4,0"/>
        <ComboBox x:Name="cmbStatus" Width="130" Margin="0,0,12,0">
          <ComboBoxItem Content="Wszystkie" IsSelected="True"/>
          <ComboBoxItem Content="NIEZGODNE"/>
          <ComboBoxItem Content="Zgodne"/>
        </ComboBox>
        <TextBlock Text="Szukaj:" VerticalAlignment="Center" Margin="0,0,4,0"/>
        <TextBox x:Name="txtSearch" Width="160" Margin="0,0,16,0"/>
        <TextBlock Text="Profil:" VerticalAlignment="Center" Margin="0,0,4,0"/>
        <ComboBox x:Name="cmbProfile" Width="180" Margin="0,0,6,0"/>
        <Button x:Name="btnApplyProfile" Content="Zastosuj profil" Width="110" Margin="0,0,6,0"/>
        <Button x:Name="btnSaveProfile"  Content="Zapisz profil" Width="100" Margin="0,0,16,0"/>
        <Button x:Name="btnSelAll"  Content="Zaznacz widoczne"  Width="120" Margin="0,0,6,0"/>
        <Button x:Name="btnSelNone" Content="Odznacz widoczne" Width="120"/>
      </WrapPanel>
    </Border>

    <!-- Tabela kontrolek -->
    <DataGrid x:Name="grid" Grid.Row="3" Margin="0,8,0,0" AutoGenerateColumns="False" CanUserAddRows="False"
              HeadersVisibility="Column" GridLinesVisibility="Horizontal" RowHeaderWidth="0"
              SelectionMode="Single" Background="White" AlternatingRowBackground="#f4f7fb">
      <DataGrid.Columns>
        <DataGridCheckBoxColumn Header="Wdroz" Binding="{Binding Selected}" Width="55"/>
        <DataGridTextColumn Header="Status"  Binding="{Binding Status}"    Width="90"  IsReadOnly="True"/>
        <DataGridTextColumn Header="Poziom"  Binding="{Binding Poziom}"    Width="60"  IsReadOnly="True"/>
        <DataGridTextColumn Header="Obszar"  Binding="{Binding Obszar}"    Width="110" IsReadOnly="True"/>
        <DataGridTextColumn Header="Kontrolka" Binding="{Binding Kontrolka}" Width="*" IsReadOnly="True"/>
        <DataGridTextColumn Header="CIS"     Binding="{Binding CIS}"       Width="70"  IsReadOnly="True"/>
        <DataGridTextColumn Header="Stan zastany" Binding="{Binding Aktualnie}" Width="2*" IsReadOnly="True"/>
      </DataGrid.Columns>
    </DataGrid>

    <!-- Akcje -->
    <Border Grid.Row="4" Background="White" CornerRadius="6" Padding="10" Margin="0,8,0,0" BorderBrush="#ddd" BorderThickness="1">
      <WrapPanel>
        <CheckBox x:Name="chkWhatIf" Content="WhatIf (tylko symulacja)" IsChecked="True" Margin="0,0,16,0" VerticalAlignment="Center"/>
        <Button x:Name="btnScan"   Content="2. Skanuj tenant" Width="150" Height="30" Margin="0,0,8,0" Background="#0b5cab" Foreground="White" FontWeight="Bold"/>
        <Button x:Name="btnApply"  Content="3. Wdroz zaznaczone" Width="170" Height="30" Margin="0,0,8,0" Background="#1a7f37" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
        <Button x:Name="btnReport" Content="4. Raport HTML" Width="140" Height="30" Margin="0,0,16,0" IsEnabled="False"/>
        <TextBlock x:Name="lblStatus" Text="Gotowy." VerticalAlignment="Center" Foreground="#444"/>
      </WrapPanel>
    </Border>

    <!-- Log -->
    <Border Grid.Row="5" Background="#1e1e1e" CornerRadius="6" Margin="0,8,0,0">
      <TextBox x:Name="txtLog" Background="#1e1e1e" Foreground="#d4d4d4" FontFamily="Consolas" FontSize="11"
               IsReadOnly="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" BorderThickness="0" Padding="8"/>
    </Border>
  </Grid>
</Window>
'@

[xml]$xaml = $xamlText
$reader = New-Object System.Xml.XmlNodeReader $xaml
$win = [Windows.Markup.XamlReader]::Load($reader)

# --- Uchwyty kontrolek ---
$ctrl = @{}
'lblTenant','lblBg','cmbCaState','chkSkipEntra','chkSkipExo','chkSkipSpo','chkSkipTeams','chkSkipIntune',
'btnConnect','cmbLevel','cmbStatus','txtSearch','cmbProfile','btnApplyProfile','btnSaveProfile','btnSelAll',
'btnSelNone','grid','chkWhatIf','btnScan','btnApply','btnReport','lblStatus','txtLog' | ForEach-Object {
    $ctrl[$_] = $win.FindName($_)
}
$ctrl.grid.ItemsSource = $script:View

# --- Logowanie do okna ---
Set-CISLogCallback {
    param($Message,$Level)
    $line = "[{0}] [{1}] {2}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    $win.Dispatcher.Invoke([action]{ $ctrl.txtLog.AppendText($line); $ctrl.txtLog.ScrollToEnd() })
}
function Set-Status { param($Text) $ctrl.lblStatus.Text = $Text; $win.Dispatcher.Invoke([action]{}, 'Background') }
function Get-Combo { param($Combo) if ($Combo.SelectedItem) { $Combo.SelectedItem.Content } else { $null } }

# --- Filtr ---
function Test-RowVisible {
    param($r)
    $lvl = Get-Combo $ctrl.cmbLevel
    if ($lvl -eq 'L1' -and $r.Level -ne 1) { return $false }
    if ($lvl -eq 'L2' -and $r.Level -ne 2) { return $false }
    $st = Get-Combo $ctrl.cmbStatus
    if ($st -eq 'NIEZGODNE' -and $r.Status -ne 'NIEZGODNE') { return $false }
    if ($st -eq 'Zgodne'    -and $r.Status -ne 'Zgodne')    { return $false }
    $q = $ctrl.txtSearch.Text
    if ($q -and ($r.Kontrolka -notmatch [regex]::Escape($q)) -and ($r.Obszar -notmatch [regex]::Escape($q)) -and ($r.Id -notmatch [regex]::Escape($q))) { return $false }
    return $true
}
function Update-View {
    $script:View.Clear()
    foreach ($r in $script:AllRows) { if (Test-RowVisible $r) { $script:View.Add($r) } }
}

# --- Profile ---
function Load-ProfileList {
    $ctrl.cmbProfile.Items.Clear()
    $dir = Join-Path $script:AppRoot 'profiles'
    if (Test-Path $dir) {
        Get-ChildItem $dir -Filter *.json | ForEach-Object { $ctrl.cmbProfile.Items.Add($_.Name) | Out-Null }
    }
    if ($ctrl.cmbProfile.Items.Count -gt 0) { $ctrl.cmbProfile.SelectedIndex = 0 }
}

# --- Akcje ---
$ctrl.btnConnect.Add_Click({
    try {
        $win.Cursor='Wait'; Set-Status 'Laczenie...'
        $state = Get-Combo $ctrl.cmbCaState
        Connect-CISServices -SkipEntra:$ctrl.chkSkipEntra.IsChecked -SkipExchange:$ctrl.chkSkipExo.IsChecked `
            -SkipSharePoint:$ctrl.chkSkipSpo.IsChecked -SkipTeams:$ctrl.chkSkipTeams.IsChecked `
            -SkipIntune:$ctrl.chkSkipIntune.IsChecked -ConditionalAccessState $state | Out-Null
        $c = Get-CISContext
        $ctrl.lblTenant.Text = "   Tenant: " + $(if($c.TenantInitialDomain){$c.TenantInitialDomain}else{'(nieznany)'})
        # Wymuszony break-glass
        if ($c.Connected.Graph) {
            Set-Status 'Wybierz konto break-glass...'
            $users = Get-CISUsers
            $pick = $users | Out-GridView -Title 'Wybierz konto BREAK-GLASS (wykluczane z CA) i kliknij OK' -OutputMode Single
            if ($pick) {
                $script:BgUser = Get-MgUser -UserId $pick.Id -ErrorAction SilentlyContinue
                Set-CISBreakGlass -User $script:BgUser | Out-Null
                $ctrl.lblBg.Text = "   Break-glass: " + $script:BgUser.UserPrincipalName
            } else {
                [System.Windows.MessageBox]::Show('Nie wybrano konta break-glass. Polityki CA beda zablokowane do czasu wyboru.','Uwaga','OK','Warning') | Out-Null
            }
        }
        Set-Status 'Polaczono. Mozesz skanowac.'
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad polaczenia','OK','Error') | Out-Null
        Set-Status 'Blad polaczenia.'
    } finally { $win.Cursor='Arrow' }
})

$ctrl.btnScan.Add_Click({
    try {
        $win.Cursor='Wait'; Set-Status 'Skanowanie...'
        $script:LastScan = Invoke-CISScan
        $script:AllRows.Clear()
        foreach ($r in $script:LastScan) { $script:AllRows.Add($r) }
        # uzupelnij filtr statusu/obszaru? statusy stale; obszary dynamiczne pomijamy dla prostoty
        Update-View
        $ctrl.btnApply.IsEnabled = $true
        $ctrl.btnReport.IsEnabled = $true
        $n = @($script:AllRows | Where-Object Status -eq 'NIEZGODNE').Count
        Set-Status ("Skan zakonczony: {0} kontrolek, {1} niezgodnych." -f $script:AllRows.Count, $n)
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad skanu','OK','Error') | Out-Null
        Set-Status 'Blad skanu.'
    } finally { $win.Cursor='Arrow' }
})

$ctrl.btnApply.Add_Click({
    $ids = @($script:AllRows | Where-Object Selected | Select-Object -ExpandProperty Id)
    if ($ids.Count -eq 0) { [System.Windows.MessageBox]::Show('Nie zaznaczono zadnej kontrolki.','Info','OK','Information')|Out-Null; return }
    $whatif = [bool]$ctrl.chkWhatIf.IsChecked
    $mode = if ($whatif) { 'SYMULACJA (WhatIf)' } else { 'REALNE WDROZENIE' }
    $r = [System.Windows.MessageBox]::Show(("Tryb: {0}`nKontrolek: {1}`n`nKontynuowac?" -f $mode,$ids.Count),'Potwierdzenie','YesNo','Question')
    if ($r -ne 'Yes') { return }
    try {
        $win.Cursor='Wait'; Set-Status "$mode..."
        $script:LastApplied = Invoke-CISApply -Ids $ids -WhatIf:$whatif
        # odswiez skan po wdrozeniu
        if (-not $whatif) {
            $script:LastScan = Invoke-CISScan
            $script:AllRows.Clear(); foreach ($r2 in $script:LastScan) { $script:AllRows.Add($r2) }
            Update-View
        }
        $ok = @($script:LastApplied | Where-Object Status -eq 'APPLIED').Count
        $er = @($script:LastApplied | Where-Object Status -eq 'ERROR').Count
        Set-Status ("Zakonczono: {0} OK, {1} bledow." -f $ok,$er)
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad wdrozenia','OK','Error')|Out-Null
    } finally { $win.Cursor='Arrow' }
})

$ctrl.btnReport.Add_Click({
    if (-not $script:LastScan) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'HTML|*.html'; $dlg.FileName = "M365-CIS-Report-$(Get-Date -Format 'yyyyMMdd-HHmm').html"
    if ($dlg.ShowDialog()) {
        try {
            New-DeploymentReport -Scan $script:LastScan -Applied $script:LastApplied -Context (Get-CISContext) -Path $dlg.FileName | Out-Null
            Set-Status ("Raport zapisany: {0}" -f $dlg.FileName)
            try { Invoke-Item $dlg.FileName } catch { }
        } catch { [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad raportu','OK','Error')|Out-Null }
    }
})

$ctrl.btnApplyProfile.Add_Click({
    if ($script:AllRows.Count -eq 0) { return }
    $name = $ctrl.cmbProfile.SelectedItem
    if (-not $name) { return }
    try {
        $p = Import-CISProfile -Path (Join-Path (Join-Path $script:AppRoot 'profiles') $name)
        $ids = Get-CISProfileSelection -Profile $p -Scan $script:AllRows
        foreach ($r in $script:AllRows) { $r.Selected = ($ids -contains $r.Id) }
        Update-View   # przerysuj checkboxy
        Set-Status ("Profil '{0}': zaznaczono {1} kontrolek." -f $p.name, @($ids).Count)
    } catch { [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad profilu','OK','Error')|Out-Null }
})

$ctrl.btnSaveProfile.Add_Click({
    $ids = @($script:AllRows | Where-Object Selected | Select-Object -ExpandProperty Id)
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.InitialDirectory = (Join-Path $script:AppRoot 'profiles')
    $dlg.Filter='JSON|*.json'; $dlg.FileName='Custom-Profile.json'
    if ($dlg.ShowDialog()) {
        Save-CISProfile -Path $dlg.FileName -Name ([IO.Path]::GetFileNameWithoutExtension($dlg.FileName)) -Description 'Zapisany z GUI' -Ids $ids | Out-Null
        Load-ProfileList
        Set-Status ("Zapisano profil ({0} kontrolek)." -f $ids.Count)
    }
})

$ctrl.btnSelAll.Add_Click({  foreach ($r in $script:View) { $r.Selected=$true };  Update-View })
$ctrl.btnSelNone.Add_Click({ foreach ($r in $script:View) { $r.Selected=$false }; Update-View })
$ctrl.cmbLevel.Add_SelectionChanged({ Update-View })
$ctrl.cmbStatus.Add_SelectionChanged({ Update-View })
$ctrl.txtSearch.Add_TextChanged({ Update-View })

Set-CISDeviceCodeCallback {
    param($Url, $Code)
    Show-DeviceCodeDialog -Url $Url -Code $Code
}

Load-ProfileList
Write-CISLog 'GUI gotowe. Krok 1: Polacz z tenantem.' INFO
$win.ShowDialog() | Out-Null
Disconnect-CISServices
