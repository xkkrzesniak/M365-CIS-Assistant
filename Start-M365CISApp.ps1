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
Add-Type -AssemblyName Microsoft.VisualBasic

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

# --- Okno wyboru / tworzenia konta Break-Glass ---
function Show-BGADialog {
    $bgaXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Konto Break-Glass" Width="680" Height="580"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        FontFamily="Segoe UI" FontSize="13" Background="#f4f6f9" ShowInTaskbar="False">
  <Grid Margin="0">
    <Grid.RowDefinitions>
      <RowDefinition Height="56"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="44"/>
    </Grid.RowDefinitions>

    <!-- Pasek tytulowy -->
    <Border Grid.Row="0" Background="#0b5cab">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="20,0">
        <TextBlock Text="🔓" FontSize="20" Foreground="White" VerticalAlignment="Center" Margin="0,0,10,0"/>
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="Konto Break-Glass" Foreground="White" FontSize="14" FontWeight="Bold"/>
          <TextBlock Text="Wybierz istniejącego Global Admina lub utwórz nowe konto" Foreground="#c5dcf0" FontSize="11"/>
        </StackPanel>
      </StackPanel>
    </Border>

    <!-- Lista Global Adminow -->
    <Grid Grid.Row="1" Margin="16,12,16,4">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <TextBlock Grid.Row="0" Text="Global Administratorzy tenanta:" FontWeight="SemiBold" Margin="0,0,0,6"/>
      <DataGrid x:Name="dgAdmins" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False"
                SelectionMode="Single" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                Background="White" AlternatingRowBackground="#f4f7fb" RowHeaderWidth="0" IsReadOnly="True">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Wyświetlana nazwa" Binding="{Binding DisplayName}" Width="2*"/>
          <DataGridTextColumn Header="UPN"               Binding="{Binding UserPrincipalName}" Width="3*"/>
          <DataGridTextColumn Header="Aktywne"           Binding="{Binding AccountEnabled}" Width="80"/>
        </DataGrid.Columns>
      </DataGrid>
      <Button x:Name="btnPickExisting" Grid.Row="2" Content="✓  Wybierz jako Break-Glass" Height="30"
              Margin="0,8,0,0" Background="#0b5cab" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
    </Grid>

    <!-- Separator -->
    <Border Grid.Row="2" BorderBrush="#ddd" BorderThickness="0,1,0,0" Margin="16,4,16,4">
      <TextBlock Text="— lub utwórz nowe konto BGA —" HorizontalAlignment="Center"
                 Foreground="#666" FontSize="12" Margin="0,6"/>
    </Border>

    <!-- Formularz tworzenia -->
    <Grid Grid.Row="3" Margin="16,0,16,4">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="180"/>
      </Grid.ColumnDefinitions>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <TextBlock Grid.Row="0" Grid.Column="0" Text="Wyświetlana nazwa:" Margin="0,0,0,3"/>
      <TextBlock Grid.Row="0" Grid.Column="2" Text="Login (UPN):" Margin="0,0,0,3"/>
      <TextBox x:Name="txtBgaName" Grid.Row="1" Grid.Column="0" Height="28" Padding="6,4"
               Text="Break Glass Account"/>
      <TextBox x:Name="txtBgaUpn"  Grid.Row="1" Grid.Column="2" Height="28" Padding="6,4"
               Text="breakglass@"/>
      <Button x:Name="btnCreateBga" Grid.Row="1" Grid.Column="4" Content="+ Utwórz konto BGA" Height="28"
              Background="#1a7f37" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
    </Grid>

    <!-- Stopka -->
    <Border Grid.Row="4" Background="White" BorderBrush="#ddd" BorderThickness="0,1,0,0">
      <DockPanel Margin="16,0">
        <Button x:Name="btnBgaCancel" DockPanel.Dock="Right" Content="Anuluj" Width="80" Height="28"
                Margin="8,0,0,0" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblBgaStatus" Foreground="#555" FontSize="11" VerticalAlignment="Center"
                   Text="Ładowanie listy Global Adminów…"/>
      </DockPanel>
    </Border>
  </Grid>
</Window>
'@

    [xml]$bgaXml = $bgaXaml
    $bgaWin = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $bgaXml))

    $dgAdmins      = $bgaWin.FindName('dgAdmins')
    $btnPick       = $bgaWin.FindName('btnPickExisting')
    $btnCreate     = $bgaWin.FindName('btnCreateBga')
    $btnCancel     = $bgaWin.FindName('btnBgaCancel')
    $txtName       = $bgaWin.FindName('txtBgaName')
    $txtUpn        = $bgaWin.FindName('txtBgaUpn')
    $lblStatus     = $bgaWin.FindName('lblBgaStatus')

    $script:BgaResult = $null

    # Zaladuj Global Adminow
    try {
        $admins = Get-CISGlobalAdmins
        $dgAdmins.ItemsSource = $admins
        $lblStatus.Text = "Znaleziono $(@($admins).Count) Global Administratorów."
    } catch {
        $lblStatus.Text = "Blad pobierania administratorow: $($_.Exception.Message)"
    }

    # Sugestia domeny UPN
    try {
        $ctx = Get-CISContext
        if ($ctx.TenantInitialDomain) { $txtUpn.Text = "breakglass@$($ctx.TenantInitialDomain)" }
    } catch { }

    $dgAdmins.Add_SelectionChanged({
        $btnPick.IsEnabled = ($dgAdmins.SelectedItem -ne $null)
    })

    $btnPick.Add_Click({
        $sel = $dgAdmins.SelectedItem
        if (-not $sel) { return }
        try {
            $script:BgaResult = Get-MgUser -UserId $sel.Id -ErrorAction Stop
        } catch {
            $script:BgaResult = $sel
        }
        $bgaWin.DialogResult = $true
    })

    $btnCreate.Add_Click({
        $upn  = $txtUpn.Text.Trim()
        $name = $txtName.Text.Trim()
        if (-not $upn -or $upn -eq 'breakglass@') {
            [System.Windows.MessageBox]::Show('Wpisz pełny adres UPN dla nowego konta.','Brak UPN','OK','Warning') | Out-Null
            return
        }
        try {
            $btnCreate.IsEnabled = $false
            $lblStatus.Text = "Tworzenie konta $upn…"
            $bgaWin.Dispatcher.Invoke([action]{}, 'Background')
            $res = New-CISBreakGlassAccount -UserPrincipalName $upn -DisplayName $name
            [System.Windows.MessageBox]::Show(
                "Konto utworzone!`n`nLogin:    $upn`nHasło:   $($res.Password)`n`nZapisz hasło teraz — nie będzie już dostępne.",
                'Konto BGA utworzone', 'OK', 'Information'
            ) | Out-Null
            $script:BgaResult = $res.User
            $bgaWin.DialogResult = $true
        } catch {
            $btnCreate.IsEnabled = $true
            $lblStatus.Text = "Blad tworzenia: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad tworzenia konta','OK','Error') | Out-Null
        }
    })

    $btnCancel.Add_Click({ $bgaWin.DialogResult = $false })

    $bgaWin.ShowDialog() | Out-Null
    return $script:BgaResult
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
      <RowDefinition Height="6"/>
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
        <CheckBox x:Name="chkSkipEntra"   Content="Skip Entra"       Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipExo"     Content="Skip Exchange"    Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipSpo"     Content="Skip SharePoint"  Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipTeams"   Content="Skip Teams"       Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipIntune"  Content="Skip Intune"      Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipPurview" Content="Skip Purview"     Margin="0,0,10,0" VerticalAlignment="Center"/>
        <CheckBox x:Name="chkSkipPP"      Content="Pomin Power Platform" Margin="0,0,16,0" VerticalAlignment="Center"/>
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
      <DataGrid.RowStyle>
        <Style TargetType="DataGridRow">
          <Style.Triggers>
            <DataTrigger Binding="{Binding Status}" Value="OK">
              <Setter Property="Background" Value="#e8f5e9"/>
            </DataTrigger>
            <DataTrigger Binding="{Binding Status}" Value="NIEZGODNE">
              <Setter Property="Background" Value="#ffebee"/>
            </DataTrigger>
            <DataTrigger Binding="{Binding Status}" Value="WARN">
              <Setter Property="Background" Value="#fff8e1"/>
            </DataTrigger>
          </Style.Triggers>
        </Style>
      </DataGrid.RowStyle>
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

    <!-- Panel szczegółów kontrolki -->
    <Border x:Name="pnlDetails" Grid.Row="4" Visibility="Collapsed" Background="#f0f6ff"
            BorderBrush="#0b5cab" BorderThickness="0,2,0,0" Padding="12,8" Margin="0,4,0,0">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="210"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Margin="0,0,12,0">
          <TextBlock x:Name="lblDetTitle" FontWeight="Bold" FontSize="13" Foreground="#0b5cab" TextWrapping="Wrap"/>
          <TextBlock x:Name="lblDetMeta"  Foreground="#666" FontSize="11" Margin="0,2,0,0"/>
          <TextBlock x:Name="lblDetDesc"  TextWrapping="Wrap" Margin="0,6,0,0" FontSize="12"/>
          <TextBlock x:Name="lblDetStan"  Foreground="#555" Margin="0,6,0,0" FontFamily="Consolas" FontSize="11" TextWrapping="Wrap"/>
        </StackPanel>
        <StackPanel Grid.Column="1" VerticalAlignment="Center">
          <Button x:Name="btnDetSelect"   Content="✓ Zaznacz do wdrozenia" Height="28" Margin="0,0,0,4"
                  Background="#1a7f37" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
          <Button x:Name="btnDetDeselect" Content="✗ Odznacz" Height="28"
                  Background="#c0392b" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- Pasek postępu -->
    <ProgressBar x:Name="prgScan" Grid.Row="5" Height="6" Minimum="0" Maximum="100"
                 Value="0" Visibility="Collapsed" Foreground="#0b5cab" Background="#dde"/>

    <!-- Akcje -->
    <Border Grid.Row="6" Background="White" CornerRadius="6" Padding="10" Margin="0,8,0,0" BorderBrush="#ddd" BorderThickness="1">
      <WrapPanel>
        <CheckBox x:Name="chkWhatIf" Content="WhatIf (tylko symulacja)" IsChecked="True" Margin="0,0,16,0" VerticalAlignment="Center"/>
        <Button x:Name="btnScan"   Content="2. Skanuj tenant" Width="150" Height="30" Margin="0,0,8,0" Background="#0b5cab" Foreground="White" FontWeight="Bold"/>
        <Button x:Name="btnApply"  Content="3. Wdroz zaznaczone" Width="170" Height="30" Margin="0,0,8,0" Background="#1a7f37" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
        <Button x:Name="btnReport"     Content="4. Raport HTML"     Width="130" Height="30" Margin="0,0,4,0" IsEnabled="False"/>
        <Button x:Name="btnExportCsv"  Content="Eksportuj CSV"      Width="110" Height="30" Margin="0,0,4,0" IsEnabled="False"/>
        <Button x:Name="btnExportWord" Content="Eksportuj Word/PDF"  Width="130" Height="30" Margin="0,0,4,0" IsEnabled="False"/>
        <Button x:Name="btnEmail"      Content="Wyslij e-mail"       Width="110" Height="30" Margin="0,0,16,0" IsEnabled="False"/>
        <TextBlock x:Name="lblStatus" Text="Gotowy." VerticalAlignment="Center" Foreground="#444"/>
      </WrapPanel>
    </Border>

    <!-- Log -->
    <Border Grid.Row="7" Background="#1e1e1e" CornerRadius="6" Margin="0,8,0,0">
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
'lblTenant','lblBg','cmbCaState','chkSkipEntra','chkSkipExo','chkSkipSpo','chkSkipTeams','chkSkipIntune','chkSkipPurview','chkSkipPP',
'btnConnect','cmbLevel','cmbStatus','txtSearch','cmbProfile','btnApplyProfile','btnSaveProfile','btnSelAll',
'btnSelNone','grid','chkWhatIf','btnScan','btnApply','btnReport','btnExportCsv','btnExportWord','btnEmail','lblStatus','txtLog',
'pnlDetails','prgScan','lblDetTitle','lblDetMeta','lblDetDesc','lblDetStan','btnDetSelect','btnDetDeselect' | ForEach-Object {
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
            -SkipIntune:$ctrl.chkSkipIntune.IsChecked -SkipPurview:$ctrl.chkSkipPurview.IsChecked `
            -SkipPowerPlatform:$ctrl.chkSkipPP.IsChecked `
            -ConditionalAccessState $state | Out-Null
        $c = Get-CISContext
        $ctrl.lblTenant.Text = "   Tenant: " + $(if($c.TenantInitialDomain){$c.TenantInitialDomain}else{'(nieznany)'})
        # Wymuszony break-glass
        if ($c.Connected.Graph) {
            Set-Status 'Wybierz konto break-glass...'
            $pick = Show-BGADialog
            if ($pick) {
                $script:BgUser = $pick
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
    # Przygotowanie
    $ctrl.btnScan.IsEnabled  = $false
    $ctrl.btnApply.IsEnabled = $false
    $ctrl.prgScan.Visibility = 'Visible'
    $ctrl.prgScan.Value      = 0
    $script:AllRows.Clear()
    $script:View.Clear()
    $ctrl.lblStatus.Foreground = '#444'
    Set-Status 'Skanowanie: pobieranie rejestru...'

    $registry = @(Get-CISControlRegistry | Where-Object { (Get-CISContext).Connected[$_.Service] })
    $total    = $registry.Count
    if ($total -eq 0) {
        Set-Status 'Brak kontrolek do skanowania. Sprawdz polaczenie.'
        $ctrl.btnScan.IsEnabled = $true; $ctrl.prgScan.Visibility = 'Collapsed'; return
    }

    $scanResults = [System.Collections.Generic.List[object]]::new()
    $scanIdx     = [ref]0

    $scanTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $scanTimer.Interval = [TimeSpan]::FromMilliseconds(5)
    $scanTimer.Add_Tick({
        $i = $scanIdx.Value
        if ($i -ge $total) {
            $scanTimer.Stop()
            $script:LastScan = $scanResults.ToArray()
            $ctrl.btnApply.IsEnabled  = $true
            $ctrl.btnReport.IsEnabled     = $true
            $ctrl.btnExportCsv.IsEnabled  = $true
            $ctrl.btnExportWord.IsEnabled = $true
            $ctrl.btnEmail.IsEnabled      = $true
            $ctrl.btnScan.IsEnabled   = $true
            $ctrl.prgScan.Visibility  = 'Collapsed'
            $ok   = @($scanResults | Where-Object Status -eq 'Zgodne').Count
            $nok  = @($scanResults | Where-Object Status -eq 'NIEZGODNE').Count
            $warn = @($scanResults | Where-Object Status -eq 'WARN').Count
            $pct  = if ($total -gt 0) { [int]([math]::Round($ok / $total * 100)) } else { 0 }
            $scoreColor = if ($pct -ge 80) { '#27ae60' } elseif ($pct -ge 50) { '#e67e22' } else { '#c0392b' }
            $ctrl.lblStatus.Foreground = $scoreColor
            Set-Status ("Zgodnosc CIS: {0}% ({1}/{2} OK) | Niezgodne: {3} | Ostrzezenia: {4}" -f $pct,$ok,$total,$nok,$warn)
            return
        }

        $c = $registry[$i]
        $status = 'Unknown'; $current = '-'
        try {
            $r = $c.Test.Invoke()
            $status  = if ($r.Compliant) { 'Zgodne' } else { 'NIEZGODNE' }
            $current = $r.Current
        } catch { $status = 'Blad'; $current = $_.Exception.Message.Substring(0, [Math]::Min(120, $_.Exception.Message.Length)) }

        $row = [pscustomobject]@{
            Selected  = ($status -eq 'NIEZGODNE')
            Id        = $c.Id
            Obszar    = $c.Area
            Kontrolka = $c.Name
            Status    = $status
            Poziom    = ("L{0}" -f $c.Level)
            Level     = $c.Level
            CIS       = $c.Cis
            Aktualnie = $current
        }
        $scanResults.Add($row)
        $script:AllRows.Add($row)
        if (Test-RowVisible $row) { $script:View.Add($row) }

        $pct2 = [int](($i + 1) / $total * 100)
        $ctrl.prgScan.Value = $pct2
        Set-Status ("[{0}/{1}] {2}" -f ($i + 1), $total, $c.Name)
        $logLvl = if ($status -eq 'Zgodne') { 'OK' } elseif ($status -eq 'NIEZGODNE') { 'WARN' } else { 'ERROR' }
        Write-CISLog ("{0,-12} L{1} {2}" -f $status, $c.Level, $c.Name) $logLvl

        $scanIdx.Value = $i + 1
    }.GetNewClosure())
    $scanTimer.Start()
})

$ctrl.btnApply.Add_Click({
    $ids = @($script:AllRows | Where-Object Selected | Select-Object -ExpandProperty Id)
    if ($ids.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Nie zaznaczono zadnej kontrolki.','Info','OK','Information') | Out-Null; return
    }
    $whatif = [bool]$ctrl.chkWhatIf.IsChecked
    $mode   = if ($whatif) { 'SYMULACJA (WhatIf)' } else { 'REALNE WDROZENIE' }
    $r      = [System.Windows.MessageBox]::Show(("Tryb: {0}`nKontrolek: {1}`n`nKontynuowac?" -f $mode,$ids.Count),'Potwierdzenie','YesNo','Question')
    if ($r -ne 'Yes') { return }

    $ctrl.btnApply.IsEnabled = $false
    $ctrl.btnScan.IsEnabled  = $false
    $ctrl.prgScan.Visibility = 'Visible'
    $ctrl.prgScan.Value      = 0
    Set-Status "$mode..."

    $registry     = Get-CISControlRegistry
    $toApply      = @($ids | ForEach-Object { $id = $_; $registry | Where-Object Id -eq $id })
    $applyTotal   = $toApply.Count
    $applyIdx     = [ref]0
    $applyResults = [System.Collections.Generic.List[object]]::new()

    $applyTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $applyTimer.Interval = [TimeSpan]::FromMilliseconds(5)
    $applyTimer.Add_Tick({
        $i = $applyIdx.Value
        if ($i -ge $applyTotal) {
            $applyTimer.Stop()
            $script:LastApplied = $applyResults.ToArray()

            if (-not $whatif) {
                # Reskan asynchroniczny po wdrozeniu
                $ctrl.btnScan.IsEnabled = $false
                $script:AllRows.Clear(); $script:View.Clear()
                $registry2 = @(Get-CISControlRegistry | Where-Object { (Get-CISContext).Connected[$_.Service] })
                $total2    = $registry2.Count
                $scanR2    = [System.Collections.Generic.List[object]]::new()
                $idx2      = [ref]0
                $rescanT   = [System.Windows.Threading.DispatcherTimer]::new()
                $rescanT.Interval = [TimeSpan]::FromMilliseconds(5)
                $rescanT.Add_Tick({
                    $j = $idx2.Value
                    if ($j -ge $total2) {
                        $rescanT.Stop()
                        $script:LastScan = $scanR2.ToArray()
                        $ctrl.btnScan.IsEnabled  = $true
                        $ctrl.btnApply.IsEnabled = $true
                        $ctrl.prgScan.Visibility = 'Collapsed'
                        $ok2  = @($scanR2 | Where-Object Status -eq 'Zgodne').Count
                        $nok2 = @($scanR2 | Where-Object Status -eq 'NIEZGODNE').Count
                        Set-Status ("Po wdrozeniu: {0}/{1} OK, {2} niezgodne." -f $ok2,$total2,$nok2)
                        return
                    }
                    $c2 = $registry2[$j]
                    $st2 = 'Unknown'; $cur2 = '-'
                    try { $res2=$c2.Test.Invoke(); $st2=if($res2.Compliant){'Zgodne'}else{'NIEZGODNE'}; $cur2=$res2.Current } catch { $st2='Blad'; $cur2=$_.Exception.Message }
                    $row2 = [pscustomobject]@{ Selected=($st2-eq'NIEZGODNE'); Id=$c2.Id; Obszar=$c2.Area; Kontrolka=$c2.Name; Status=$st2; Poziom=("L{0}"-f $c2.Level); Level=$c2.Level; CIS=$c2.Cis; Aktualnie=$cur2 }
                    $scanR2.Add($row2); $script:AllRows.Add($row2); if (Test-RowVisible $row2) { $script:View.Add($row2) }
                    $ctrl.prgScan.Value = [int](($j+1)/$total2*100)
                    Set-Status ("Reskan: [{0}/{1}] {2}" -f ($j+1),$total2,$c2.Name)
                    $idx2.Value = $j + 1
                }.GetNewClosure())
                $rescanT.Start()
            } else {
                $ctrl.btnApply.IsEnabled = $true
                $ctrl.btnScan.IsEnabled  = $true
                $ctrl.prgScan.Visibility = 'Collapsed'
            }
            $ok3 = @($applyResults | Where-Object Status -eq 'APPLIED').Count
            $er3 = @($applyResults | Where-Object Status -eq 'ERROR').Count
            Write-CISLog ("Wdrozenie zakonczone: {0} OK, {1} bledow." -f $ok3,$er3) OK
            return
        }

        $c = $toApply[$i]
        if ($whatif) {
            Write-CISLog ("WHATIF: {0}" -f $c.Name) SKIP
            $applyResults.Add([pscustomobject]@{ Id=$c.Id; Name=$c.Name; Status='WHATIF'; Detail='' })
        } else {
            try {
                & $c.Apply
                Write-CISLog ("WDROZONO: {0}" -f $c.Name) OK
                $applyResults.Add([pscustomobject]@{ Id=$c.Id; Name=$c.Name; Status='APPLIED'; Detail='' })
            } catch {
                Write-CISLog ("BLAD [{0}]: {1}" -f $c.Id,$_.Exception.Message) ERROR
                $applyResults.Add([pscustomobject]@{ Id=$c.Id; Name=$c.Name; Status='ERROR'; Detail=$_.Exception.Message })
            }
        }
        $ctrl.prgScan.Value = [int](($i+1)/$applyTotal*100)
        Set-Status ("{0}: [{1}/{2}] {3}" -f $mode, ($i+1), $applyTotal, $c.Name)
        $applyIdx.Value = $i + 1
    }.GetNewClosure())
    $applyTimer.Start()
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

$ctrl.btnExportCsv.Add_Click({
    if (-not $script:LastScan) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV|*.csv'; $dlg.FileName = "M365-CIS-Scan-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
    if ($dlg.ShowDialog()) {
        try {
            Export-CISScanToCsv -Scan $script:LastScan -Path $dlg.FileName
            Set-Status ("CSV zapisany: {0}" -f $dlg.FileName)
            try { Invoke-Item $dlg.FileName } catch { }
        } catch { [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad CSV','OK','Error')|Out-Null }
    }
})

$ctrl.btnExportWord.Add_Click({
    if (-not $script:LastScan) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Word (docx)|*.docx'; $dlg.FileName = "M365-CIS-Report-$(Get-Date -Format 'yyyyMMdd-HHmm').docx"
    if ($dlg.ShowDialog()) {
        try {
            $win.Cursor = 'Wait'; Set-Status 'Generowanie raportu Word...'
            Export-CISReportToWord -Scan $script:LastScan -Applied $script:LastApplied -Context (Get-CISContext) -Path $dlg.FileName -SaveAsPdf
            Set-Status ("Raport Word+PDF: {0}" -f $dlg.FileName)
            try { Invoke-Item $dlg.FileName } catch { }
        } catch { [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad Word','OK','Error')|Out-Null } finally { $win.Cursor = 'Arrow' }
    }
})

$ctrl.btnEmail.Add_Click({
    if (-not $script:LastScan) { return }
    $to = [Microsoft.VisualBasic.Interaction]::InputBox('Podaj adres e-mail odbiorcy raportu:', 'Wyslij raport e-mail', '')
    if (-not $to -or $to -notmatch '@') { return }
    try {
        $win.Cursor = 'Wait'; Set-Status 'Wysylanie e-mail...'
        $htmlPath = $null
        Send-CISReportByEmail -To $to -Scan $script:LastScan -Applied $script:LastApplied -Context (Get-CISContext) -HtmlReportPath $htmlPath
        Set-Status "E-mail wyslany do: $to"
    } catch { [System.Windows.MessageBox]::Show($_.Exception.Message,'Blad e-mail','OK','Error')|Out-Null } finally { $win.Cursor = 'Arrow' }
})

$ctrl.btnApplyProfile.Add_Click({
    if ($script:AllRows.Count -eq 0) { return }
    $name = $ctrl.cmbProfile.SelectedItem
    if (-not $name) { return }
    try {
        $p = Import-CISProfile -Path (Join-Path (Join-Path $script:AppRoot 'profiles') $name)
        $ids = Get-CISProfileSelection -CisProfile $p -Scan $script:AllRows
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

# --- Panel szczegółów ---
$ctrl.grid.Add_SelectionChanged({
    $item = $ctrl.grid.SelectedItem
    if (-not $item) { $ctrl.pnlDetails.Visibility = 'Collapsed'; return }
    $ctrl.pnlDetails.Visibility = 'Visible'
    $ctrl.lblDetTitle.Text = "$($item.Id)  —  $($item.Kontrolka)"
    $ctrl.lblDetMeta.Text  = "CIS $($item.CIS)  |  $($item.Poziom)  |  Obszar: $($item.Obszar)  |  Status: $($item.Status)"
    $desc = (Get-CISControlDocs)[$item.Id]; if (-not $desc) { $desc = '(brak opisu — dodaj do $script:ControlDocs w M365CISCore.psm1)' }
    $ctrl.lblDetDesc.Text  = $desc
    $ctrl.lblDetStan.Text  = "Stan zastany: $($item.Aktualnie)"
})

$ctrl.btnDetSelect.Add_Click({
    $item = $ctrl.grid.SelectedItem
    if ($item) { $item.Selected = $true; $ctrl.grid.Items.Refresh() }
})

$ctrl.btnDetDeselect.Add_Click({
    $item = $ctrl.grid.SelectedItem
    if ($item) { $item.Selected = $false; $ctrl.grid.Items.Refresh() }
})

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
