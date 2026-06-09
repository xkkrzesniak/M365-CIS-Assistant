#Requires -Version 5.1
<#
.SYNOPSIS
    Buduje plik EXE z GUI (Start-M365CISApp.ps1) przy uzyciu modulu ps2exe.
.DESCRIPTION
    EXE jest cienkim launcherem PowerShell. Po zbudowaniu MUSI lezec w katalogu glownym repo,
    obok M365CISCore.psm1 oraz folderu profiles\ (sa wczytywane w czasie dzialania).
    EXE nie zawiera modulow M365 (Microsoft.Graph itd.) - musza byc zainstalowane w systemie.
.EXAMPLE
    .\build\Build-Exe.ps1
#>
param(
    [string]$Source = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Start-M365CISApp.ps1'),
    [string]$Output = (Join-Path (Split-Path $PSScriptRoot -Parent) 'M365-CIS-Assistant.exe'),
    [string]$IconPath
)

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "Instaluje modul ps2exe (CurrentUser)..." -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force -AllowClobber
}
Import-Module ps2exe -Force

$params = @{
    InputFile  = $Source
    OutputFile = $Output
    STA        = $true          # WPF wymaga STA
    NoConsole  = $true          # tylko okno GUI, bez konsoli
    Title      = 'M365 CIS Assistant'
    Description= 'CIS Microsoft 365 hardening assistant'
    Company    = 'YOUR ORG'
    Product    = 'M365 CIS Assistant'
    Version    = '0.4.0.0'
}
if ($IconPath -and (Test-Path $IconPath)) { $params.IconFile = $IconPath }

Write-Host "Buduje: $Output" -ForegroundColor Cyan
Invoke-ps2exe @params
Write-Host "Gotowe. Umiesc EXE obok M365CISCore.psm1 i folderu profiles\." -ForegroundColor Green
