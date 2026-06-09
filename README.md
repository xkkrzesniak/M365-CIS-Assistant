# M365 CIS Assistant

Aplikacja do hardeningu tenanta **Microsoft 365** wedlug **CIS Microsoft 365 Foundations Benchmark v6.x** (poziomy L1/L2), dla licencji **Microsoft 365 Business Premium** (Entra ID P1 + Defender for Office 365 P1 + Intune).

Skanuje tenant, pokazuje wynik w **GUI (WPF)** lub w oknie wyboru CLI, wdraza **tylko zaznaczone** kontrolki (z profilami/zestawami), a na koncu generuje **dokumentacje powdrozeniowa HTML**.

> ⚠️ Narzedzie wprowadza zmiany w tenancie. Testuj najpierw na tenancie testowym; pierwszy raz uruchamiaj w trybie WhatIf. Uzywasz na wlasna odpowiedzialnosc.

## Architektura

| Plik | Rola |
|---|---|
| `M365CISCore.psm1` | **Silnik** - jedno zrodlo prawdy: rejestr kontrolek, polaczenia, skan, wdrozenie, profile, raport. |
| `Start-M365CISApp.ps1` | **GUI (WPF)** - graficzny interfejs (DataGrid, filtry, profile, log na zywo). |
| `Invoke-M365-CIS-Assistant.ps1` | **CLI** - cienka nakladka do uruchomien konsolowych / automatyzacji. |
| `profiles/*.json` | **Profile / zestawy kontrolek** (Baseline L1, Strict L1+L2, Identity Core, Email Security, All). |
| `build/Build-Exe.ps1` | Builder pliku **.exe** (ps2exe). |
| `profiles/*.json` | Gotowe zestawy kontrolek (Baseline L1, Strict L1+L2, Identity Core, Email Security, All). |

GUI i CLI korzystaja z tego samego silnika - kontrolki definiuje sie raz w `M365CISCore.psm1`.

## Wymagania

- **Windows PowerShell 5.1** (zalecane; GUI WPF wymaga sesji STA).
- Rola **Global Administrator**.
- Moduly (instalowane automatycznie w razie braku): `Microsoft.Graph`, `ExchangeOnlineManagement`, `Microsoft.Online.SharePoint.PowerShell`, `MicrosoftTeams`.

## Uruchomienie - GUI

```powershell
powershell -STA -ExecutionPolicy Bypass -File .\Start-M365CISApp.ps1
```

Przeplyw w oknie: **1. Polacz** (auto-domena + wymuszony wybor konta break-glass) -> **2. Skanuj** -> filtruj / zaznacz / zastosuj profil -> **3. Wdroz zaznaczone** (z opcja WhatIf) -> **4. Raport HTML**.

## Uruchomienie - CLI

```powershell
# Podglad (nic nie zapisuje):
.\Invoke-M365-CIS-Assistant.ps1 -WhatIf

# Interaktywnie (okno wyboru / konsola):
.\Invoke-M365-CIS-Assistant.ps1

# Z profilem, bez interakcji:
.\Invoke-M365-CIS-Assistant.ps1 -Profile .\profiles\Baseline-L1.json -Unattended

# Sam audyt zgodnosci:
.\Invoke-M365-CIS-Assistant.ps1 -ScanOnly

# Polityki CA na ostro (po weryfikacji raportow logowan):
.\Invoke-M365-CIS-Assistant.ps1 -ConditionalAccessState Enabled
```

## Budowa EXE

```powershell
.\build\Build-Exe.ps1
```

Powstaje `M365-CIS-Assistant.exe` w katalogu glownym. **EXE musi lezec obok** `M365CISCore.psm1` i folderu `profiles\` (sa wczytywane w runtime). EXE to launcher - moduly M365 musza byc zainstalowane w systemie.

## Profile (zestawy kontrolek)

Format JSON:

```json
{
  "name": "Baseline L1",
  "description": "Opis",
  "select": { "ids": [], "levels": [1], "areas": [], "excludeIds": [] }
}
```

Regula doboru: kontrolka trafia do zestawu, gdy pasuje do `ids` **lub** `levels` **lub** `areas` (pusty selektor = wszystkie), minus `excludeIds`. Profile mozna zapisywac z GUI ("Zapisz profil").

## Polityki Conditional Access

`CAP01: Block Legacy Authentication`, `CAP02: Require MFA for All Users`, `CAP03: Require MFA for Admin Roles`, `CAP04: Require Compliant or Hybrid Joined Device`. Domyslnie tworzone w trybie **Report-only**; konto break-glass zawsze wykluczone.

## Dodawanie kontrolek

W `M365CISCore.psm1` (sekcja `Get-CISControlRegistry`) dopisz obiekt z polami `Id, Service, Area, Cis, Level, Name, Test, Apply`. Skan, GUI, profile i raport obsluza go automatycznie. Opis do dokumentacji dodaj w `$script:ControlDocs`.

## Roadmap / licencja

[ROADMAP.md](ROADMAP.md) · [MIT](LICENSE)
