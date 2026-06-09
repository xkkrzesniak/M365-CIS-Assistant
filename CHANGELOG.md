# Changelog

Wszystkie istotne zmiany w projekcie M365 CIS Assistant.
Format: [Semantic Versioning](https://semver.org/).

---

## [Unreleased]
### Planowane
- Eksport wyników skanu do CSV (otwieralny w Excelu)
- Wysyłka raportu HTML przez Microsoft Graph (`Send-MgUserMail`)
- Porównanie skanów before/after z wykresem trendu
- Kolorowanie wierszy DataGrid (OK=zielony, NIEZGODNE=czerwony)
- Zaplanowane skany + Task Scheduler
- Kontrolki: `MDO-ZAP`, `EXO-FORWARDINGRULES`, `TEAMS-PRESENTER`, `INTUNE-ENCRYPT`, `ENTRA-SMARTLOCKOUT`

---

## [0.9.0] — 2026-06-09
### Dodano
- **8 nowych kontrolek:**
  - `EXO-SMTPAUTH` (6.5.1): wyłącz SMTP AUTH globalnie
  - `EXO-CALENDAR` (6.3): kalendarze zewnętrzne — tylko FreeBusySimple
  - `EXO-MAILTIPS` (6.x): MailTips dla zewnętrznych odbiorców
  - `SPO-UNMANAGED` (7.1): urządzenia niezarządzane → tylko podgląd
  - `SPO-GUESTEXPIRY` (7.4): goście SPO wygasają po 60 dniach
  - `TEAMS-EXTCONTROL` (8.4): blokada przejęcia kontroli nad ekranem
  - `TEAMS-MEETINGCHAT` (8.1): anonimowi bez dostępu do czatu
  - `ENTRA-BREAKGLASS` (1.4): audyt MFA konta break-glass (read-only)
- **Wynik zgodności w GUI** — po skanie: `Zgodnosc CIS: 72% (18/25 OK) | Niezgodne: 5 | Ostrzezenia: 2`, kolor: zielony ≥80% / pomarańczowy ≥50% / czerwony <50%

## [0.8.0] — 2026-06-09
### Dodano
- **9 nowych kontrolek:**
  - `ENTRA-M365GROUP` (1.1.8): blokada tworzenia grup M365 przez userów
  - `MDO-MALWARE` (2.1.7): filtr niebezpiecznych typów plików
  - `MDO-ANTISPAM-IN` (2.1.5): HC spam i phishing → kwarantanna
  - `MDO-ANTISPAM-OUT` (2.1.9): konto spamujące → blokada
  - `TEAMS-EXTERNAL` (8.2.1): blokada Skype/publicznych kont
  - `TEAMS-GUESTCALL` (8.3): goście nie mogą dzwonić prywatnie
  - `TEAMS-RECORDING` (8.5 L2): wyłącz nagrywanie spotkań

## [0.7.0] — 2026-06-09
### Dodano
- **6 kontrolek ustawień użytkowników Entra ID (sekcja CIS 1.1):**
  - `ENTRA-APPREG` (1.1.3), `ENTRA-TENANT-CREATE` (1.1.4), `ENTRA-SECGROUP` (1.1.6)
  - `ENTRA-GUEST-PERMS` (1.1.7), `ENTRA-GUEST-INVITE` (1.1.5), `ENTRA-PORTAL` (1.1.2)

## [0.6.0] — 2026-06-09
### Naprawiono
- **Auth Graph**: MSAL assembly `Microsoft.Identity.Client.dll` ładowany jawnie z katalogu modułu Graph gdy lazy-load nie wpisuje go do AppDomain (PS 5.1)
- **Auth EXO/Purview**: zastąpiono device code (`-Device`) przez `Connect-ExchangeOnline -UserPrincipalName` z SSO — eliminuje błąd "parametr Device nie istnieje" w EXO v3.10+
- **Auth SPO**: przywrócono `Microsoft.Online.SharePoint.PowerShell` (PnP.PowerShell 3.x wymaga PS 7+); błędy auth widoczne zamiast cichego pomijania
- **MSAL refaktor**: `Initialize-CISMsal` + `Get-CISMsalToken` jako współdzielone helpery; `WithRedirectUri('http://localhost')` + `WithUseEmbeddedWebView($false)`

## [0.5.0] — 2026-06-09
### Dodano
- Autoryzacja Graph przez MSAL.NET `PublicClientApplicationBuilder` z interaktywną przeglądarką (zastępuje WAM broker i device code dla Graph)
- `DispatcherFrame` do nieblokującego czekania na token w UI WPF STA

## [0.4.0] - 2026-06-09
### Added
- Pelne **GUI WPF** (`Start-M365CISApp.ps1`): DataGrid z checkboxami, filtry poziomu/statusu/szukaj, profile, log na zywo, przyciski Polacz/Skanuj/Wdroz/Raport.
- **Silnik jako modul** `M365CISCore.psm1` - jedno zrodlo prawdy dla CLI i GUI.
- **Profile / zestawy kontrolek** (JSON): Baseline L1, Strict L1+L2, Identity Core, Email Security, All. Zapis profilu z GUI.
- **Builder EXE** (`build/Build-Exe.ps1`, ps2exe), odporne ustalanie katalogu aplikacji.
### Changed
- CLI przepisany na cienka nakladke nad modulem; dodane `-Profile` i `-Unattended`.

## [0.3.0] - 2026-06-09
### Added
- Dokumentacja powdrozeniowa HTML, wymuszony wybor break-glass po logowaniu, auto-domena z Graph.

## [0.2.0] - 2026-06-09
### Added
- Filtr L1/L2, blok Intune, DKIM/DMARC/SPF, nazwy CAP01-CAP04, sprzatanie starych polityk.

## [0.1.0] - 2026-06-09
### Added
- Architektura rejestru kontrolek: skan + okno wyboru + wdrozenie (Entra/EXO/Defender/SPO/Teams).
