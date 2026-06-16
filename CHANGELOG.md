# Changelog

Wszystkie istotne zmiany w projekcie M365 CIS Assistant.
Format: [Semantic Versioning](https://semver.org/).

---

## [Unreleased]
### Planowane
- Porównanie skanów before/after z wykresem trendu
- Zaplanowane skany + Task Scheduler
- Integracja z Maester (opcjonalny backend skanu)
- Kontrolki: `TEAMS-PRESENTER`, `ENTRA-SMARTLOCKOUT`

---

## [1.2.1] — 2026-06-16
### Naprawiono
- **Bug: zielone tło wierszy w DataGrid nigdy nie odpala** — `DataTrigger Value="OK"` zmienione na `Value="Zgodne"` (rzeczywista wartość statusu w `Start-M365CISApp.ps1`).
- **Bug: `Export-CISReportToWord` pokazuje zawsze 0% zgodności** — `Where-Object Status -eq 'OK'` oraz `switch 'OK'` zmienione na `'Zgodne'` w `M365CISCore.psm1`.
- **Bug: `Send-CISReportByEmail` pokazuje zawsze 0 w temacie maila** — ta sama poprawka `'OK'` → `'Zgodne'`.

---

## [1.2.0] — 2026-06-10
### Dodano
- **Kreator Start Tenant** (przycisk `★ Kreator Start Tenant`): 5-krokowy kreator WPF dla nowych tenantów — grupuje wyniki skanu wg obszaru (Entra ID → Exchange/MDO → SharePoint → Teams → Podsumowanie), pokazuje szczegółowe opisy z ControlDocs, checkboxy do wyboru poprawek, jedno kliknięcie `Wdroz zaznaczone` aplikuje wszystkie zaznaczone kontrolki jednocześnie. Pre-zaznaczone wszystkie NIEZGODNE.
- **15 nowych kontrolek** (zainspirowanych Maester/ORCA/EIDSCA):
  - `ENTRA-GA-COUNT` (CIS 1.1.1 L1): liczba Global Adminów w zakresie 2-4
  - `ENTRA-CA-RISK-USER` (CIS 1.2.4 L2): CA blokuje użytkowników wysokiego ryzyka (Identity Protection)
  - `ENTRA-CA-RISK-SIGNIN` (CIS 1.2.3 L2): CA wymaga MFA dla ryzykownych logowań
  - `ENTRA-AUTHFIDO2` (L1): metoda FIDO2/Passkeys włączona w Authentication Methods Policy
  - `MDO-PHISH-THRESHOLD` (ORCA 101 / CIS 2.1.4 L1): PhishThresholdLevel ≥3 (Aggressive)
  - `MDO-DMARC-POLICY` (CIS 2.1.x L1): HonorDmarcPolicy=true — respektowanie DMARC p=reject
  - `MDO-BULK-THRESHOLD` (ORCA / CIS 2.1.5 L1): BCL ≤6 w polityce antyspam
  - `MDO-QUARANTINE-NOTIFY` (ORCA L1): ESN — powiadomienia kwarantanny dla użytkowników
  - `MDO-MAILBOX-INTEL` (ORCA / CIS 2.1.4 L1): Mailbox Intelligence + IntelligenceProtection=Quarantine
  - `EXO-BYPASS-RULES` (CIS 6.x L1): audyt reguł transportu z SCL=-1 omijających spam
  - `EXO-CONNECTORS-TLS` (CIS 6.x L1): aktywne connectory przychodzące wymagają TLS
  - `SPO-DEFAULT-LINK` (CIS 7.2.3 L1): DefaultSharingLinkType=Direct (tylko konkretne osoby)
  - `SPO-LINK-PERMISSION` (CIS 7.2.4 L1): DefaultLinkPermission=View
  - `TEAMS-LOBBY` (CIS 8.1.1 L1): AutoAdmittedUsers=EveryoneInCompanyExcludingGuests
  - `TEAMS-PSTN-LOBBY` (CIS 8.1.2 L1): AllowPSTNUsersToBypassLobby=false
- Łącznie: **78 kontrolek** (było 63)
- Przycisk `btnWizard` aktywowany po każdym skanie (zarówno głównym jak i po reskan po wdrożeniu)

---

## [1.1.0] — 2026-06-09
### Dodano
- **11 nowych kontrolek:**
  - `PP-ENVCONFIG` (9.1.1 L1): Power Platform — tworzenie środowisk tylko przez adminów
  - `PP-TRIALENV` (9.1.2 L1): Power Platform — blokada trial environments
  - `PP-SHAREWITHTENANT` (9.1.3 L2): Power Platform — blokada udostępniania canvas app całemu tenantowi
  - `COPILOT-PLUGINS` (1.3.1 L1): Copilot M365 — zarządzanie wtyczkami zewnętrznych wydawców
  - `COPILOT-M365GROUPS` (1.3.2 L1): Copilot M365 — audyt przypisania licencji
  - `INTUNE-ENCRYPT-WIN` (5.2.1 L1): Intune — BitLocker wymagany w polityce Windows
  - `INTUNE-AV-WIN` (5.2.2 L1): Intune — antywirus wymagany (Windows)
  - `INTUNE-FIREWALL-WIN` (5.2.3 L1): Intune — zapora sieciowa wymagana (Windows)
  - `INTUNE-JAILBREAK` (5.1.1 L1): Intune — blokada urządzeń jailbreak/root (iOS/Android)
  - `PUR-DLP-TEAMS` (3.3.3 L1): Purview DLP — polityka obejmująca Teams
  - `PUR-INSIDER-RISK` (3.5 L2): Purview — Insider Risk Management
- **Asynchroniczny skan** — DispatcherTimer procesuje jedną kontrolkę na tick (5ms); UI nie zamraża; pasek postępu `prgScan` w czasie rzeczywistym; DataGrid aktualizowany live; po wdrożeniu automatyczny reskan async
- **Panel szczegółów** — widoczny po kliknięciu wiersza: pełna nazwa, CIS ref, poziom, opis z `$script:ControlDocs`, stan zastany; przyciski „Zaznacz do wdrożenia" / „Odznacz"
- **Power Platform connection** — `Connect-CISServices -SkipPowerPlatform`; nowy checkbox w GUI; `$script:Ctx.Connected.PowerPlatform` inicjalizowany w `Reset-CISContext`

---

## [1.0.0] — 2026-06-09
### Dodano
- **Kolorowanie wierszy DataGrid** — wiersze tabeli kontrolek kolorowane wg statusu: OK=jasna zieleń (`#e8f5e9`), NIEZGODNE=jasna czerwień (`#ffebee`), WARN=jasny pomarańcz (`#fff8e1`); zaznaczenie (IsSelected) nadpisuje kolory.
- **2 nowe kontrolki:**
  - `MDO-ZAP` (CIS 2.1.6, L1): Zero-hour Auto Purge — weryfikuje i włącza ZAP dla spam, phishing i malware w politykach EXO.
  - `EXO-FORWARDINGRULES` (CIS 6.2.2, L1): audyt zewnętrznych reguł przekazywania poczty (ForwardingSmtpAddress + Inbox Rules); tylko odczyt, wymaga ręcznego przeglądu.
- **Eksport CSV** (`Export-CISScanToCsv`) — eksportuje wyniki skanu do pliku CSV (separator `;`, UTF-8) z kolumnami ID/Obszar/CIS/Poziom/Kontrolka/Status/Stan/Opis; przycisk „Eksportuj CSV" w GUI.
- **Raport Word/PDF** (`Export-CISReportToWord`) — generuje raport DOCX (+ opcjonalnie PDF) przez COM Word.Application: strona tytułowa z wynikiem %, podsumowanie wykonawcze, tabela wyników z kolorami statusów i opisami kontrolek; przycisk „Eksportuj Word/PDF" w GUI.
- **Wysyłka e-mail przez Graph** (`Send-CISReportByEmail`) — wysyła raport (inline HTML lub wygenerowany) przez Microsoft Graph `sendMail`; przycisk „Wyślij e-mail" z InputBox do podania adresata.
- **GUI**: dodano przyciski `btnExportCsv`, `btnExportWord`, `btnEmail` aktywowane po skanie; załadowanie `Microsoft.VisualBasic` dla InputBox.

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
