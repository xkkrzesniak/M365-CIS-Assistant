# Changelog

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
