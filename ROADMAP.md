# Roadmap

## Zrobione
- Silnik (modul) + CLI + GUI WPF na wspolnym rejestrze kontrolek.
- Skan -> wybor (checkboxy, filtry, profile) -> wdrozenie -> raport HTML.
- Profile JSON + zapis z GUI. Builder EXE (ps2exe).
- Tryb Report-only dla CA, wymuszony break-glass, auto-domena, dokumentacja powdrozeniowa.

## Planowane
- Asynchroniczny skan/wdrozenie (runspace + pasek postepu) - GUI bez zamrazania.
- Podglad/edycja wartosci kontrolki przed wdrozeniem (panel szczegolow).
- Eksport/import calej konfiguracji jako kod (JSON) + tryb "diff" (stan zastany vs docelowy).
- Uwierzytelnianie app-only (certyfikat / app registration) dla CI i uruchomien nieinteraktywnych.
- Rozbudowa rejestru (Purview/DLP, Power Platform, dodatkowe Intune, Copilot wg CIS v6).
- Testy Pester + walidacja w GitHub Actions; podpisywanie skryptu (Authenticode).
- Lokalizacja (PL/EN), motyw ciemny GUI.
