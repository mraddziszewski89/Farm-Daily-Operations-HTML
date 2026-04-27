# Farm Daily Operations HTML

Automatyczny raport operacyjny dla farmy **SharePoint Subscription Edition on-premises**, generowany przez skrypt PowerShell i zapisywany jako interaktywny dashboard HTML.

Rozwiązanie zostało przygotowane z myślą o codziennej kontroli stanu farmy SharePoint. Skrypt zbiera dane z serwerów farmy, usług SharePoint, IIS, baz danych, logów systemowych oraz certyfikatów, a następnie agreguje je do jednego raportu HTML z czytelnym podsumowaniem problemów i sekcjami diagnostycznymi.

## Cel rozwiązania

Głównym celem projektu jest uproszczenie codziennego nadzoru administracyjnego nad farmą SharePoint poprzez:

- centralizację najważniejszych danych operacyjnych w jednym raporcie,
- szybkie wykrywanie problemów krytycznych i ostrzeżeń,
- ograniczenie czasu potrzebnego na ręczne sprawdzanie serwerów, usług, IIS i baz danych,
- umożliwienie automatyzacji raportowania oraz opcjonalnej wysyłki e-mail.

## Najważniejsze funkcje

- generowanie pojedynczego raportu HTML w formie dashboardu administracyjnego,
- globalna ocena stanu farmy: `OK`, `WARNING`, `CRITICAL`,
- sekcja zbiorcza z listą wykrytych problemów i rekomendacjami,
- analiza serwerów farmy i ról `MinRole`,
- kontrola usług SharePoint i `Service Applications`,
- analiza witryn IIS i `Application Pools`,
- przegląd aplikacji webowych SharePoint wraz z testem dostępności,
- analiza baz danych SharePoint, ich stanu, rozmiaru i flag `ReadOnly` / `NeedsUpgrade`,
- kontrola `Timer Jobs` i reguł `Health Analyzer`,
- przegląd wpisów z `Windows Event Log` i logów `ULS`,
- kontrola certyfikatów SSL/TLS powiązanych z IIS,
- analiza zasobów systemowych: dyski i pamięć RAM,
- weryfikacja spójności wersji buildów w całej farmie,
- opcjonalna archiwizacja raportów oraz wysyłka przez SMTP.

## Zakres raportu

Raport HTML zawiera następujące obszary:

1. Podsumowanie ogólne farmy
2. Serwery farmy
3. Service Applications
4. Usługi SharePoint
5. Witryny IIS
6. Application Pools IIS
7. Aplikacje webowe SharePoint
8. Bazy danych SharePoint
9. Timer Jobs
10. Health Analyzer
11. Windows Event Log
12. Logi ULS SharePoint
13. Certyfikaty SSL/TLS
14. Dyski serwerów
15. Pamięć RAM
16. Spójność wersji farmy

## Architektura rozwiązania

Rozwiązanie składa się z jednej głównej warstwy wykonawczej oraz jednej warstwy prezentacyjnej:

- `FarmDailyOperations.ps1` odpowiada za walidację środowiska, zbieranie danych, klasyfikację stanu oraz wygenerowanie raportu.
- raport wynikowy HTML jest budowany dynamicznie przez funkcję `Build-HTMLReport`, która osadza style CSS, logikę JavaScript i komplet sekcji danych.

Przepływ działania:

1. Walidacja uprawnień i inicjalizacja środowiska SharePoint.
2. Pobranie danych z lokalnego serwera oraz z pozostałych serwerów farmy.
3. Analiza stanu poszczególnych komponentów.
4. Agregacja wykrytych problemów do sekcji podsumowania.
5. Wygenerowanie raportu HTML.
6. Zapis raportu i logu wykonania.
7. Opcjonalna archiwizacja oraz wysyłka e-mail.

## Cechy raportu HTML

- ciemny, czytelny dashboard do pracy operacyjnej,
- boczna nawigacja po sekcjach,
- rozwijanie i zwijanie sekcji,
- filtrowanie po statusie i serwerze,
- globalne wyszukiwanie w tabelach,
- wyróżnianie problemów krytycznych i ostrzeżeń,
- podsumowanie liczby incydentów oraz zasobów.

## Wymagania

- Windows PowerShell `5.1+`,
- SharePoint Subscription Edition,
- uruchomienie z uprawnieniami lokalnego administratora,
- konto z rolą `Farm Administrator`,
- moduł `Microsoft.SharePoint.PowerShell`,
- moduł `WebAdministration`,
- dostęp sieciowy do pozostałych serwerów farmy, w tym `WinRM / PowerShell Remoting`.

## Przykładowe uruchomienie

```powershell
.\FarmDailyOperations.ps1
```

Przykład z archiwizacją i wysyłką e-mail:

```powershell
.\FarmDailyOperations.ps1 `
    -ReportOutputPath "D:\Reports\SharePoint" `
    -LogHours 24 `
    -DiskWarningThresholdGB 25 `
    -DiskCriticalThresholdGB 10 `
    -RAMWarningThreshold 80 `
    -RAMCriticalThreshold 90 `
    -ArchiveReports `
    -SendEmail `
    -SMTPServer "mail.contoso.com" `
    -SMTPPort 25 `
    -EmailFrom "sp-monitor@contoso.com" `
    -EmailTo @("admin@contoso.com","helpdesk@contoso.com")
```

## Najważniejsze parametry

- `ReportOutputPath` - katalog zapisu raportu HTML,
- `ReportFileName` - nazwa pliku raportu,
- `LogHours` - liczba godzin analizy dla `Event Log` i `ULS`,
- `DiskWarningThresholdGB` / `DiskCriticalThresholdGB` - progi dla wolnego miejsca na dysku,
- `RAMWarningThreshold` / `RAMCriticalThreshold` - progi użycia pamięci RAM,
- `ArchiveReports` - zapis dodatkowej kopii archiwalnej,
- `SendEmail` - włączenie wysyłki raportu e-mailem,
- `SMTPServer`, `SMTPPort`, `EmailFrom`, `EmailTo`, `UseSSL`, `SMTPCredential` - konfiguracja SMTP.

## Odporność i bezpieczeństwo

- każda sekcja zbierania danych działa w osobnym bloku `try/catch`,
- błąd pojedynczej sekcji nie przerywa wykonania całego raportu,
- brakujące dane są oznaczane w raporcie zamiast powodować awarię,
- skrypt zapisuje niezależny log wykonania,
- treści wstawiane do HTML są zabezpieczane przez escapowanie podstawowych znaków specjalnych.

## Zawartość repozytorium

- [`FarmDailyOperations.ps1`](./FarmDailyOperations.ps1) - główny skrypt raportujący,
- [`FarmDailyOperations_Dokumentacja_Techniczna.md`](./FarmDailyOperations_Dokumentacja_Techniczna.md) - dokumentacja techniczna do przeglądania bezpośrednio w GitHub,
- [`FarmDailyOperations_Dokumentacja_Techniczna.html`](./FarmDailyOperations_Dokumentacja_Techniczna.html) - oryginalna, rozbudowana wersja HTML dokumentacji,
- [`README.txt`](./README.txt) - wcześniejszy opis techniczny projektu,
- `README.md` - opis repozytorium do wyświetlania na GitHub.

## Ograniczenia

- rozwiązanie jest przeznaczone dla środowisk **SharePoint Subscription Edition on-premises**,
- część danych wymaga poprawnie skonfigurowanego dostępu zdalnego do serwerów farmy,
- skuteczność sekcji logów i IIS zależy od dostępności modułów oraz uprawnień wykonawczych,
- wysyłka SMTP opiera się na `Send-MailMessage`, więc w niektórych środowiskach może wymagać dostosowania do lokalnych standardów.

## Zastosowanie

Projekt może być używany jako:

- codzienny raport operacyjny dla administratora farmy,
- wsparcie dla zespołu utrzymaniowego i helpdesku,
- materiał diagnostyczny do analizy incydentów,
- podstawa do harmonogramowania zadania w `Task Scheduler`.

## Dokumentacja

Szczegółowy opis techniczny znajduje się w pliku:

- [`FarmDailyOperations_Dokumentacja_Techniczna.md`](https://github.com/mraddziszewski89/Farm-Daily-Operations-HTML/wiki#farmdailyoperationsps1---dokumentacja-techniczna)
