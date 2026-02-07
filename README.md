# Sales Reporting Automation Tool

## Opis projektu
Narzędzie stanowi rozwiązanie opracowane w języku VBA, przeznaczone do automatyzacji procesów raportowania sprzedaży. Aplikacja umożliwia konsolidację rozproszonych danych z wielu plików tekstowych oraz ich automatyczne przekształcenie w interaktywny panel menedżerski. Projekt bazuje na danych sprzedażowych marek elektroniki użytkowej, takich jak Sony, Philips, Panasonic, Toshiba, LG oraz Samsung

## Kluczowe funkcjonalności
* **Zautomatyzowany import danych**: System pozwala na masowe wczytywanie plików tekstowych.
* **Zaawansowana analityka trendów**: Algorytm wylicza nachylenie linii trendu dla szeregów czasowych, dynamicznie modyfikując atrybuty wizualne wykresów w zależności od kierunku zmian sprzedaży
* **Interaktywny interfejs użytkownika**: 
    * Implementacja nawigacji za pomocą mapy geograficznej Polski, gdzie zdarzenia kliknięcia w kształty sterują filtrami tabel przestawnych
    * System dynamicznych kafelków KPI informujących o całkowitym wolumenie sprzedaży, liczbie transakcji oraz średniej wartości koszyka.
* **Zarządzanie bezpieczeństwem**: Automatyczne procedury blokowania arkuszy chronią strukturę raportu przed nieumyślną ingerencją użytkownika.

## Struktura techniczna projektu
Projekt został podzielony na moduły tematyczne w celu zachowania czytelności i łatwości utrzymania kodu:

* **data/**: Katalog zawierający surowe pliki źródłowe w formacie .txt.
* **src/Import_danych.bas**: Moduł odpowiedzialny za operacje na systemie plików, czyszczenie arkuszy oraz procedury zabezpieczające interfejs.
* **src/Filtry.bas**: Logika biznesowa odpowiedzialna za budowę silnika tabel przestawnych, konfigurację obiektów `SlicerCache` oraz obsługę interaktywnej mapy.
* **src/Wykresy.bas**: Procedury generujące obiekty `ChartObject`, formatowanie wizualizacji oraz przeliczanie wskaźników KPI.

## Instrukcja wdrożenia
1. Skopiować plik `Raport sprzedażowy.xlsm` oraz katalog `data` na dysk lokalny.
2. Otworzyć arkusz i uruchomić procedurę główną `ZbudujRaport`.
3. Wybrać pliki tekstowe do analizy z folderu źródłowego.
4. Po zakończeniu importu system automatycznie wygeneruje dashboard oraz umożliwi eksport wyników do formatu PDF.
