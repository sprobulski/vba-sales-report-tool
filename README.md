# Sales Reporting Automation Tool

## Opis projektu
Narzędzie stanowi kompleksowe rozwiązanie klasy ETL (Extract, Transform, Load) opracowane w języku VBA, przeznaczone do automatyzacji procesów raportowania sprzedaży. [cite_start]Aplikacja umożliwia konsolidację rozproszonych danych z wielu plików tekstowych oraz ich automatyczne przekształcenie w interaktywny panel menedżerski[cite: 113, 119]. [cite_start]Projekt bazuje na danych sprzedażowych marek elektroniki użytkowej, takich jak Sony, Philips, Panasonic, Toshiba, LG oraz Samsung[cite: 1, 2, 3, 4, 5, 6, 7, 8, 9].

## Kluczowe funkcjonalności
* [cite_start]**Zautomatyzowany import danych**: System pozwala na masowe wczytywanie plików tekstowych przy użyciu obiektów `FileDialog`, obsługując kodowanie UTF-8 dla zapewnienia integralności polskich znaków[cite: 121, 122, 123].
* **Przetwarzanie i standaryzacja**: Kod automatycznie parsuje złożone ciągi znaków, rozdzielając informacje o województwie i mieście, co umożliwia precyzyjną analizę regionalną[cite: 126, 127].
* [cite_start]**Zaawansowana analityka trendów**: Algorytm wylicza nachylenie linii trendu (funkcja `Slope`) dla szeregów czasowych, dynamicznie modyfikując atrybuty wizualne wykresów w zależności od kierunku zmian sprzedaży[cite: 141, 142, 144].
* **Interaktywny interfejs użytkownika**: 
    * [cite_start]Implementacja nawigacji za pomocą mapy geograficznej Polski, gdzie zdarzenia kliknięcia w kształty sterują filtrami tabel przestawnych[cite: 104, 106, 109].
    * System dynamicznych kafelków KPI informujących o całkowitym wolumenie sprzedaży, liczbie transakcji oraz średniej wartości koszyka[cite: 157, 158, 159].
* [cite_start]**Zarządzanie bezpieczeństwem**: Automatyczne procedury blokowania arkuszy chronią strukturę raportu przed nieumyślną ingerencją użytkownika[cite: 113, 118, 119].

## Struktura techniczna projektu
Projekt został podzielony na moduły tematyczne w celu zachowania czytelności i łatwości utrzymania kodu:

* [cite_start]**data/**: Katalog zawierający surowe pliki źródłowe w formacie .txt (raporty tygodniowe od 35 do 43)[cite: 1, 2, 3, 4, 5, 6, 7, 8, 9].
* **src/Import_danych.bas**: Moduł odpowiedzialny za operacje na systemie plików, czyszczenie arkuszy oraz procedury zabezpieczające interfejs[cite: 113, 122, 134].
* **src/Filtry.bas**: Logika biznesowa odpowiedzialna za budowę silnika tabel przestawnych, konfigurację obiektów `SlicerCache` oraz obsługę interaktywnej mapy[cite: 88, 92, 100].
* [cite_start]**src/Wykresy.bas**: Procedury generujące obiekty `ChartObject`, formatowanie wizualizacji oraz przeliczanie wskaźników KPI[cite: 137, 143, 157].

## Instrukcja wdrożenia
1. Skopiować plik `Raport sprzedażowy.xlsm` oraz katalog `data` na dysk lokalny.
2. [cite_start]Otworzyć arkusz i uruchomić procedurę główną `ZbudujRaport`[cite: 127].
3. [cite_start]Wybrać pliki tekstowe do analizy z folderu źródłowego[cite: 122].
4. [cite_start]Po zakończeniu importu system automatycznie wygeneruje dashboard oraz umożliwi eksport wyników do formatu PDF[cite: 129, 132].
