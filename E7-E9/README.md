# TI.1 – Excel i Word: zadania E7, E8 i E9 (kompletny przewodnik)

Ten folder zawiera komplet materiałów do części III laboratoriów z przedmiotu **Technologie Informacyjne** (spółka **Hog-Dog sp. z o.o.**).

---

## 1. Zawartość folderu `E7-E9`

| Plik | Opis |
| :--- | :--- |
| [`TI.1_Excel.E7+E8+E9_InstrukcjaLab.pdf`](./TI.1_Excel.E7+E8+E9_InstrukcjaLab.pdf) | Oryginalna treść zadań laboratoryjnych E7, E8 i E9. |
| [`TI.1_Excel.E7+E8+E9_PlikRoboczy_rozwiazanie.xlsx`](./TI.1_Excel.E7+E8+E9_PlikRoboczy_rozwiazanie.xlsx) | Ukończony skoroszyt Excela: 10 arkuszy, 2 tabele (zielony styl `TableStyleMedium11`), 7 tabel przestawnych (`PivotStyleMedium4`) i 2 wykresy przestawne. |
| [`TI.1_Word.E9_Raport_HogDog.docx`](./TI.1_Word.E9_Raport_HogDog.docx) | Czysty raport dla zarządu firmy Hog-Dog sp. z o.o. (tekst podsumowujący, dwie tabele i dwa wykresy przestawne). Plik działa całkowicie niezależnie od zewnętrznego pliku Excela. |
| [`README.md`](./README.md) | Ten przewodnik: teoria, instrukcja krok po kroku do każdego podpunktu, liczby kontrolne i gotowe odpowiedzi na pytania z PDF-a. |

---

## 2. Mapa arkuszy i liczby kontrolne w pliku Excela

| Nr | Nazwa arkusza | Zadanie | Kolor karty | Najważniejsze elementy i suma kontrolna |
| :--- | :--- | :--- | :--- | :--- |
| 1 | `EWIDENCJA` | E7 zad. 2 | Czerwony | Tabela `SprzedażTygodniowa` (`A1:E386`, 385 rekordów), kolumna `Kwartał`, nowy wpis z `31.12.2024` (`500`), formuły w `G2:G3`. Suma: **2 683 073**. |
| 2 | `KONSPEKT` | E7 zad. 1 | Brak | Kopia danych przed dodaniem `31.12.2024` (384 rekordy), posortowana `Kraj -> Miasto -> Data`, sumy częściowe na 2 poziomach. Suma końcowa: **2 682 573**. |
| 3 | `PRZEFILTROWANE` | E7 zad. 3 | Czerwony | Kopia arkusza `EWIDENCJA`, tabela `SprzedażTygodniowa_Filtr`, sort malejąco po `Kwota`, filtr `Hogwart` + `01.07–31.08.2024` (32 wiersze), wiersz sumy tylko dla `Kwota`: **333 584**. |
| 4 | `Sprzedaż w miastach` | E8 1a | Brak | Tabela przestawna `PTSprzedazWMiastach`: `Miasto` w wierszach, `Kraj` w filtrze, `Suma z Kwota` = **2 683 073**. |
| 5 | `Kwartalnie w miastach` | E8 1b + 1f | Brak | Tabela przestawna `PTKwartalnieWMiastach` w układzie **konspekt**: `Kwartał` i `Miasto` w wierszach, `Kraj` w filtrze, `Suma z Kwota` i `% Całkowitej sumy`. |
| 6 | `Miasta kwartałami` | E8 1c | Brak | Tabela przestawna `PTMiastaKwartalami` w układzie **konspekt**: `Kraj`, `Miasto` i `Kwartał` w wierszach, `Suma z Kwota`. |
| 7 | `Tabelarycznie` | E8 1d + 2a | Brak | Tabela przestawna `PTTabelarycznie`: `Miasto` w wierszach, `Kwartał` w kolumnach, `Kraj` w filtrze + przestawny wykres **skumulowany kolumnowy**. |
| 8 | `Miasta w miesiącach` | E8 1e | Brak | Tabela przestawna `PTMiastaWMiesiacach`: `Miesiące (Data)` w wierszach, `Kraj` i `Miasto` w kolumnach. |
| 9 | `Dodatkowa tabela` | E8 1g | Brak | Dodatkowa tabela `PTDodatkowa` do ćwiczeń przestawności (`Kraj` w wierszach, `Kwartał` + miary w kolumnach, `Miasto` w filtrze, Suma i Średnia z Kwoty). |
| 10 | `Wykres przestawny` | E8 2b | Brak | Wykres przestawny utworzony bezpośrednio z `EWIDENCJA` (`PTWykresPrzestawny` + wykres kolumnowy grupowany `Kwartał` × `Kraj`). |

### Zestawienie liczb kontrolnych

* **Suma całkowita (385 rekordów, z `31.12.2024`):** `2 683 073`
  * **Hogwart:** `1 832 430` (`68,30%`) — `Gryffindor`: `729 745`, `Slytherin`: `532 215`, `Hufflepuff`: `379 080`, `Ravenclaw`: `191 390`.
  * **Ringland:** `850 643` (`31,70%`) — `Gondor`: `305 588`, `Isengard`: `271 013`, `Rohan`: `169 866`, `Mordor`: `104 176`.
* **Suma w `KONSPEKT` (384 rekordy, przed `31.12.2024`):** `2 682 573` (`Isengard` ma tu `270 513`, a `Ringland` `850 143` — dokładnie o `500` mniej).
* **Kwartały (wszystkie dane):** `Q1` = `470 170` (`17,52%`), `Q2` = `605 787` (`22,58%`), `Q3` = `765 600` (`28,53%`), `Q4` = `841 516` (`31,36%`).
* **Filtr w `PRZEFILTROWANE` (Hogwart, 01.07–31.08.2024):** `32` widoczne wiersze, suma = `333 584`.

---

## 3. Temat E7 – Organizowanie danych w tabele (szczegółowy przewodnik)

### Zadanie 1: Konspekt i obliczanie sum pośrednich (`KONSPEKT`)

#### Czym jest konspekt i po co go używamy?
**Konspekt** (ang. *Outline*) w Excelu porządkuje długą listę danych w hierarchiczne bloki i zwija je przyciskami poziomów (`1`, `2`, `3`, `4` oraz `+`/`−`). Narzędzie **Suma częściowa** automatycznie wstawia wiersze podsumowań pod każdą grupą i tworzy strukturę konspektu na lewym marginesie arkusza.

> **Złota zasada:** Narzędzie **Suma częściowa** wstawia podsumowanie za każdym razem, gdy zmienia się wartość w wybranej kolumnie. Jeśli najpierw nie posortujesz tabeli według kolumn grupujących, ten sam kraj lub miasto pojawi się wielokrotnie w osobnych blokach.

#### Instrukcja krok po kroku (`1a–1e`)

* **1a) Utwórz kopię arkusza `EWIDENCJA` i nazwij ją `KONSPEKT`:**
  1. Kliknij prawym przyciskiem myszy kartę `EWIDENCJA` na dole ekranu.
  2. Wybierz **Przenieś lub kopiuj...**, zaznacz pole **Utwórz kopię** i kliknij **OK**.
  3. Kliknij dwukrotnie nazwę nowej karty i zmień ją na `KONSPEKT`.
* **1b) Posortuj dane wielopoziomowo (`Kraj`, `Miasto`, `Data`):**
  1. Kliknij dowolną pojedynczą komórkę wewnątrz danych (np. `A2`).
  2. Wejdź na wstążkę **Dane → Sortowanie i filtrowanie → Sortuj**.
  3. Ustaw pierwszy poziom: *Sortuj według* `Kraj` (od A do Z).
  4. Kliknij **Dodaj poziom** i ustaw *Następnie według* `Miasto` (od A do Z).
  5. Kliknij **Dodaj poziom** i ustaw *Następnie według* `Data` (od najstarszej do najnowszej), po czym kliknij **OK**.
* **1c) Zastosuj dwa poziomy sum częściowych:**
  1. Ustaw kursor wewnątrz posortowanych danych (np. `A2`) i wybierz **Dane → Konspekt → Suma częściowa**.
  2. W pierwszym oknie ustaw:
     * *Dla każdej zmiany w:* `Kraj`
     * *Użyj funkcji:* `Suma`
     * *Dodaj sumę częściową do:* zaznacz wyłącznie `Kwota`
     * *Zamień bieżące sumy częściowe:* może pozostać zaznaczone przy pierwszym kroku. Kliknij **OK**.
  3. Ponownie otwórz **Dane → Konspekt → Suma częściowa** (nie usuwaj poprzednich sum!).
  4. W drugim oknie ustaw:
     * *Dla każdej zmiany w:* `Miasto`
     * *Użyj funkcji:* `Suma`
     * *Dodaj sumę częściową do:* `Kwota`
     * **ODZNACZ** pole wyboru **Zamień bieżące sumy częściowe** (kluczowe — inaczej stracisz sumy dla krajów!). Kliknij **OK**.
* **1d) Nawiguj przyciskami konspektu:**
  * Po lewej stronie numerów wierszy pojawi się pasek z przyciskami `1`, `2`, `3`, `4` oraz symbolami `−` i `+`.
  * `1` pokazuje tylko `Suma końcowa` (`2 682 573`).
  * `2` pokazuje sumy krajów (`Hogwart Suma`, `Ringland Suma`).
  * `3` pokazuje sumy poszczególnych miast.
  * `4` rozwija wszystkie pojedyncze tygodnie.
  * Jeśli każde z 8 miast i oba kraje występują na widoku poziomu 3 tylko po jednym razie, konspekt działa bezbłędnie.
* **1e) Sprawdź wiersze podsumowań i przetestuj zmianę kwoty z `28.12`:**
  1. Kliknij komórkę z sumą miasta (np. `D50` dla `Gryffindor Suma`) albo z sumą kraju (`D198`).
  2. Spójrz na pasek formuły.
  3. Zmień tymczasowo kwotę w wierszu z datą `28.12` (np. `D49` lub `D393`), obserwuj zmianę sum, a następnie cofnij zmianę (`Ctrl+Z`), aby zachować oryginalne dane.

#### Odpowiedzi na pytania z podpunktu 1e
* **Czy w komórkach podsumowania są formuły czy wartości?**
  Są tam **formuły**, a konkretnie funkcja `=SUMA.CZĘŚCIOWA(9; D2:D49)` (ang. `=SUBTOTAL(9, D2:D49)`). Liczba `9` oznacza funkcję `SUMA`, która automatycznie pomija inne zagnieżdżone formuły `SUBTOTAL` w tym samym zakresie, dlatego sumy krajów i suma końcowa nie liczą kwot podwójnie.
* **Czy coś uległo zmianie po zmianie wartości z `28.12`?**
  Tak — automatycznie przeliczyła się suma danego miasta, suma odpowiedniego kraju oraz `Suma końcowa`.

---

### Zadanie 2: Definiowanie, formatowanie i rozszerzanie tabel (`EWIDENCJA`)

#### Czym jest Tabela Excela i po co ją tworzymy?
Zwykły zakres komórek nie ma własnej nazwy ani struktury bazy danych. Gdy konwertujesz zakres na **Tabelę Excela** (`Wstawianie → Tabela` lub `Ctrl+T`):
1. Zakres sam rozszerza się po dopisaniu nowej kolumny z prawej strony lub nowego wiersza na dole.
2. Wpisana w jednym wierszu formuła kopiuje się automatycznie na całą kolumnę (*kolumna obliczeniowa*).
3. Formuły korzystają z czytelnych **odwołań strukturalnych** (np. `SprzedażTygodniowa[@Kwota]`) zamiast sztywnych adresów typu `D3`.
4. Tabele przestawne podpięte pod nazwę tabeli zawsze widzą wszystkie nowo dopisane rekordy po odświeżeniu.

#### Instrukcja krok po kroku (`2a–2g`)

* **2a) Zmień kolor zakładki `EWIDENCJA` i utwórz Tabelę:**
  1. Kliknij prawym przyciskiem kartę `EWIDENCJA` → **Kolor karty** → wybierz czerwony (`#FF0000`).
  2. Kliknij dowolną komórkę wewnątrz danych `A1:D385` (np. `A1`).
  3. Wybierz **Wstawianie → Tabele → Tabela** (upewnij się, że zaznaczono „Moja tabela ma nagłówki” i zakres to `=$A$1:$D$385`). Kliknij **OK**.
* **2b) Nazwij tabelę `SprzedażTygodniowa`:**
  1. Kliknij wewnątrz tabeli, aby pojawiła się kontekstowa karta **Projekt tabeli** (w starszym Excelu: *Narzędzia tabel / Projektowanie*).
  2. W polu **Właściwości → Nazwa tabeli** wpisz dokładnie `SprzedażTygodniowa` (bez spacji!) i naciśnij `Enter`.
* **2c) Zastosuj własny styl tabeli:**
  1. Na karcie **Projekt tabeli** otwórz galerię **Style tabeli**.
  2. Wybierz wybrany styl różniący się od domyślnego niebieskiego — w naszym pliku zastosowaliśmy zielony styl **`TableStyleMedium11`** (wyróżnia się na tle domyślnego niebieskiego oraz czerwonego stylu).
* **2d) Dodaj kolumnę `Kwartał`:**
  1. Kliknij w komórkę `E1` (bezpośrednio obok nagłówka `Kwota`) i wpisz `Kwartał`, a następnie naciśnij `Enter` (tabela automatycznie włączy kolumnę `E` do swojego zakresu; możesz też przeciągnąć mały niebieski uchwyt w prawym dolnym rogu tabeli `D385`).
* **2e) Oblicz numer kwartału formułą:**
  1. Kliknij komórkę `E2` (pierwsza komórka danych pod nagłówkiem `Kwartał`).
  2. Wpisz dokładnie:
     ```excel
     =ZAOKR.GÓRA(MIESIĄC(C2)/3; 0)
     ```
     (w wersji angielskiej `=ROUNDUP(MONTH(C2)/3, 0)`).
  3. Naciśnij `Enter`. Excel automatycznie wypełni wszystkie komórki aż do wiersza `385`.
* **2f) Dodaj nowy wiersz na końcu tabeli (`31 grudnia 2024`):**
  1. Kliknij ostatnią komórkę tabeli (`E385`) i naciśnij klawisz `Tab` (albo wpisz dane w `A386` — tabela sama rozszerzy zakres do `A1:E386`).
  2. Wpisz brakujące dane:
     * `A386` (`Kraj`): `Ringland`
     * `B386` (`Miasto`): `Isengard`
     * `C386` (`Data`): `2024-12-31` (czysta data bez dopisanej godziny!)
     * `D386` (`Kwota`): `500`
     * `E386` (`Kwartał`): obliczy się automatycznie jako `4`.
* **2g) Pozostałe opcje wstążki `Projekt tabeli`:**
  * *Opcje stylu tabeli:* wiersz nagłówka, wiersz sumy, paski wierszy/kolumn, przycisk filtru.
  * *Narzędzia:* **Podsumuj za pomocą tabeli przestawnej**, **Usuń duplikaty**, **Konwertuj na zakres**, **Wstaw fragmentator**.

#### Pytanie z żółtej ramki na arkuszu `EWIDENCJA` (kolumna `G`)
* **Treść zadania z ramki:** W komórce `G2` stoi zwykła formuła `=D2/7`. Po zamianie danych na tabelę wpisz w `G3` znak `=`, wskaż myszką komórkę `D3` i dopisz `/7`. Dlaczego formuła ma inną postać?
* **Odpowiedź:** Excel zamienia adres `D3` na **odwołanie strukturalne**:
  ```excel
  =SprzedażTygodniowa[@Kwota]/7
  ```
  (lub w starszej składni `=SprzedażTygodniowa[[#This Row],[Kwota]]/7`). Znak `@` oznacza *ten sam wiersz*, a `[Kwota]` odwołuje się do nazwy kolumny tabeli zamiast do współrzędnych komórki. Taka formuła działa niezależnie od tego, czy przesuniesz kolumnę lub posortujesz tabelę.

---

### Zadanie 3: Szybkie sortowanie i filtrowanie tabeli (`PRZEFILTROWANE`)

#### Po co używamy filtrów i funkcji `SUBTOTAL(109)`?
W tabeli możesz szybko ograniczać widok do wybranych rekordów i sumować wyłącznie widoczne komórki, nie niszcząc reszty danych.

#### Instrukcja krok po kroku (`3a–3d`)

* **3a) Utwórz kopię `EWIDENCJA` i nazwij ją `PRZEFILTROWANE`:**
  1. Kliknij prawym przyciskiem kartę `EWIDENCJA` → **Przenieś lub kopiuj...** → zaznacz **Utwórz kopię** → **OK**.
  2. Zmień nazwę kopii na `PRZEFILTROWANE`. Ponieważ `EWIDENCJA` ma już czerwoną zakładkę, kopia naturalnie dziedziczy czerwony kolor karty i styl tabeli (`TableStyleMedium11`).
* **3b) Posortuj tabelę według `Kwoty`:**
  1. Kliknij strzałkę rozwijaną przy nagłówku `Kwota` (`D1`).
  2. Wybierz **Sortuj od największych do najmniejszych** (lub rosnąco).
* **3c) Przefiltruj: `Hogwart` w wakacje letnie (`1 lipca – 31 sierpnia`):**
  1. Kliknij strzałkę w nagłówku `Kraj` (`A1`), zostaw zaznaczony tylko `Hogwart` i kliknij **OK**.
  2. Kliknij strzałkę w nagłówku `Data` (`C1`) → **Filtry dat → Między...** (ang. *Between...*).
  3. Wpisz datę początkową `2024-07-01` (*jest po lub równa*) oraz końcową `2024-08-31` (*jest przed lub równa*) i kliknij **OK**.
  4. Na ekranie zostanie dokładnie **32 widoczne wiersze**.
* **3d) Włącz `Wiersz sumy` tylko dla `Kwota`:**
  1. Kliknij dowolną komórkę wewnątrz tabeli.
  2. Wejdź na kartę **Projekt tabeli → Opcje stylu tabeli → zaznacz `Wiersz sumy`**.
  3. Excel domyślnie wstawi sumę w ostatniej kolumnie (`Kwartał` w `E387`). Kliknij `E387`, otwórz listę rozwijaną i wybierz **Brak**.
  4. Kliknij `D387` (pod kolumną `Kwota`), otwórz listę rozwijaną i wybierz **Suma**.
  5. Wynik dla widocznych wierszy to **`333 584`**.

#### Odpowiedź na pytanie z podpunktu 3d
* **Co jest faktycznie sumowane — widoczne (filtrowane) dane, czy wszystko w kolumnie `Kwota`?**
  Sumowane są **wyłącznie widoczne (przefiltrowane) dane** (`333 584`), a nie cała kolumna (`2 683 073`). Wiersz sumy wstawia funkcję `=SUBTOTAL(109; [Kwota])` (`109` oznacza sumowanie wyłącznie nieukrytych wierszy).

---

## 4. Temat E8 – Tabele przestawne i wykresy przestawne

### Czym jest Tabela przestawna i po co jej używamy?
**Tabela przestawna** (ang. *PivotTable*) w kilka sekund agreguje tysiące wierszy w dowolnych przekrojach — bez pisania ani jednej formuły `SUMA.WARUNKÓW`. Możesz swobodnie przeciągać pola między czterema obszarami: **Filtry**, **Kolumny**, **Wiersze** i **Wartości**.

### Trzy układy raportu (ważne w zadaniach 1b i 1c!)
Na karcie **Projektowanie → Układ → Układ raportu** wybierasz sposób wyświetlania wierszy:

| Układ raportu | Jak wygląda? | Gdzie stoją pola wierszy? | Gdzie stoją sumy częściowe? |
| :--- | :--- | :--- | :--- |
| **Kompaktowy** *(domyślny)* | Wszystkie pola wierszy siedzą w jednej kolumnie `A` z wcięciami | W jednej kolumnie | Na górze grupy |
| **Konspekt** *(wymagany w 1b i 1c!)* | Każde pole dostaje osobną kolumnę, a wiersz nadrzędny ma własny wiersz sumy nad dziećmi | W osobnych kolumnach | **Na górze grupy** |
| **Tabelaryczny** | Każde pole ma osobną kolumnę, a pierwszy element podrzędny stoi w tym samym wierszu co rodzic | W osobnych kolumnach | **Na dole grupy** |

> **Uwaga przy tworzeniu:** Ustaw kursor w tabeli `SprzedażTygodniowa` w arkuszu `EWIDENCJA` i wybierz **Wstawianie → Tabele → Tabela przestawna → Z tabeli/zakresu → Nowy arkusz**. W naszym pliku wszystkie 7 tabel przestawnych korzysta ze wspólnej pamięci podręcznej (*PivotCache*) oraz zielonego stylu `PivotStyleMedium4`.

---

### Zadanie 1: Raporty w tabelach przestawnych (`1a–1i`)

* **1a) Arkusz `Sprzedaż w miastach`:**
  * *Wiersze:* `Miasto` | *Wartości:* `Suma z Kwota` | *Filtry:* `Kraj`
  * *Cel:* Szybkie porównanie łącznej sprzedaży w każdym z 8 miast z możliwością odfiltrowania jednego kraju (`Hogwart` lub `Ringland`).
* **1b) Arkusz `Kwartalnie w miastach` (+ modyfikacja z pkt 1f):**
  * *Wiersze:* najpierw `Kwartał`, pod nim `Miasto` | *Filtry:* `Kraj`
  * *Wartości (1b):* `Suma z Kwota`
  * *Wartości (1f):* Przeciągnij pole `Kwota` drugi raz do obszaru **Wartości**. Kliknij je prawym przyciskiem myszy → **Pokaż wartości jako → % Całkowitej sumy** (w pliku nazwaliśmy tę kolumnę `% Całkowitej sumy`).
  * *Układ:* **Projektowanie → Układ raportu → Pokaż w formie konspektu**.
  * *Ćwiczenie z 1b:* Ustaw kursor na numerze kwartału w kolumnie `A` i sprawdź na karcie **Analiza tabeli przestawnej → Aktywne pole** przyciski **Rozwiń pole** / **Zwiń pole** (zwijają lub rozwijają wszystkie kwartały naraz).
* **1c) Arkusz `Miasta kwartałami`:**
  * *Wiersze:* w kolejności od góry `Kraj`, `Miasto`, `Kwartał` | *Wartości:* `Suma z Kwota`
  * *Układ:* **Projektowanie → Układ raportu → Pokaż w formie konspektu** (`Hogwart` w wierszu 4, `Gryffindor` w wierszu 5, kwartały `1–4` w wierszach 6–9).
* **1d) Arkusz `Tabelarycznie`:**
  * *Wiersze:* `Miasto` | *Kolumny:* `Kwartał` | *Filtry:* `Kraj` | *Wartości:* `Suma z Kwota`
  * *Cel:* Czytelna macierz (krzyżówka) porównująca 8 miast w 4 kwartałach obok siebie.
* **1e) Arkusz `Miasta w miesiącach`:**
  * *Wiersze:* `Data` → kliknij prawym przyciskiem dowolną datę w tabeli przestawnej → **Grupuj...** (albo **Analiza tabeli przestawnej → Grupowanie → Grupuj pole**) → zaznacz **tylko `Miesiące`** i kliknij **OK**.
  * *Kolumny:* `Kraj`, a pod nim `Miasto` | *Wartości:* `Suma z Kwota`.
  * *Uwaga dla nowszego Excela (365 / Mac):* Excel przy wrzuceniu pola `Data` do wierszy potrafi sam dodać dodatkowe pole `Dni (Data)`. W oknie **Grupuj** odznacz `Dni` i zostaw same `Miesiące`, żeby uzyskać czyste 12 wierszy `sty–gru`.
* **1g) Arkusz `Dodatkowa tabela`:**
  * Dowolna tabela ćwiczeniowa do pokazania „przestawności” (przeciągania pól między osiami i zmiany funkcji agregacji).
  * W pliku zawiera: *Wiersze:* `Kraj`, *Kolumny:* `Kwartał` i `Wartości`, *Filtry:* `Miasto`, *Wartości:* `Suma z Kwota` oraz `Średnia z Kwota`.

#### Odpowiedzi na pytania z podpunktu 1h
* **Co znajduje się w komórkach dowolnej tabeli przestawnej — formuły czy obliczone wartości?**
  W komórkach tabeli przestawnej znajdują się **obliczone wartości** (liczby), a nie klasyczne formuły arkusza typu `=SUMA(...)`. Silnik tabeli przestawnej wylicza je w pamięci podręcznej (*PivotCache*) i wpisuje gotowy wynik do siatki.
* **Czy tabela przestawna zostanie automatycznie zaktualizowana po zmianie wartości w danych źródłowych?**
  **Nie.** Zmiana liczby w arkuszu `EWIDENCJA` nie aktualizuje tabeli przestawnej od razu. Musisz kliknąć prawym przyciskiem myszy tabelę przestawną i wybrać **Odśwież** (lub na wstążce **Analiza tabeli przestawnej → Odśwież → Odśwież wszystko**).

---

### Zadanie 2: Wykresy przestawne (`2a–2c`)

#### Czym jest Wykres przestawny?
**Wykres przestawny** jest bezpośrednio połączony z tabelą przestawną. Ma na sobie interaktywne przyciski pól (do filtrowania wprost na wykresie), a każda zmiana układu osi automatycznie przebudowuje zarówno wykres, jak i powiązaną tabelę.

#### Instrukcja krok po kroku (`2a–2b`)

* **2a) Wykres w arkuszu `Tabelarycznie`:**
  1. Kliknij dowolną komórkę wewnątrz tabeli przestawnej w arkuszu `Tabelarycznie`.
  2. Wybierz **Analiza tabeli przestawnej → Narzędzia → Wykres przestawny**.
  3. Wybierz typ **Kolumnowy → Skumulowany kolumnowy** i kliknij **OK**.
  4. Nadaj czytelny tytuł, np. `Sprzedaż kwartalna w miastach`. Każde miasto ma jeden słupek podzielony kolorystycznie na 4 kwartały.
* **2b) Wykres przestawny bez wcześniejszej tabeli (`Wykres przestawny`):**
  1. Przejdź do arkusza `EWIDENCJA` i kliknij wewnątrz tabeli `SprzedażTygodniowa`.
  2. Wybierz **Wstawianie → Wykresy → Wykres przestawny → Wykres przestawny** (lub dolną część przycisku *Tabela przestawna*).
  3. Wybierz **Nowy arkusz**, nazwij go `Wykres przestawny` i zaprojektuj przejrzysty układ:
     * *Oś (Kategorie):* `Kwartał`
     * *Legenda (Serie):* `Kraj`
     * *Filtry:* `Miasto`
     * *Wartości:* `Suma z Kwota`
     * Typ wykresu: **Kolumnowy grupowany** z tytułem `Kwartalna sprzedaż według krajów`.

#### Odpowiedzi na pytania z podpunktu 2c
* **Czy wykres przestawny może być „przestawiany” i filtrowany?**
  **Tak.** Możesz przeciągać pola między obszarami *Filtry*, *Legenda (Serie)*, *Oś (Kategorie)* i *Wartości* dokładnie tak samo jak w tabeli przestawnej oraz używać przycisków filtrów widocznych na samym wykresie.
* **Co dzieje się z tabelą leżącą u podstaw wykresu, gdy wykres jest modyfikowany?**
  Tabela przestawna zmienia się **synchronicznie i natychmiast** w ten sam sposób: filtr włączony na wykresie filtruje tabelę, a zamiana pól między osią i legendą zamienia w tabeli wiersze z kolumnami.

---

## 5. Temat E9 – Tworzenie złożonych dokumentów (osadzanie i łączenie obiektów w Wordzie)

### Po co osadzamy i łączymy obiekty oraz co wybraliśmy do gotowego raportu?
Windowsowy mechanizm **OLE** (*Object Linking and Embedding*) pozwala umieszczać fragmenty jednego programu (np. Excela) wewnątrz dokumentu innego programu (np. Worda). W laboratorium porównujemy trzy metody wklejania:

| Metoda w Wordzie | Co powstaje? | Czy zależy od zewnętrznego pliku `.xlsx`? | Jak reaguje na edycję? |
| :--- | :--- | :--- | :--- |
| **Zwykłe `Wklej`** *(dla tabeli)* | Natywna tabela Worda | **Nie** | Edytujesz komórki jak zwykły tekst i styl tabeli w Wordzie. |
| **`Wklej specjalnie → Arkusz MS Excel - obiekt` (bez łącza)** | **Osadzony obiekt OLE** (`word/embeddings/`) | **Nie** (kopia siedzi wewnątrz pliku `.docx`) | Dwuklik otwiera wstążkę Excela wewnątrz Worda, ale nie zmienia źródłowego pliku `.xlsx`. |
| **`Wklej specjalnie → Wklej łącze`** | **Połączony obiekt OLE** | **Tak** (wskazuje na ścieżkę pliku `.xlsx`) | Dwuklik otwiera zewnętrzny plik Excela; zmiany w Excelu aktualizują tabelę w Wordzie. |
| **Wklejony wykres z przerwanym łączem** | Natywny wektorowy wykres w `.docx` | **Nie** (po użyciu *Przerwij łącze*) | Zachowuje idealną jakość wektorową i działa na każdym komputerze bez pliku Excela. |

> **Co zawiera plik [`TI.1_Word.E9_Raport_HogDog.docx`](./TI.1_Word.E9_Raport_HogDog.docx)?**
> Zgodnie z poleceniem dokument zawiera wyłącznie zwięzły raport dla zarządu firmy Hog-Dog sp. z o.o. oraz obiekty z Excela (bez odpowiedzi na pytania ćwiczeniowe):
> 1. Krótkie wprowadzenie liczbowe (sprzedaż całkowita, podział na kraje, najlepsze/najsłabsze miasto, trend kwartalny).
> 2. Tabelę `Sprzedaż w miastach` wklejoną jako natywną tabelę Worda.
> 3. Tabelę `Sprzedaż kwartalna w miastach` (`Tabelarycznie`) osadzoną jako obiekt Excela **bez zewnętrznego łącza** (`word/embeddings/Microsoft_Excel_Worksheet.xlsx`).
> 4. Oba wykresy przestawne (`Sprzedaż kwartalna w miastach` oraz `Kwartalna sprzedaż według krajów`) wklejone i odłączone przez *Przerwij łącze*, dzięki czemu cały dokument `.docx` jest **w 100% samodzielny** i otwiera się poprawnie nawet po przeniesieniu na inny komputer bez Excela.

---

### Instrukcja wykonania ćwiczeń z E9 i odpowiedzi na wszystkie pytania z PDF-a

#### Zadanie 1: Osadzanie nowych (pustych) obiektów (`1a–1e`)

* **1a) Wstaw pusty arkusz Excela:**
  * Ścieżka: **Wstawianie → Tekst → Obiekt → zakładka `Utwórz nowy` → `Arkusz programu Microsoft Excel`**.
  * Wpisz kilka liczb, powiększ ramkę za narożnik i kliknij na zewnątrz obiektu.
* **1b) Pytanie: Czy obiekt poddaje się edycji? Co się dzieje ze wstążką po dwukrotnym kliknięciu obiektu?**
  * **Tak**, obiekt poddaje się pełnej edycji (możesz zmieniać liczby, formuły i formaty).
  * Po dwukrotnym kliknięciu następuje tzw. *aktywacja w miejscu* (*in-place activation*): **wstążka Worda zamienia się na wstążkę programu Excel** (pojawiają się karty *Formuły*, *Dane*, *Narzędzia główne* z Excela, siatka kolumn `A, B, C` i wierszy `1, 2, 3`), przy czym pasek tytułu okna nadal należy do Worda. Kliknięcie poza ramką przywraca wstążkę Worda.
* **1c) Dwa wykresy w Wordzie — jaka jest między nimi różnica?**
  * *Metoda 1 (`Wstawianie → Ilustracje → Wykres`):* Tworzy **natywny wykres Worda**. Używa kart Worda (*Projekt wykresu* i *Formatowanie*) oraz otwiera osobne małe okienko Excela tylko z tabelką danych.
  * *Metoda 2 (`Wstawianie → Tekst → Obiekt → Wykres programu Microsoft Excel`):* Tworzy **osadzony obiekt OLE starego typu (skoroszyt Excela z zakładkami `Wykres1` i `Arkusz1`)**. Dwukrotne kliknięcie zamienia wstążkę Worda na interfejs Excela.
* **1d) Arkusz wstawiony przez `Wstawianie → Tabele → Tabela → Arkusz kalkulacyjny programu Excel` — czy jest różny od tego z pkt 1a?**
  * **Nie.** Technicznie powstaje dokładnie **ten sam typ obiektu OLE** (`Excel.Sheet.12` / `Arkusz programu Microsoft Excel`). Różni się wyłącznie przycisk na wstążce, którym go wywołujesz.
* **1e) Zapisz dokument Worda — ile plików się zapisało i gdzie znajdują się wstawione skoroszyty Excela?**
  * Na dysku zapisał się **jeden plik** dokumentu Worda (`.docx`).
  * Osadzone skoroszyty Excela **nie leżą na dysku jako osobne pliki** — znajdują się **wewnątrz samego pliku `.docx`** (ponieważ format `.docx` jest archiwum ZIP; po rozpakowaniu zobaczysz je w podfolderze `word/embeddings/`).

#### Zadanie 2: Osadzanie i dołączanie istniejących obiektów (`2a–2g`)

* **2a) Zwykłe `Wklej` tabeli przestawnej z Excela do Worda:**
  * Skopiuj zakres tabeli przestawnej w Excelu (`Ctrl+C`) i wklej w Wordzie (`Ctrl+V`). Powstaje **zwykła tabela Worda**.
* **2b) `Wklej specjalnie → Arkusz programu Microsoft Excel - obiekt`:**
  * Dwa wiersze poniżej kliknij **Narzędzia główne → Schowek → Wklej (strzałka) → Wklej specjalnie... → `Arkusz programu Microsoft Excel - obiekt`** (opcja *Wklej*, **bez** *Wklej łącze*). Powstaje osadzona kopia arkusza OLE.
* **2c) Pytanie: Jaka jest różnica przy kliknięciu 1 raz i 2 razy na obu tabelach?**
  * *Tabela z 2a (zwykłe Wklej):* Jedno kliknięcie stawia kursor tekstowy wewnątrz komórki i wyświetla karty Worda (*Projekt tabeli*, *Układ*). Dwuklik zaznacza pojedyncze słowo lub liczbę.
  * *Obiekt z 2b (Wklej specjalnie jako obiekt):* Jedno kliknięcie zaznacza całą tabelę jako jedną ramkę graficzną z uchwytami do skalowania. Dwuklik uruchamia edytor Excela wewnątrz dokumentu Worda.
* **2d) Pytanie: Czy zmiana wewnątrz osadzonego obiektu (z 2b) ma odzwierciedlenie w pliku Excela? Czy odwrotna aktualizacja zadziała?**
  * **Nie i nie.** Wklejenie obiektu bez zaznaczenia *Wklej łącze* tworzy niezależną kopię arkusza zamkniętą wewnątrz pliku `.docx`. Zmiany zrobione po dwukliku w Wordzie nie trafiają do źródłowego pliku `.xlsx`, a zmiany w pliku `.xlsx` nie aktualizują kopii w Wordzie.
* **2e) `Wklej specjalnie → Wklej łącze` — wyjaśnij zachowanie obiektu:**
  * Zaznaczony zostaje adres do pliku zewnętrznego na dysku.
  * Po dwukrotnym kliknięciu obiekt **nie otwiera się wewnątrz Worda**, lecz uruchamia lub przenosi Cię bezpośrednio do **oryginalnego okna programu Excel** z otwartym plikiem źródłowym.
  * Zmiany zapisane w pliku Excela automatycznie aktualizują obiekt w Wordzie (albo po użyciu *Aktualizuj łącze* / ponownym otwarciu dokumentu). Jeśli przeniesiesz plik `.docx` bez pliku `.xlsx` na inny komputer, łącze przestanie działać.
* **2f) Wykres Excela wklejony do Worda — co po 1 kliknięciu, co po 2 kliknięciach, jak zmienić dane i co się aktualizuje?**
  * *Pojedyncze kliknięcie:* Zaznacza obiekt wykresu i wyświetla w Wordzie narzędzia wykresu (*Projekt wykresu*, *Formatowanie*) albo uchwyty ramki (dla obiektu OLE).
  * *Dwukrotne kliknięcie:* Dla wykresu wklejonego domyślnie otwiera panel boczny *Formatuj obszar wykresu*; dla wykresu wklejonego jako obiekt OLE otwiera edycję w Excelu.
  * *Zmiana danych źródłowych:* Kliknij wykres → **Projekt wykresu → Dane → Edytuj dane** lub zmień wartość bezpośrednio w otwartym pliku Excela.
  * *Aktualizacja (przy zachowanym łączu):* Zmiana komórki w Excelu automatycznie przelicza słupki i skalę osi zarówno na wykresie w Excelu, jak i na wykresie w Wordzie. (W naszym finalnym pliku `TI.1_Word.E9_Raport_HogDog.docx` celowo przerwaliśmy łącze, żeby dokument działał samodzielnie bez pliku `.xlsx`).
* **2g) Jakie istnieją inne metody wstawiania obiektów w Wordzie? Czy obiekty można też osadzać w Excelu?**
  * W oknie **Wstawianie → Obiekt** masz dwie zakładki:
    1. **Utwórz nowy** — tworzysz nowy obiekt innego programu (np. *Równanie Microsoft*, *Obraz Paint*, *Dokument OpenDocument*, *Arkusz/Wykres Excel*, *Prezentacja PowerPoint*).
    2. **Utwórz z pliku** — wskazujesz istniejący plik z dysku; możesz dodatkowo zaznaczyć **Połącz z plikiem** (utrzymuje synchronizację) oraz **Wyświetl jako ikonę** (zamiast podglądu treści pokazuje klikalną ikonę pliku).
  * Dodatkowo polecenie **Wstawianie → Tekst → Obiekt (strzałka) → Tekst z pliku...** wstawia bezpośrednio treść innego dokumentu tekstowego.
  * **Czy obiekty można osadzać w arkuszach Excela?** **Tak** — identyczny mechanizm działa w Excelu na wstążce **Wstawianie → Tekst → Obiekt** (możesz tam osadzić lub połączyć np. dokument Worda, plik PDF czy grafikę).
