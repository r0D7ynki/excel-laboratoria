# TI.1 – Arkusz kalkulacyjny: zadania E1, E2 i E3

Repozytorium zawiera gotowy skoroszyt zaliczeniowy do laboratorium TI.1.
Plik `TI.1_Excel.E1+E2+E3_PlikRoboczy.xlsx` przechodzi przez wszystkie zadania
z tematów E1, E2 i E3.
Ten dokument pokazuje, co trzeba było zrobić i jak to zostało zrobione.

---

## 1. Zawartość skoroszytu

| Arkusz | Kolor karty | Czego dotyczy |
| --- | --- | --- |
| `BusinessPlan` | domyślny | Tematy E2 i E3 |
| `Potter` | czerwony | Temat E1 – tabela wyników w Quidditchu |
| `Weasley` | czerwony | Temat E1.4a – pusty arkusz |

Arkusz `Malfoy` powstał w zadaniu E1.4b. Instrukcja każe go usunąć w zadaniu E1.4d.
Dlatego nie ma go w gotowym pliku.

---

## 2. Jak otworzyć plik

1. Pobierz plik `TI.1_Excel.E1+E2+E3_PlikRoboczy.xlsx`.
2. Otwórz go w programie Microsoft Excel.
3. Plik ma włączoną opcję pełnego przeliczania przy otwarciu.
4. Jeśli zobaczysz stare wyniki, naciśnij `Ctrl` + `Alt` + `F9`.

Nazwy funkcji są zapisane po angielsku. Excel pokazuje je w języku programu.
Na przykład `ROUND` wyświetla się jako `ZAOKRĄGLIJ`, a `AVERAGE` jako `ŚREDNIA`.

---

## 3. Temat E1 – arkusze kalkulacyjne i skoroszyty

Zadania E1 wykonuje się w nowym, pustym skoroszycie.
Na końcu arkusze trafiają do pliku roboczego. Efekt to arkusz `Potter`.

### 3.1. Zadanie 1 – skoroszyty i szablony

Ta część to ćwiczenie. Nie zostawia śladu w pliku.

1. Wybierz `Plik` / `Nowy` i utwórz kilka skoroszytów z różnych szablonów.
2. Przełączaj się między oknami klawiszami `Alt` + `Tab`.
3. Przełączaj się między skoroszytami klawiszami `Ctrl` + `Tab` oraz `Ctrl` + `Shift` + `Tab`.
4. Obejrzyj wstążkę `Widok`: widoki arkusza, linie siatki, powiększenie, dzielenie okna.

### 3.2. Zadanie 2 – wprowadzanie wartości i dopasowanie kolumn

| Krok | Co zrobić |
| --- | --- |
| 2a | Wpisz `Potter Harry` w A1, `Weasley Ron` w B1, `Malfoy Draco` w C1. |
| 2b | Dopasuj szerokość kolumn A–C na trzy sposoby: przeciągnij krawędź w nagłówku, kliknij dwukrotnie krawędź (autodopasowanie) oraz `Narzędzia główne` / `Komórki` / `Formatuj` / `Szerokość kolumny`. |
| 2c | Wpisz `1000` w A2. Przeciągnij uchwyt wypełnienia do A13. |
| 2d | Wpisz `0,1` w B2 i `0,15` w B3. Zaznacz B2:B3 i przeciągnij uchwyt do B13. Excel sam doda przyrost 0,05. |
| 2e | Wpisz `10` w C2 i `900` w C13. Zaznacz C2:C13. Wybierz `Narzędzia główne` / `Edytowanie` / `Wypełnij` / `Seria danych`. Ustaw typ `Liniowa` i wartość końcową `900`. |
| 2f | Kliknij prawym przyciskiem nagłówek kolumny A i wybierz `Wstaw`. Wpisz `Miesiąc` w A1. |
| 2g | Wpisz `1.1.15` w A2 i `2.1.15` w A3. Zaznacz A2:A3 i przeciągnij uchwyt do A13. Excel doda po jednym miesiącu. |
| 2h | Kliknij prawym przyciskiem nagłówek wiersza 1 i wybierz `Wstaw`. Wpisz `Tabela wyników w Quidditchu` w A1. |

Uwaga do kolejności. Po kroku 2f kolumny przesuwają się w prawo.
Dane z A trafiają do B, z B do C, a z C do D.
Po kroku 2h wszystkie wiersze przesuwają się o jeden w dół.
Dlatego gotowa tabela zajmuje wiersze 2–14.

Jeśli Excel nie rozpozna daty, wpisz `2015-01-01` i `2015-02-01`.

### 3.3. Zadanie 3 – formatowanie komórek

| Krok | Co zrobić | Co jest w pliku |
| --- | --- | --- |
| 3a | Zaznacz A2:D2. Włącz pogrubienie i wyrównanie do środka. Zaznacz A3:A14 i wyrównaj daty do środka. | Nagłówek pogrubiony, cała tabela wyrównana do środka |
| 3b | Zaznacz B3:D14. Naciśnij `Ctrl` + `1`. Wybierz kategorię `Liczba` i ustaw 1 miejsce dziesiętne. | Format `0,0` |
| 3c | Zaznacz A3:A14. Naciśnij `Ctrl` + `1`. Wybierz kategorię `Niestandardowe` i wpisz format `mmmm`. | Pełna nazwa miesiąca, np. „październik" |
| 3d | Zaznacz A2:D14. W oknie `Ctrl` + `1` ustaw kartę `Obramowanie`. Zewnętrzne krawędzie: podwójna cienka linia. Wewnętrzne: pojedyncza cienka linia. Potem wypełnij A2:D2 na żółto. | Żółty nagłówek i ramka podwójna |
| 3e | Zaznacz A1:D1 i wybierz `Narzędzia główne` / `Wyrównanie` / `Scal i wyśrodkuj`. | Tytuł scalony nad tabelą |
| 3f | Ustaw szerokość kolumn B–D na 60 pikseli. Zaznacz A2:D2 i w oknie `Ctrl` + `1` włącz `Zawijaj tekst` oraz wyrównanie pionowe do środka. | Zawinięty nagłówek |

Uwaga do kroku 3f. Excel liczy piksele wzorem `px = szerokość * 7 + 5`.
Dlatego 60 pikseli to szerokość **7,86**.
Kolumna A ma szerokość **10**, aby zmieściła pełną nazwę miesiąca.

Gotowy arkusz `Potter` wygląda tak:

| Wiersz | A | B | C | D |
| --- | --- | --- | --- | --- |
| 1 | `Tabela wyników w Quidditchu` (scalone A1:D1) | | | |
| 2 | `Miesiąc` | `Potter Harry` | `Weasley Ron` | `Malfoy Draco` |
| 3–14 | 1.01.2015 – 1.12.2015, format `mmmm` | 1000 w każdym wierszu | 0,10 – 0,65, przyrost 0,05 | 10 – 900, seria liniowa |

### 3.4. Zadanie 4 – zarządzanie arkuszami

| Krok | Co zrobić | Co jest w pliku |
| --- | --- | --- |
| 4a | Wstaw dwa nowe arkusze. | Arkusze `Weasley` i `Malfoy` |
| 4b | Zmień nazwę arkusza z tabelą na `Potter`, a dwóch nowych na `Weasley` i `Malfoy`. Zaznacz karty z klawiszem `Ctrl` i zmień kolor wszystkich kart na czerwony. | Czerwone karty |
| 4c | Ustaw karty alfabetycznie, przeciągając je myszką. | Kolejność: `Malfoy`, `Potter`, `Weasley` |
| 4d | Usuń arkusz `Malfoy`. | Zostają `Potter` i `Weasley` |
| 4e | Otwórz plik roboczy `TI.1_Excel.E1+E2+E3_PlikRoboczy.xlsx` i przenieś do niego oba arkusze. | Kolejność: `BusinessPlan`, `Potter`, `Weasley` |

Karta arkusza `BusinessPlan` ma domyślny kolor. Ten arkusz powstał wcześniej.

---

## 4. Temat E2 – formuły, funkcje i adresy komórek

Zadania E2 wykonuje się w arkuszu `BusinessPlan`, w pliku roboczym.
Szablon ma gotowe dane i założenia.

| Komórka | Zawartość | Kolor |
| --- | --- | --- |
| C12 | Cena parówki | zielony |
| C13 | Koszty zmienne | niebieski |
| C14 | Koszty stałe | czerwony |
| C15 | Wzrost produkcji | szary |

Tabela 1 zajmuje wiersze 18–22. Kolumny C–N to miesiące od stycznia do grudnia.

### 4.1. Zadanie 1 – proste formuły dla stycznia

| Komórka | Formuła | Znaczenie |
| --- | --- | --- |
| C20 | `=C19*$C$12` | liczba parówek razy cena |
| C21 | `=C19*$C$13+$C$14` | koszt zmienny razy liczba parówek plus koszt stały |
| C22 | `=C20-C21` | wpływy minus koszty |

### 4.2. Zadanie 2 – adresy względne

Firma produkuje o 350 parówek więcej niż w poprzednim miesiącu.

1. Wpisz `=C19+350` w D19.
2. Zaznacz D19 i przeciągnij uchwyt wypełnienia do N19.
3. Adres `C19` jest względny, dlatego przesuwa się razem z kolumną.
4. Nie wpisuj każdej formuły ręcznie.

### 4.3. Zadanie 3 – adresy bezwzględne

Krok 3a. Sprawdź, czy można skopiować formuły stycznia na cały rok.

1. Skopiuj C20:C22 do zakresu D20:N22.
2. Zobaczysz błąd. Formuła bez znaku `$` przesuwa odwołanie do ceny.
   W komórce D20 powstaje `=D19*D12`, a D12 jest puste.

Krok 3b. Popraw formuły stycznia.

1. Wstaw `$` przed kolumną i wierszem: `$C$12`, `$C$13`, `$C$14`.
2. Zaznacz C20:C22 i przeciągnij uchwyt do N22.
3. Sprawdź wybrane komórki na pasku formuły. Adresy miesięcy się zmieniają, adresy założeń nie.

Krok 3c. W wierszu „Liczba parówek" nie może być liczby 350 z klawiatury.

1. Wpisz `=C19+$C$15` w D19.
2. Przeciągnij uchwyt do N19.
3. Styczeń zostaje zerem, bo to rozruch linii produkcyjnej.
4. Teraz zmiana komórki C15 zmienia cały rok.

Krok 3d. Druga tabela.

1. Zaznacz całą Tabelę 1, czyli `B18:N22`.
2. Naciśnij `Ctrl` + `C`.
3. Zaznacz komórkę `B24` i naciśnij `Ctrl` + `V`.
4. Tabela 2 zajmuje wiersze 24–28. Między tabelami zostaje jeden pusty wiersz.
5. Wpisz w C25 formułę `=N19+$C$15`.
6. Liczba parówek w styczniu w Tabeli 2 to grudzień z Tabeli 1 plus wzrost produkcji.

### 4.4. Zadanie 4 – funkcje

| Krok | Co zrobić |
| --- | --- |
| 4a | Wpisz `Suma` w O18 i `Średnia` w P18. |
| 4b | Zaznacz O19 i kliknij `Autosumowanie`. Powstanie `=SUM(C19:N19)`. |
| 4c | Zaznacz P19. Kliknij strzałkę obok `Autosumowanie` i wybierz `Średnia`. |
| 4d | Zmień formułę na `=ROUND(AVERAGE(C19:N19),0)`. To funkcja zagnieżdżona. |
| 4e | Skopiuj O19:P19 w dół do wiersza 22. Nadaj tym komórkom obramowanie takie samo jak w tabeli. |
| 4f | Skopiuj kolumny `Suma` i `Średnia` do Tabeli 2, czyli do zakresu O24:P28. |
| 4g | Odpowiedzi na pytania są w punkcie 7 tego dokumentu. |

Dlaczego w komórce widać `####`? Odpowiedź jest w punkcie 7.1.
Rozwiązanie: ustaw format `#,##0`, czyli liczbę całkowitą z separatorem tysięcy.

---

## 5. Temat E3 – formatowanie warunkowe i poprawność danych

Zadania E3 wykonuje się w arkuszu `BusinessPlan`, po zadaniu E2.
Wszystkie polecenia są w `Narzędzia główne` / `Style` / `Formatowanie warunkowe`.

### 5.1. Zadanie 1 – czcionka, krawędzie i wypełnienie

1a. Pogrubienie dla liczby parówek:

1. Zaznacz `C19:N19`.
2. Wybierz `Nowa reguła` / `Formatuj tylko komórki, które zawierają`.
3. Ustaw warunek: wartość większa niż `0`.
4. Kliknij `Formatuj` i włącz pogrubienie.

1b. Kolor dla wpływów:

1. Zaznacz `C20:N20`.
2. Dodaj regułę `Użyj formuły do określenia komórek do sformatowania`.
3. Wpisz formułę `=C20>AVERAGE($C$20:$N$20)` i ustaw zielone wypełnienie.
4. Dodaj drugą regułę z formułą `=C20<AVERAGE($C$20:$N$20)` i czerwonym wypełnieniem.

Formuła nie ma znaku `$` przy `C20`. Dzięki temu przesuwa się na kolejne komórki wiersza.

### 5.2. Zadanie 2 – paski danych, skala kolorów i ikony

| Krok | Co zrobić |
| --- | --- |
| 2a | Zaznacz `C21:N21` i wybierz `Paski danych`. |
| 2b | Zaznacz `C22:N22` i wybierz `Skale kolorów` / `Inne reguły`. Ustaw 3 kolory: czerwony, żółty i zielony. Punkt środkowy ustaw jako `Liczba` równą `0`, a nie domyślny percentyl. |
| 2c | Zaznacz `C22:N22` ponownie i dodaj `Zestawy ikon`, na przykład sygnalizację świetlną. |
| 2d | Wpisz `2` w C12. To niższa cena parówki. Obserwuj tabelę. |
| 2e | Obejrzyj pozostałe opcje: reguły dla tekstu, reguły dla dat, reguły dla wartości średnich. |

Cena w komórce C12 wynosi **2,00 €**. To wartość z kroku 2d.
Instrukcja każe ją wpisać i obserwować reguły. Nie każe wracać do starej ceny.

### 5.3. Zadanie 3 – poprawność danych

3a. Blokada ujemnej ceny:

1. Zaznacz `C12`.
2. Wybierz `Dane` / `Narzędzia danych` / `Poprawność danych`.
3. Ustaw `Zezwalaj` na `Liczba dziesiętna`, `Dane` na `większe niż`, `Wartość` na `0`.
4. Na karcie `Alert błędu` wpisz komunikat.
5. Wpisz `-1` w C12. Excel odrzuci tę wartość.

3b. Koszt zmienny niższy od ceny:

1. Zaznacz `C13`.
2. Wybierz `Poprawność danych` i ustaw `Zezwalaj` na `Niestandardowe`.
3. Wpisz formułę `=C13<C12`.
4. Wpisz `5` w C13. Excel odrzuci wartość, bo 5 jest większe od ceny.

3c. Odpowiedź na pytanie o kryterium `Lista` jest w punkcie 7 tego dokumentu.

---

## 6. Wyniki liczbowe

Założenia: cena 2,00 €, koszt zmienny 1,45 €, koszt stały 3 000 €.

Tabela 1:

| Wiersz | Styczeń | Grudzień | Suma | Średnia |
| --- | --- | --- | --- | --- |
| Liczba parówek (szt.) | 0 | 3 850 | 23 100 | 1 925 |
| Wpływy ze sprzedaży (zł) | 0 | 7 700 | 46 200 | 3 850 |
| Koszty | 3 000 | 8 583 | 69 495 | 5 791 |
| Zysk | −3 000 | −883 | −23 295 | −1 941 |

Tabela 2 startuje od 4 200 parówek w styczniu:

| Wiersz | Styczeń | Grudzień | Suma | Średnia |
| --- | --- | --- | --- | --- |
| Liczba parówek (szt.) | 4 200 | 8 050 | 73 500 | 6 125 |
| Wpływy ze sprzedaży (zł) | 8 400 | 16 100 | 147 000 | 12 250 |
| Koszty | 9 090 | 14 673 | 142 575 | 11 881 |
| Zysk | −690 | 1 428 | 4 425 | 369 |

Koszty w grudniu to dokładnie 8 582,50 zł. Format liczbowy pokazuje 8 583 zł.
Kolumna Średnia używa funkcji `ROUND`, dlatego pokazuje liczby całkowite.

---

## 7. Odpowiedzi na pytania z instrukcji

### 7.1. Dlaczego w komórce widać `####`?

W komórce nie mieści się gotowy napis z liczbą.
Kolumny Suma i Średnia mają w szablonie format `0,0000`.
Suma 23 100 daje napis `23100,0000`, czyli 10 znaków.
Kolumna O ma 9 znaków szerokości, więc Excel pokazuje `####`.
Rozwiązanie: format `#,##0` albo szersza kolumna.

### 7.2. Jaka funkcja liczy minimum w wierszu?

Funkcja `MIN`, na przykład `=MIN(C19:N19)`.
W tym wierszu wynik to 0, bo styczeń ma zero parówek.
Podobne funkcje to `MAX`, `MEDIAN` i `MODE`.

### 7.3. Przykłady funkcji z trzech kategorii

| Kategoria | Funkcja | Co robi |
| --- | --- | --- |
| Tekstowe | `CONCATENATE` (`ZŁĄCZ.TEKSTY`) | Skleja tekst z kilku części w jeden napis. |
| Tekstowe | `TEXT` (`TEKST`) | Zamienia liczbę na tekst o wybranym formacie. |
| Data i godzina | `TODAY` (`DZIŚ`) | Zwraca dzisiejszą datę. |
| Data i godzina | `MONTH` (`MIESIĄC`) | Zwraca numer miesiąca z podanej daty. |
| Matem. i tryg. | `ROUND` (`ZAOKRĄGLIJ`) | Zaokrągla liczbę do wybranej liczby miejsc. |
| Matem. i tryg. | `SUM` (`SUMA`) | Dodaje liczby z podanego zakresu. |

Przykłady na danych z arkusza:

* `=MIN(C19:N19)` zwraca 0.
* `=MAX(C19:N19)` zwraca 3 850.
* `=CONCATENATE("Średnia liczba parówek: ",TEXT(P19,"#,##0"))` zwraca gotowy napis.
* `=TEXT(TODAY(),"yyyy-mm-dd")` zwraca dzisiejszą datę jako tekst.
* `=ROUND(AVERAGE(C22:N22),2)` zwraca średni zysk z dokładnością do groszy.

### 7.4. Jak działają adresy względne i bezwzględne?

| Zapis | Nazwa | Zachowanie przy kopiowaniu |
| --- | --- | --- |
| `C19` | względny | Przesuwa się razem z komórką. |
| `$C19` | mieszany | Kolumna stoi, wiersz się przesuwa. |
| `C$19` | mieszany | Wiersz stoi, kolumna się przesuwa. |
| `$C$19` | bezwzględny | Nic się nie przesuwa. |

Klawisz `F4` przełącza te cztery warianty podczas edycji formuły.

### 7.5. Jak stosować kryterium `Lista`?

1. Wpisz dozwolone wartości w osobnych komórkach, na przykład `2,99`, `2,50`, `2,00`.
2. Zaznacz komórkę, w której ma działać lista.
3. Wybierz `Dane` / `Narzędzia danych` / `Poprawność danych`.
4. Ustaw `Zezwalaj` na `Lista` i wskaż zakres z wartościami.
5. Zaznacz pole `Lista rozwijana`.

W komórce pojawia się strzałka. Po kliknięciu wybierasz wartość z listy.
Użytkownik nie może wpisać wartości spoza listy.
