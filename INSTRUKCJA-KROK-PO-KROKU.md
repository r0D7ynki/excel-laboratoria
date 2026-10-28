# Instrukcja krok po kroku – TI.1 Arkusz kalkulacyjny (E1–E6)

Ten dokument opisuje całą drogę do gotowych skoroszytów.
Każdy krok podaje ścieżkę na wstążce, komórkę i wynik, który ma się pojawić.
Kolejność jest taka sama jak w plikach PDF z poleceniami.

**Spis treści**

0. Zanim zaczniesz
1. E1 – tabela wyników w Quidditchu
2. E2 – formuły i adresy
3. E3 – formatowanie warunkowe i poprawność danych
4. E4 – składniki i koszty
5. E5 – analiza warunkowa
6. E6 – wykresy
7. Skróty klawiszowe
8. Najczęstsze błędy
9. Kontrola wyników
10. Gdzie szukać odpowiedzi

---

## 0. Zanim zaczniesz

### 0.1. Pliki

| Część | Plik roboczy | Instrukcja |
| --- | --- | --- |
| E1–E3 | `E1-E3/TI.1_Excel.E1+E2+E3_PlikRoboczy.xlsx` | `E1-E3/TI.1_Excel.E1+E2+E3_InstrukcjaLab.pdf` |
| E4–E6 | `E4-E6/TI.1_Excel.E4+E5+E6_PlikRoboczy.xlsx` | `E4-E6/TI.1_Excel.E4+E5+E6_InstrukcjaLab.pdf` |

### 0.2. Oznaczenia w tym dokumencie

* `Narzędzia główne / Liczba / Więcej…` – kolejne elementy wstążki.
* `Ctrl+1` – skrót klawiszowy.
* `C12` – adres komórki.
* `=SUMA(C19:N19)` – formuła do wpisania.

### 0.3. Język formuł

Formuły w tym dokumencie są po polsku i mają średniki.
Tak wygląda wpisywanie w polskim Excelu.
Plik `.xlsx` zapisuje je po angielsku i z przecinkami.
To normalne, format pliku ma własny język.
Na przykład `=JEŻELI(A1>0;1;0)` zapisze się jako `=IF(A1>0,1,0)`.

### 0.4. Pierwsza zasada

Zapisuj plik często (`Ctrl+S`).
Do zaliczenia oddajesz plik `.xlsx`, nie PDF.

---

# Część I – E1, E2, E3

## 1. Temat E1 – tabela wyników w Quidditchu

Zadania 1–3 wykonujesz w **nowym, pustym skoroszycie**.
Dopiero na końcu przenosisz arkusze do pliku roboczego.

### 1.1. Zadanie 1 – szablony (ćwiczenie)

1. Wybierz `Plik / Nowy`.
2. Otwórz kilka szablonów.
3. Przełączaj okna klawiszami `Alt+Tab`.
4. Przełączaj skoroszyty klawiszami `Ctrl+Tab` i `Ctrl+Shift+Tab`.
5. Obejrzyj wstążkę `Widok`.

Ten punkt nie zostawia śladu w pliku.

### 1.2. Zadanie 2 – dane i szerokości kolumn

1. Wpisz `Potter Harry` w `A1`, `Weasley Ron` w `B1`, `Malfoy Draco` w `C1`.
2. Dopasuj szerokość kolumn A–C na trzy sposoby:
   * przeciągnij krawędź nagłówka kolumny,
   * kliknij dwukrotnie krawędź (autodopasowanie),
   * `Narzędzia główne / Komórki / Formatuj / Szerokość kolumny`.
3. Wpisz `1000` w `A2`.
4. Przeciągnij uchwyt wypełnienia do `A13`.
5. Wpisz `0,1` w `B2` i `0,15` w `B3`.
6. Zaznacz `B2:B3` i przeciągnij uchwyt do `B13`. Excel doda po 0,05.
7. Wpisz `10` w `C2` i `900` w `C13`.
8. Zaznacz `C2:C13`.
9. Wybierz `Narzędzia główne / Edytowanie / Wypełnij / Seria danych`.
10. Ustaw typ `Liniowa`, wartość końcową `900`, kliknij `OK`.
11. Kliknij prawym przyciskiem nagłówek kolumny `A` i wybierz `Wstaw`.
12. Wpisz `Miesiąc` w `A1`.
13. Wpisz `1.1.15` w `A2` i `2.1.15` w `A3`.
    Jeśli Excel potraktuje to jako tekst, wpisz `2015-01-01` i `2015-02-01`.
14. Zaznacz `A2:A3` i przeciągnij uchwyt do `A13`. Powstaną miesiące.
15. Kliknij prawym przyciskiem nagłówek wiersza `1` i wybierz `Wstaw`.
16. Wpisz `Tabela wyników w Quidditchu` w `A1`.

**Uwaga do adresów.** Po kroku 11 kolumny przesunęły się w prawo.
Po kroku 15 wszystkie wiersze przesunęły się w dół.
Dlatego gotowa tabela zajmuje `A2:D14`, a nie `A1:C13`.

### 1.3. Zadanie 3 – formatowanie tabeli

1. Zaznacz `A2:D2`. Włącz `Pogrubienie` i `Wyśrodkuj`.
2. Zaznacz `A3:A14`. Włącz `Wyśrodkuj`.
3. Zaznacz `B3:D14`. Naciśnij `Ctrl+1`.
4. Wybierz kategorię `Liczba` i ustaw `1` miejsce dziesiętne.
5. Zaznacz `A3:A14`. Naciśnij `Ctrl+1`.
6. Wybierz kategorię `Niestandardowe` i wpisz format `mmmm`.
7. Zaznacz `A2:D14`. Naciśnij `Ctrl+1`.
8. Karta `Obramowanie`:
   * `Zewnętrzne` – linia podwójna cienka,
   * `Wewnętrzne` – linia pojedyncza cienka.
9. Zaznacz `A2:D2`. Ustaw wypełnienie `Żółty`.
10. Zaznacz `A1:D1`. Wybierz `Narzędzia główne / Wyrównanie / Scal i wyśrodkuj`.
11. Zaznacz kolumny `B:D`. Ustaw szerokość `60` pikseli.
    W oknie `Szerokość kolumny` wpisz `7,86`.
12. Zaznacz `A2:D2`. Włącz `Zawijaj tekst` i wyrównanie pionowe `Do środka`.

**Sprawdź.** Kolumna A pokazuje `styczeń`, `luty` itd.
Nagłówek jest żółty i pogrubiony.
Cała tabela ma podwójną ramkę zewnętrzną.

### 1.4. Zadanie 4 – arkusze

1. Wstaw dwa arkusze: `Shift+F11` dwa razy.
2. Kliknij dwukrotnie kartę arkusza z tabelą. Wpisz `Potter`.
3. Zmień nazwy nowych arkuszy na `Weasley` i `Malfoy`.
4. Zaznacz karty z klawiszem `Ctrl` i kliknij prawym przyciskiem.
5. Wybierz `Kolor karty` i ustaw `Czerwony`.
6. Przeciągnij karty myszką do kolejności `Malfoy`, `Potter`, `Weasley`.
7. Kliknij prawym przyciskiem kartę `Malfoy` i wybierz `Usuń`.
8. Otwórz plik `TI.1_Excel.E1+E2+E3_PlikRoboczy.xlsx`.
9. Kliknij prawym przyciskiem kartę `Potter`.
10. Wybierz `Przenieś lub kopiuj`.
11. Wskaż skoroszyt `TI.1_Excel.E1+E2+E3_PlikRoboczy.xlsx`.
    W polu `Przed arkuszem` wybierz `(przenieś na koniec)`.
    Nie zaznaczaj `Utwórz kopię`.
12. Powtórz krok 9–11 dla arkusza `Weasley`.

**Wynik.** Kolejność kart: `BusinessPlan`, `Potter`, `Weasley`.
Karty `Potter` i `Weasley` są czerwone.

---

## 2. Temat E2 – formuły i adresy

Zadania wykonujesz w arkuszu `BusinessPlan`, w pliku roboczym.
Tabela 1 zajmuje wiersze 18–22.
Kolumny C–N to miesiące od stycznia do grudnia.

### 2.1. Zadanie 1 – formuły dla stycznia

1. W `C20` wpisz `=C19*C12`.
   To wpływy: liczba parówek razy cena.
2. W `C21` wpisz `=C19*C13+C14`.
   To koszty: zmienne razy sztuki plus stałe.
3. W `C22` wpisz `=C20-C21`.
   To zysk.

Formuły są na razie względne.
W zadaniu 3 poprawisz je na bezwzględne.

### 2.2. Zadanie 2 – adresy względne

1. W `D19` wpisz `=C19+350`.
2. Przeciągnij uchwyt wypełnienia do `N19`.
3. Excel sam przesuwa `C19` na `D19`, `E19` i tak dalej.

Liczba parówek rośnie o 350 co miesiąc.
Nie wpisuj dwunastu formuł ręcznie.

### 2.3. Zadanie 3 – adresy bezwzględne

1. Skopiuj `C20:C22` do `D20:N22`.
2. Zobaczysz błąd. Formuła `=C19*C12` w kolumnie D zmienia się w `=D19*D12`.
   Komórka `D12` jest pusta.
3. Popraw `C20` na `=C19*$C$12`.
4. Popraw `C21` na `=C19*$C$13+$C$14`.
5. `C22` zostaje bez zmian: `=C20-C21`.
6. Zaznacz `C20:C22` i przeciągnij do `N22`.
7. Kliknij kilka komórek w wierszu 20. Na pasku formuły sprawdź adresy.
   Adres miesiąca się zmienia, adresy założeń stoją.
8. W `D19` wpisz `=C19+$C$15`.
9. Przeciągnij do `N19`. Styczeń zostaje zerem.
10. Zaznacz `B18:N22`. Naciśnij `Ctrl+C`.
11. Zaznacz `B24` i naciśnij `Ctrl+V`.
12. W `C25` wpisz `=N19+$C$15`.

**Wynik.** Tabela 2 zajmuje wiersze 24–28.
Styczeń w Tabeli 2 to grudzień z Tabeli 1 plus wzrost produkcji.

### 2.4. Zadanie 4 – funkcje

1. Wpisz `Suma` w `O18` i `Średnia` w `P18`.
2. Zaznacz `O19`. Kliknij `Narzędzia główne / Edytowanie / Autosumowanie`.
3. W komórce pojawi się `=SUMA(C19:N19)`.
4. Widać `####` zamiast liczby. Kolumna jest za wąska, a format ma 4 miejsca.
5. Naciśnij `Ctrl+1` i ustaw format `#,##0`. Poszerz kolumnę `O`.
6. Zaznacz `P19`. Kliknij strzałkę obok `Autosumowanie` i wybierz `Średnia`.
7. Popraw formułę na `=ZAOKRĄGLIJ(ŚREDNIA(C19:N19);0)`.
8. Zaznacz `O19:P19` i przeciągnij do `P22`.
9. Ustaw obramowanie `O19:P22` takie samo jak w tabeli.
10. Skopiuj `O18:P22` i wklej w `O24`.
11. Odpowiedzi na pytania są w `E1-E3/README.md`, punkcie 7.

**Kontrola.** `O19` = 23 100, `P19` = 1 925, `O22` = −23 295.

---

## 3. Temat E3 – formatowanie warunkowe i poprawność

Wszystko robisz w arkuszu `BusinessPlan`, po zadaniu E2.

### 3.1. Zadanie 1 – czcionka i wypełnienie

1. Zaznacz `C19:N19`.
2. Wybierz `Narzędzia główne / Style / Formatowanie warunkowe / Nowa reguła`.
3. Wybierz `Formatuj tylko komórki, które zawierają`.
4. Ustaw: `Większe niż`, wartość `0`.
5. Kliknij `Formatuj`, karta `Czcionka`, włącz `Pogrubienie`.
6. Zaznacz `C20:N20`.
7. `Formatowanie warunkowe / Nowa reguła / Użyj formuły do określenia komórek do sformatowania`.
8. Wpisz `=C20>ŚREDNIA($C$20:$N$20)`.
9. Ustaw wypełnienie `Zielone` i kliknij `OK`.
10. Dodaj drugą regułę z formułą `=C20<ŚREDNIA($C$20:$N$20)` i wypełnieniem `Czerwonym`.

Brak `$` przy `C20` jest celowy. Dzięki temu formuła działa w każdej kolumnie.

### 3.2. Zadanie 2 – paski, skala, ikony

1. Zaznacz `C21:N21`.
2. `Formatowanie warunkowe / Paski danych / Niebieski pasek danych`.
3. Zaznacz `C22:N22`.
4. `Formatowanie warunkowe / Skale kolorów / Inne reguły`.
5. Wybierz skalę 3-kolorową: czerwony, żółty, zielony.
6. Dla punktu środkowego wybierz `Typ: Liczba` i wpisz `0`.
7. Zaznacz `C22:N22` ponownie.
8. `Formatowanie warunkowe / Zestawy ikon / Sygnalizacja świetlna`.
9. Wpisz `2` w `C12`. To niższa cena parówki.
10. Obserwuj, jak zmieniają się kolory w tabeli.

Cena `2` zostaje w pliku. Instrukcja nie każe wracać do starej ceny.

### 3.3. Zadanie 3 – poprawność danych

1. Zaznacz `C12`.
2. Wybierz `Dane / Narzędzia danych / Poprawność danych`.
3. Ustaw `Zezwalaj: Liczba dziesiętna`, `Dane: Większe niż`, `Wartość: 0`.
4. Karta `Alert błędu`: wpisz `Cena parówki musi być większa od 0`.
5. Wpisz `-1` w `C12`. Excel odrzuci wartość.
6. Zaznacz `C13`.
7. `Poprawność danych`, `Zezwalaj: Niestandardowe`.
8. Wpisz formułę `=C13<C12`.
9. Karta `Alert błędu`: wpisz `Koszty zmienne muszą być niższe od ceny`.
10. Wpisz `5` w `C13`. Excel odrzuci wartość.

**Odpowiedź na pytanie o listę.** Opis jest w `E1-E3/README.md`, punkcie 7.5.

---

# Część II – E4, E5, E6

## 4. Temat E4 – składniki i koszty

Wszystko robisz w arkuszu `Wykonalność`.

### 4.1. Zadanie 1a – zabezpieczenie dzielenia

1. Zaznacz `C7`. To koszt składników jednej parówki.
2. Wpisz `=JEŻELI(C31<>0;I48/C31;"Nie można dzielić przez 0")`.
3. W `C31` nie ma jeszcze sumy, więc formuła pokaże komunikat.
4. Po wypełnieniu Tabeli 3 wróci liczba.

### 4.2. Zadanie 1b–1c – waga składników

1. Zaznacz `D19`. To smocze sadło w styczniu.
2. Wpisz `=$C19*$C$9/1000*D$13`.
3. Przeciągnij uchwyt w prawo do `H19`.
4. Zaznacz `D19:H19` i przeciągnij w dół do wiersza `30`.

Dwa ruchy myszką wypełniają całą Tabelę 3.
Adresy mają specjalne znaczenie:

| Zapis | Co się dzieje przy kopiowaniu |
| --- | --- |
| `$C19` | kolumna C stoi, wiersz się zmienia |
| `$C$9` | nic się nie zmienia (waga parówki) |
| `D$13` | wiersz 13 stoi, kolumna się zmienia |

### 4.3. Zadanie 1d – sumy Tabeli 3

1. Zaznacz `C31:H31`.
2. Kliknij `Autosumowanie`.
3. Powstaną sumy roczne pod każdą kolumną.

### 4.4. Zadanie 1e–1f – koszt składników

1. Zaznacz `D36`.
2. Wpisz `=D19*D$14`.
3. Przeciągnij w prawo do `H36`.
4. Zaznacz `D36:H36` i przeciągnij w dół do wiersza `47`.
5. Uzupełnij kolumnę `Razem` (jeśli jest pusta): w `I36` wpisz `=SUMA(D36:H36)`.
6. Przeciągnij `I36` w dół do `I47`.
7. W `I48` wpisz `=SUMA(I36:I47)`.

### 4.5. Zadanie 1g – rachunek wyników

Tabela 6 stoi w kolumnach H–I, wiersze 51–57.
Uzupełnij lub sprawdź formuły:

| Komórka | Formuła | Znaczenie |
| --- | --- | --- |
| `I51` | `=C31*C6` | przychód: sztuki razy cena |
| `I52` | `=I48` | koszty składników |
| `I53` | `=12*C8` | koszty stałe razy 12 |
| `I54` | `=I48+I53` | koszty całkowite |
| `I55` | `=I51-I54` | dochód |
| `I56` | `=15000` | spłata kredytu |
| `I57` | `=I55-I56` | końcowy ZYSK |

**Kontrola.** Przy cenie 3,00 € i wadze 100 g `I57` wynosi −2 019 €.

### 4.6. Zadanie 2 – funkcje wyszukiwania

1. W `D51` wpisz `=MAX(I36:I47)`.
   To najwyższy miesięczny koszt składników.
2. W `D52` wpisz `=INDEKS(B36:B47;PODAJ.POZYCJĘ(D51;I36:I47;0))`.
   To nazwa miesiąca z tym kosztem.
3. W `D53` wpisz `=WYSZUKAJ.PIONOWO(D52;B19:C30;2;FAŁSZ)`.
   To liczba parówek w tym miesiącu.

**Dlaczego nie WYSZUKAJ?**
Funkcja `WYSZUKAJ` (`LOOKUP`) wymaga danych posortowanych rosnąco.
Kolumna `I36:I47` nie jest posortowana.
Wynik byłby przypadkowy.
Dlatego używamy `INDEKS` + `PODAJ.POZYCJĘ`.

---

## 5. Temat E5 – analiza warunkowa

### 5.1. Zadanie 1a – próg rentowności metodą prób

1. Zaznacz `C9`. To waga parówki, startowo 100 g.
2. Wpisuj kolejno `99`, `98`, `98,5`.
3. Po każdej zmianie patrz na `I57`.
4. Przy wadze około `98,39` zysk wynosi 0.

To ćwiczenie ma pokazać zależność. Nie zostawia śladu w pliku.

### 5.2. Zadanie 1b – Szukaj wyniku

1. Zaznacz komórkę `I57`.
2. Wybierz `Dane / Analiza warunkowa / Szukaj wyniku`.
3. Ustaw:
   * `Ustaw komórkę:` `$I$57`,
   * `Na wartość:` `10000`,
   * `Zmieniając komórkę:` `$C$9`.
4. Kliknij `OK`.
5. Excel znajdzie wagę około `90,386261` g.
6. Wynik zostanie wpisany w `C9`.

**Sprawdź.** `I57` pokazuje `10 000 €`, a `C9` około `90,4` g.

### 5.3. Zadanie 2a – Tabela 7 (jedna zmienna)

1. W `B61` wpisz `waga`.
2. W `C61` wpisz `=I57`.
3. W `B62` wpisz `85`.
4. W `B63` wpisz `=B62+1`.
5. Przeciągnij `B63` do `B82`. Ostatnia waga to `105`.
6. Zaznacz `B61:C82`.
7. Wybierz `Dane / Analiza warunkowa / Tabela danych`.
8. W polu `Komórka wejściowa kolumny:` wpisz `$C$9`.
9. Pole `Komórka wejściowa wiersza` zostaw puste.
10. Kliknij `OK`.

**Sprawdź.** `C62` = 16 733,85 €, `C77` = −2 019,00 €.
Komórki `C62:C82` przeliczają się same. Nie edytuj ich ręcznie.

### 5.4. Zadanie 2b – Tabela 8 (dwie zmienne)

1. W `E61` wpisz `=I57`.
2. W `F61` wpisz `70`.
3. W `G61` wpisz `=F61+5` i przeciągnij do `N61`.
   Ostatnia wartość to `110`.
4. W `E62` wpisz `2,5`.
5. W `E63` wpisz `=E62+0,1` i przeciągnij do `E72`.
   Ostatnia cena to `3,5`.
6. Zaznacz `E61:N72`.
7. Wybierz `Dane / Analiza warunkowa / Tabela danych`.
8. `Komórka wejściowa wiersza:` wpisz `$E$14`.
   To koszt mrówczego mięsa.
9. `Komórka wejściowa kolumny:` wpisz `$C$6`.
   To cena parówki.
10. Kliknij `OK`.

**Sprawdź.** `H67` = 10 000 €, `F72` = 49 484,81 €, `N62` = −50 454,42 €.

### 5.5. Zadanie 3a–3b – scenariusze

1. Wybierz `Dane / Analiza warunkowa / Menedżer scenariuszy`.
2. Kliknij `Dodaj`.
3. Nazwa: `Bieżący`.
4. Kliknij pole `Komórki zmieniane:`.
   Zaznacz `C6`, potem przytrzymaj `Ctrl` i zaznacz `D14` oraz `E14`.
5. Wartości: `3`, `15`, `80`. Kliknij `OK`.
6. Kliknij `Dodaj` i utwórz `Optymistyczny`: `3,5`, `10`, `70`.
7. Kliknij `Dodaj` i utwórz `Pesymistyczny`: `2,5`, `20`, `90`.
8. Zaznacz scenariusz na liście i kliknij `Pokaż`.
9. Obserwuj `I57` po każdym scenariuszu.

**Sprawdź.** Bieżący 10 000 €, Optymistyczny 57 348,41 €, Pesymistyczny −37 348,41 €.

### 5.6. Zadanie 3c – raport scenariuszy

1. Kliknij `Podsumowanie`.
2. `Typ raportu:` wybierz `Podsumowanie scenariuszy`.
   Nie wybieraj raportu tabeli przestawnej.
3. W polu `Komórki wynikowe:` wskaż zakres `D48:H48`.
4. Przytrzymaj `Ctrl` i wskaż komórkę `I57`.
5. Kliknij `OK`.
6. Excel utworzy arkusz `Podsumowanie scenariuszy`.

Arkusz pokazuje koszty składników i końcowy ZYSK dla trzech scenariuszy.

### 5.7. Zadanie 3d – nazwy zdefiniowane

1. Zaznacz `C6`.
2. Wybierz `Formuły / Nazwy zdefiniowane / Definiuj nazwę`.
3. Nazwa: `Cena_1_parowki`. Zakres: arkusz `Wykonalność`.
4. Kliknij `OK`.
5. Powtórz dla `D14` z nazwą `Koszt_Smoczego_sadla`.
6. Powtórz dla `E14` z nazwą `Koszt_Mrowczego_miesa`.
7. Otwórz `Menedżer scenariuszy` i ponownie zrób `Podsumowanie`.
8. W nowym raporcie zamiast adresów `$C$6`, `$D$14`, `$E$14` są nazwy.

To jest różnica, o którą pyta instrukcja.

---

## 6. Temat E6 – wykresy

Wykresy wstawiasz w arkuszu `Wykonalność`, pod tabelami.

### 6.1. Wykres kolumnowy (zadanie 6.1a)

1. Zaznacz `B35:H47`.
   To nagłówki i zielone wnętrze Tabeli 4, bez kolumny `Razem`.
2. Wybierz `Wstawianie / Wykresy / Wstaw wykres kolumnowy / Grupowany`.
3. Kliknij wykres i wybierz `Narzędzia wykresu / Projektowanie / Dodaj element wykresu`.
4. Włącz `Tytuł wykresu` i wpisz `Koszty składników w czasie (styczeń–grudzień)`.
5. Włącz `Tytuły osi`. Oś pozioma: `Miesiąc`. Oś pionowa: `Koszt [€]`.
6. Włącz `Legenda` i ustaw pozycję `Dolna`.
7. Kliknij serię `Makulatura`, wybierz `Formatuj serię danych`.
8. Ustaw `Wypełnienie / Kolor` na `Żółty`.
9. Przeciągnij wykres tak, aby lewy górny róg wypadł w `A85`.

**Sprawdź.** Wykres ma 5 serii, 12 miesięcy i żółtą Makulaturę.

### 6.2. Wykres kołowy (zadanie 6.1b)

1. Zaznacz `D48:H48`. To roczne sumy składników.
2. Wybierz `Wstawianie / Wykresy / Wstaw wykres kołowy / Kołowy`.
3. Kliknij wykres i wybierz `Projektowanie / Zaznacz dane`.
4. W polu `Etykiety osi poziomej (kategorii)` kliknij `Edytuj`.
5. Wpisz `=Wykonalność!$D$35:$H$35` i kliknij `OK`.
6. Dodaj tytuł `Struktura rocznych kosztów składników`.
7. Ustaw legendę po prawej stronie.
8. Kliknij wykres, potem `Dodaj element wykresu / Etykiety danych / Więcej opcji`.
9. Włącz `Nazwa kategorii` i `Wartość procentowa`.
10. W sekcji `Format liczby` wybierz `Procent` i 1 miejsce dziesiętne.
11. Przeciągnij wykres do `A106`.

**Sprawdź.** Etykiety pokazują np. `Mrówcze mięso 74,2%` i `Woda 0,1%`.

### 6.3. Wykres przebiegu w czasie (zadanie 6.2a)

Typ `Wykresy przebiegu w czasie` w Excelu wymaga tabeli przestawnej.
Dlatego w zwykłym arkuszu robi się wykres liniowy po miesiącach.
Efekt jest ten sam: pokazuje zmianę wartości w czasie.

1. Zaznacz `I35:I47`. To kolumna `Razem`.
2. Wybierz `Wstawianie / Wykresy / Wstaw wykres liniowy / Liniowy z markerami`.
3. Kliknij wykres i wybierz `Projektowanie / Zaznacz dane`.
4. W polu `Etykiety osi poziomej (kategorii)` kliknij `Edytuj`.
5. Wpisz `=Wykonalność!$B$36:$B$47` i kliknij `OK`.
6. Dodaj tytuł `Przebieg kosztów całkowitych w czasie (styczeń–grudzień)`.
7. Oś pozioma: `Miesiąc`. Oś pionowa: `Koszt [€]`.
8. Legenda na dole.
9. Przeciągnij wykres do `A132`.

**Sprawdź.** Linia rośnie od stycznia do czerwca.
Potem spada w lipcu i sierpniu, a od września znowu rośnie.

### 6.4. Uwagi do wykresów

* Każdy wykres ma tytuł, legendę i tytuły osi.
* Wykres kołowy nie ma osi, więc nie ma tytułów osi.
* Zmiana danych w tabeli odświeża wykresy po `F9`.
* Aby zmienić zakres serii, kliknij wykres i wybierz `Projektowanie / Zaznacz dane`.

---

## 7. Skróty, które oszczędzają czas

| Skrót | Działanie |
| --- | --- |
| `Ctrl+1` | okno `Formatuj komórki` |
| `F4` | przełącza `A1`, `$A$1`, `A$1`, `$A1` podczas edycji formuły |
| `Alt+=` | Autosumowanie |
| `Ctrl+Shift+F3` | nazwy z zaznaczenia |
| `F9` | przeliczenie arkusza |
| `Ctrl+Shift+F9` | pełne przeliczenie |
| `Shift+F11` | nowy arkusz |
| `Ctrl+Page Down` | następny arkusz |
| `Ctrl+Alt+F9` | przeliczenie wszystkich skoroszytów |

---

## 8. Najczęstsze błędy

| Objaw | Przyczyna | Poprawka |
| --- | --- | --- |
| `####` w komórce | za wąska kolumna albo zły format | poszerz kolumnę, ustaw `#,##0` |
| `#DZIEL/0!` | dzielenie przez pustą komórkę | użyj `JEŻELI`, jak w `C7` |
| złe liczby po skopiowaniu | adres względny zamiast bezwzględnego | dodaj `$` i naciśnij `F4` |
| data pokazuje się jako tekst | zły separator daty | wpisz `2015-01-01` |
| wykres pokazuje `Seria1` | brak zaznaczonych nagłówków | zaznacz zakres razem z wierszem nagłówka |
| tabela danych się nie zmienia | brak przeliczenia | naciśnij `F9` |
| scenariusz nie zmienia komórek | zły zakres komórek zmienianych | zaznacz `C6`, `D14` i `E14` w polu `Komórki zmieniane` |

---

## 9. Kontrola wyników

### 9.1. Skoroszyt E1–E3

| Komórka | Wartość |
| --- | --- |
| `C12` – cena | 2,00 € |
| `O19` – suma parówek | 23 100 |
| `P19` – średnia parówek | 1 925 |
| `O20` – suma wpływów | 46 200 |
| `O21` – suma kosztów | 69 495 |
| `O22` – suma zysku | −23 295 |
| `C25` – styczeń w Tabeli 2 | 4 200 |
| `N25` – grudzień w Tabeli 2 | 8 050 |
| `O28` – suma zysku Tabeli 2 | 4 425 |

### 9.2. Skoroszyt E4–E6

| Komórka | Wartość |
| --- | --- |
| `C9` – waga po Szukaj wyniku | 90,386261 g |
| `C31` – liczba parówek w roku | 58 000 |
| `D51` – maks. koszt miesięczny | 15 586,21 € |
| `D52` – miesiąc | Cze |
| `D53` – liczba parówek | 8 000 |
| `I48` – roczny koszt składników | 113 000,00 € |
| `I57` – ZYSK | 10 000,00 € |
| `C77` – ZYSK dla 100 g | −2 019,00 € |
| `H67` – ZYSK dla 3,00 € i 80 €/kg | 10 000,00 € |

Jeśli liczby się nie zgadzają, sprawdź kolejność:
najpierw Tabela 3, potem Tabela 4, na końcu Tabela 6.
Błąd w jednej komórce przenosi się na cały rachunek.

---

## 10. Gdzie szukać odpowiedzi na pytania

Pytania z instrukcji mają odpowiedzi w plikach README:

* `E1-E3/README.md`, punkt 7 – `####`, funkcje `MIN`, `MAX`, adresy, lista.
* `E4-E6/README.md`, punkty 2 i 3 – `WYSZUKAJ`, tabele danych, scenariusze.
