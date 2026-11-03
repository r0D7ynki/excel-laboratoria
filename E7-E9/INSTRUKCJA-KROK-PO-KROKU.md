# Instrukcja krok po kroku – TI.1 Excel E7, E8 i E9

Ten dokument opisuje wykonanie wszystkich poleceń z pliku PDF.
Każdy krok podaje ścieżkę na wstążce, adres komórki i wynik.
Kolejność jest taka sama jak w poleceniach.

**Spis treści**

0. Zanim zaczniesz
1. E7 – organizowanie danych w tabele
2. E8 – tabele i wykresy przestawne
3. E9 – osadzanie obiektów w programie Word
4. Kontrola wyników
5. Najczęstsze błędy
6. Pliki

---

## 0. Zanim zaczniesz

### 0.1. Pliki

| Plik | Rola |
| --- | --- |
| `TI.1_Excel.E7+E8+E9_PlikRoboczy.xlsx` | Plik roboczy i plik zaliczeniowy. |
| `TI.1_Excel.E7+E8+E9_InstrukcjaLab.pdf` | Treść poleceń. |
| `RAPORT-E9-Hog-Dog.docx` | Gotowy raport Word. |

Zrób kopię pliku roboczego.
Pracuj na kopii.
Plik źródłowy z PDF ma tylko arkusz `EWIDENCJA`.

### 0.2. Oznaczenia

* `Dane / Sortowanie i filtrowanie / Sortuj` – kolejne elementy wstążki.
* `Ctrl+1` – skrót klawiszowy.
* `C12` – adres komórki.
* `=SUMA(C19:N19)` – formuła do wpisania.

### 0.3. Język formuł

Formuły w tym dokumencie są po polsku i mają średniki.
Tak wygląda wpisywanie w polskim Excelu.
Plik `.xlsx` zapisuje formuły po angielsku i z przecinkami.
To normalne. Format pliku ma własny język.
Na przykład `=JEŻELI(A1>0;1;0)` zapisze się jako `=IF(A1>0,1,0)`.

### 0.4. Kolejność pracy

1. Najpierw wykonaj E7.
2. Potem E8. Użyj tabeli z arkusza `EWIDENCJA`.
3. Na końcu wykonaj E9.

Nie dodawaj wiersza z 31.12.2024 przed wykonaniem konspektu.
Konspekt ma powstać z danych 384 wierszy.
Wiersz dodasz w zadaniu 2f.

---

## 1. Temat E7 – organizowanie danych w tabele

### 1.1. Kopia arkusza EWIDENCJA

1. Kliknij prawym przyciskiem myszy kartę `EWIDENCJA`.
2. Wybierz `Przenieś lub kopiuj`.
3. Zaznacz `Utwórz kopię`.
4. Kliknij `OK`.
5. Kliknij dwukrotnie nową kartę.
6. Wpisz nazwę `KONSPEKT` i naciśnij `Enter`.

Sprawdź, że arkusz `KONSPEKT` ma 384 wiersze danych.
Zakres danych to `A1:D385`.

### 1.2. Sortowanie wielopoziomowe

1. Zaznacz zakres `A1:D385` w arkuszu `KONSPEKT`.
2. Wybierz `Dane / Sortowanie i filtrowanie / Sortuj`.
3. Zaznacz `Moje dane mają nagłówki`.
4. W polu `Sortuj według` wybierz `Kraj`.
5. W polu `Sortuj według` obok wybierz `Wartości`.
6. Wybierz `Od A do Z`.
7. Kliknij `Dodaj poziom`.
8. Wybierz `Miasto`, `Wartości`, `Od A do Z`.
9. Kliknij `Dodaj poziom`.
10. Wybierz `Data`, `Wartości`, `Od najstarszych do najnowszych`.
11. Kliknij `OK`.

Sprawdź pierwsze wiersze:

| Wiersz | Kraj | Miasto | Data |
| --- | --- | --- | --- |
| 2 | Hogwart | Gryffindor | 07.01.2024 |
| 3 | Hogwart | Gryffindor | 14.01.2024 |
| 4 | Hogwart | Gryffindor | 21.01.2024 |

Miasta w kraju Hogwart występują w kolejności:
Gryffindor, Hufflepuff, Ravenclaw, Slytherin.
W kraju Ringland:
Gondor, Isengard, Mordor, Rohan.

### 1.3. Suma częściowa dla kraju

1. Zaznacz `A1:D385`.
2. Wybierz `Dane / Suma częściowa`.
3. W polu `Przy każdej zmianie w` wybierz `Kraj`.
4. W polu `Użyj funkcji` wybierz `Suma`.
5. W polu `Dodaj sumę częściową do` zaznacz tylko `Kwota`.
6. Kliknij `OK`.

Excel wstawi wiersz `Hogwart Suma częściowa` i `Ringland Suma częściowa`.
Suma kraju Hogwart wynosi `1 832 430`.
Suma kraju Ringland wynosi `850 143`.

### 1.4. Suma częściowa dla miasta

1. Zaznacz ponownie zakres danych. Uwaga: nie zaznaczaj wierszy sum.
2. Wybierz `Dane / Suma częściowa`.
3. W polu `Przy każdej zmianie w` wybierz `Miasto`.
4. W polu `Użyj funkcji` wybierz `Suma`.
5. W polu `Dodaj sumę częściową do` zaznacz tylko `Kwota`.
6. **Odznacz** pole `Zamień bieżące sumy częściowe`.
7. Kliknij `OK`.

Teraz są dwa poziomy sum:

* suma miasta,
* suma kraju.

Sumy miast:

| Miasto | Kwota |
| --- | ---: |
| Gryffindor | 729 745 |
| Hufflepuff | 379 080 |
| Ravenclaw | 191 390 |
| Slytherin | 532 215 |
| Gondor | 305 588 |
| Isengard | 270 513 |
| Mordor | 104 176 |
| Rohan | 169 866 |

Suma końcowa wynosi `2 682 573`.

### 1.5. Konspekt – przyciski 1, 2, 3 i +/−

1. Spójrz na lewą krawędź okna.
2. Kliknij przycisk `1`. Widoczne będą tylko sumy krajów i suma końcowa.
3. Kliknij przycisk `2`. Widoczne będą sumy miast.
4. Kliknij przycisk `3`. Widoczne będą wszystkie szczegóły.
5. Kliknij `+` przy kraju. Rozwinie się jeden kraj.
6. Kliknij `−` przy kraju. Kraj się zwinie.

Sprawdź, że każdy kraj i każde miasto występują tylko raz.
Jeśli kraj występuje dwa razy, dane nie były posortowane.
Wróć do punktu 1.2.

### 1.6. Formuły czy wartości

1. Kliknij komórkę sumy miasta, np. `D50`.
2. Spójrz na pasek formuły.
3. Zobaczysz `=SUMA.CZĘŚCIOWA(9;D2:D49)`.
4. Kliknij sumę kraju, np. `D198`.
5. Zobaczysz `=SUMA.CZĘŚCIOWA(9;D2:D197)`.
6. Kliknij sumę końcową `D396`.
7. Zobaczysz `=SUMA.CZĘŚCIOWA(9;D2:D395)`.

W komórkach są formuły, nie wartości.

Test z punktu 1e:

1. Znajdź dowolny wiersz z datą `28.12.2024`.
2. Zmień jego `Kwota` na `9999`.
3. Sprawdź sumę miasta, sumę kraju i sumę końcową.
4. Wszystkie sumy się zmieniły.
5. Cofnij zmianę klawiszem `Ctrl+Z`.

### 1.7. Czerwona karta arkusza EWIDENCJA

1. Przejdź do arkusza `EWIDENCJA`.
2. Kliknij prawym przyciskiem myszy kartę.
3. Wybierz `Kolor karty`.
4. Wybierz kolor czerwony.
5. Kliknij poza menu.

### 1.8. Tabela SprzedażTygodniowa

1. W arkuszu `EWIDENCJA` kliknij komórkę `A1`.
2. Wybierz `Wstawianie / Tabele / Tabela`.
3. Sprawdź zakres `A1:D385`.
4. Zaznacz `Mój zakres ma nagłówki`.
5. Kliknij `OK`.
6. Wybierz `Narzędzia tabel / Projektowanie`.
7. W polu `Nazwa tabeli` wpisz `SprzedażTygodniowa`.
8. Naciśnij `Enter`.
9. W galerii `Style tabeli` wybierz `TableStyleMedium9`.
10. Zaznacz `Paski wierszy`.

Tabela ma teraz zakres `A1:D385` i nazwę `SprzedażTygodniowa`.

### 1.9. Kolumna Kwartał

1. Kliknij komórkę `E1`.
2. Wpisz `Kwartał`.
3. Naciśnij `Enter`. Excel rozszerzy tabelę o jedną kolumnę.
4. Kliknij `E2`.
5. Wpisz formułę:

```
=ZAOKR.GÓRA(MIESIĄC(C2)/3;0)
```

6. Naciśnij `Enter`.
7. Excel powinien wypełnić całą kolumnę.
8. Jeśli tego nie zrobi, przeciągnij uchwyt wypełnienia do `E385`.

Sprawdź wartości:

| Data | Kwartał |
| --- | ---: |
| 07.01.2024 | 1 |
| 07.04.2024 | 2 |
| 07.07.2024 | 3 |
| 07.10.2024 | 4 |

### 1.10. Kolumna Miesiąc

Ta kolumna jest potrzebna do raportu miesięcznego w E8.

1. Kliknij `F1`.
2. Wpisz `Miesiąc`.
3. Naciśnij `Enter`. Tabela rozszerzy się do kolumny `F`.
4. Kliknij `F2`.
5. Wpisz formułę:

```
=TEKST(C2;"mm mmm")
```

6. Naciśnij `Enter`.
7. Przeciągnij formułę do `F385`.

W komórkach pojawią się wartości `01 sty`, `02 lut`, aż do `12 gru`.
Format `mm mmm` daje numer miesiąca i skrót nazwy.
Numer na początku zapewnia właściwą kolejność miesięcy.

### 1.11. Nowy wiersz 31.12.2024

1. Kliknij ostatnią komórkę tabeli `F385`.
2. Naciśnij klawisz `Tab`. Excel doda nowy wiersz 386.
3. Wpisz dane:

| Kolumna | Wartość |
| --- | --- |
| `Kraj` | `Ringland` |
| `Miasto` | `Isengard` |
| `Data` | `31.12.2024` |
| `Kwota` | `500` |
| `Kwartał` | formuła `=ZAOKR.GÓRA(MIESIĄC(C386)/3;0)` |
| `Miesiąc` | formuła `=TEKST(C386;"mm mmm")` |

4. Sprawdź, że tabela ma zakres `A1:F386`.
5. Sprawdź, że `Kwartał` wynosi `4`, a `Miesiąc` wynosi `12 gru`.

Suma w tabeli wynosi teraz `2 683 073`.

### 1.12. Kolumna Średnia dzienna

Ta kolumna jest przykładem ze wskazówki w pliku źródłowym.
Leży poza tabelą.

1. Kliknij `G1`.
2. Wpisz `Średnia dzienna`.
3. Kliknij `G2`.
4. Wpisz `=D2/7`.
5. Naciśnij `Enter`.
6. Przeciągnij formułę do `G386`.
7. Ustaw format `#,##0.00`.

Wskazówka w pliku mówi o formule w komórce `G3`.
Gdy komórka należy do tabeli, Excel zapisuje tą samą formułę jako odwołanie strukturalne, np. `=SprzedażTygodniowa[[#This Row],[Kwota]]/7`.
Tutaj kolumna `G` leży poza tabelą, więc formuła pozostaje zwykła.

### 1.13. Kopia PRZEFILTROWANE

1. Kliknij prawym przyciskiem myszy kartę `EWIDENCJA`.
2. Wybierz `Przenieś lub kopiuj`.
3. Zaznacz `Utwórz kopię`.
4. Kliknij `OK`.
5. Zmień nazwę nowej karty na `PRZEFILTROWANE`.

Kopia powstaje po dodaniu wiersza 31.12.2024.
Dlatego ma 385 wierszy danych.

### 1.14. Sortowanie po kwocie

1. Kliknij strzałkę filtra w nagłówku `Kwota`.
2. Wybierz `Sortuj od największego do najmniejszego`.
3. Sprawdź, że pierwszy wiersz ma największą kwotę.

### 1.15. Filtr kraju i daty

1. Kliknij strzałkę filtra w nagłówku `Kraj`.
2. Odznacz `(Zaznacz wszystko)`.
3. Zaznacz tylko `Hogwart`.
4. Kliknij `OK`.
5. Kliknij strzałkę filtra w nagłówku `Data`.
6. Wybierz `Filtry daty / Między`.
7. Wpisz datę od `01.07.2024`.
8. Wpisz datę do `31.08.2024`.
9. Kliknij `OK`.

Widoczne są 32 wiersze.
Wszystkie dotyczą Hogwartu i wakacji letnich.

### 1.16. Wiersz sumy

1. Kliknij dowolną komórkę tabeli.
2. Wybierz `Narzędzia tabel / Projektowanie`.
3. Zaznacz `Wiersz sumy`.
4. W wierszu sumy kliknij komórkę pod `Kwota`.
5. Wybierz `Suma`.
6. W komórce pod `Kwartał` wybierz `Brak`.
7. W komórce pod `Miesiąc` wybierz `Brak`.

Suma wynosi `333 584`.
To suma tylko widocznych wierszy.
Formuła to `=SUMA.CZĘŚCIOWA(109;SprzedażTygodniowa_Filtr[Kwota])`.

### 1.17. Odpowiedzi do tematu E7

* W komórkach podsumowań są formuły `SUMA.CZĘŚCIOWA`.
* Zmiana kwoty z 28.12 zmienia sumy.
* Kraje i miasta nie powtarzają się, bo dane są posortowane.
* Wiersz sumy sumuje tylko widoczne wiersze.
* Tabela nazywa się `SprzedażTygodniowa`.
* Karta arkusza `EWIDENCJA` jest czerwona.

---

## 2. Temat E8 – tabele i wykresy przestawne

### 2.1. Zasada pracy z tabelami przestawymi

1. Kliknij dowolną komórkę tabeli `SprzedażTygodniowa`.
2. Wybierz `Wstawianie / Tabele / Tabela przestawna`.
3. Wybierz `Nowy arkusz`.
4. Kliknij `OK`.
5. Zmień nazwę arkusza na nazwę raportu.
6. Przeciągnij pola z panelu po prawej stronie.

Każdy raport ma osobny arkusz.
Nazwa arkusza jest taka sama jak nazwa raportu.

### 2.2. Raport „Sprzedaż w miastach”

1. Utwórz nową tabelę przestawną.
2. Zmień nazwę arkusza na `Sprzedaż w miastach`.
3. Przeciągnij `Miasto` do obszaru `Wiersze`.
4. Przeciągnij `Kwota` do obszaru `Wartości`.
5. Sprawdź, że operacja to `Suma`.
6. Przeciągnij `Kraj` do obszaru `Filtry`.
7. Ustaw format wartości `#,##0`.

Wynik:

| Miasto | Suma z Kwota |
| --- | ---: |
| Gondor | 305 588 |
| Gryffindor | 729 745 |
| Hufflepuff | 379 080 |
| Isengard | 271 013 |
| Mordor | 104 176 |
| Ravenclaw | 191 390 |
| Rohan | 169 866 |
| Slytherin | 532 215 |
| Razem | 2 683 073 |

### 2.3. Raport „Kwartalnie w miastach”

1. Utwórz nową tabelę przestawną.
2. Zmień nazwę arkusza na `Kwartalnie w miastach`.
3. Przeciągnij `Kwartał` do `Wiersze`.
4. Przeciągnij `Miasto` pod `Kwartał`.
5. Przeciągnij `Kwota` do `Wartości`.
6. Przeciągnij `Kraj` do `Filtry`.
7. Wybierz `Projektowanie / Układ raportu / Pokaż w formie konspektu`.
8. Kliknij `+` i `−` przy kwartale, aby sprawdzić zwijanie.

Sumy kwartałów:

| Kwartał | Kwota |
| --- | ---: |
| 1 | 470 170 |
| 2 | 605 787 |
| 3 | 765 600 |
| 4 | 841 516 |

### 2.4. Raport „Miasta kwartałami”

1. Utwórz nową tabelę przestawną.
2. Zmień nazwę arkusza na `Miasta kwartałami`.
3. Przeciągnij `Kraj` do `Wiersze`.
4. Przeciągnij `Miasto` pod `Kraj`.
5. Przeciągnij `Kwartał` pod `Miasto`.
6. Przeciągnij `Kwota` do `Wartości`.
7. Ustaw układ `Pokaż w formie konspektu`.

Sprawdź sumy krajów:

| Kraj | Kwota |
| --- | ---: |
| Hogwart | 1 832 430 |
| Ringland | 850 643 |
| Razem | 2 683 073 |

### 2.5. Raport „Tabelarycznie”

1. Utwórz nową tabelę przestawną.
2. Zmień nazwę arkusza na `Tabelarycznie`.
3. Przeciągnij `Miasto` do `Wiersze`.
4. Przeciągnij `Kwartał` do `Kolumny`.
5. Przeciągnij `Kwota` do `Wartości`.
6. Przeciągnij `Kraj` do `Filtry`.
7. Wybierz `Projektowanie / Układ raportu / Pokaż w formie tabelarycznej`.

Wynik to tabela miast i kwartałów.
Wiersz `Razem` pokazuje sumy kwartałów.
Kolumna `Razem` pokazuje sumy miast.

### 2.6. Raport „Miasta w miesiącach”

1. Utwórz nową tabelę przestawną.
2. Zmień nazwę arkusza na `Miasta w miesiącach`.
3. Przeciągnij `Miesiąc` do `Wiersze`.
4. Przeciągnij `Kraj` do `Kolumny`.
5. Przeciągnij `Miasto` pod `Kraj` w obszarze `Kolumny`.
6. Przeciągnij `Kwota` do `Wartości`.
7. Ustaw układ `Pokaż w formie tabelarycznej`.

Wiersze to 12 miesięcy.
Kolumny to kraje i miasta.
Suma w wierszu `12 gru` wynosi `232 414`.

Uwaga: w pliku użyto kolumny `Miesiąc`.
Gdybyś używał pełnej daty, kliknij komórkę z datą i wybierz `Analiza / Grupowanie / Grupuj pole`.
Zaznacz `Miesiące` i kliknij `OK`.
Kolumna `Miesiąc` daje ten sam wynik i działa w każdym programie.

### 2.7. Udział procentowy w raporcie „Kwartalnie w miastach”

1. Przejdź do arkusza `Kwartalnie w miastach`.
2. Przeciągnij `Kwota` drugi raz do obszaru `Wartości`.
3. Kliknij strzałkę przy nowym polu `Suma z Kwota`.
4. Wybierz `Ustawienia pola wartości`.
5. Kliknij `Pokaż wartości jako`.
6. Wybierz `% sumy końcowej`.
7. Kliknij `OK`.
8. Zmień nazwę pola na `% z całości`.
9. Ustaw format `0,00%`.

W kolumnie obok sumy pojawi się udział procentowy.
Suma na dole wynosi `100,00%`.

### 2.8. Dodatkowa tabela przestawna

1. Utwórz nową tabelę przestawną.
2. Zmień nazwę arkusza na `Cwiczenia`.
3. Przeciągnij `Kraj` do `Wiersze`.
4. Przeciągnij `Miasto` do `Kolumny`.
5. Przeciągnij `Kwota` do `Wartości`.
6. Przeciągnij `Kwartał` do `Filtry`.
7. Ustaw układ tabelaryczny.

Ta tabela służy do ćwiczenia zmiany układu i filtrów.

### 2.9. Wartości czy formuły

1. Kliknij dowolną komórkę tabeli przestawnej.
2. Spójrz na pasek formuły.
3. Zobaczysz wartość, nie formułę.
4. Zmień kwotę w arkuszu `EWIDENCJA`.
5. Wróć do tabeli przestawnej.
6. Tabela się nie zmieni.
7. Kliknij `Analiza / Odśwież`.
8. Tabela pokaże nową wartość.

Tabela przestawna nie odświeża się sama w trakcie pracy.
Plik ma włączone odświeżanie przy otwarciu.
Po ponownym otwarciu raport jest aktualny.

### 2.10. Wykres przestawny w arkuszu „Tabelarycznie”

1. Kliknij komórkę wewnątrz tabeli przestawnej `Tabelarycznie`.
2. Wybierz `Narzędzia tabel przestawnych / Analiza / Wykres przestawny`.
3. Wybierz `Kolumnowy skumulowany`.
4. Kliknij `OK`.
5. Ustaw tytuł `Sprzedaż kwartalna w miastach`.
6. Dodaj tytuły osi: `Miasto` i `Kwota`.
7. Przenieś wykres pod tabelę.

Wykres pokazuje sprzedaż miast w czterech kwartałach.

### 2.11. Wykres przestawny bez gotowej tabeli

1. Przejdź do arkusza `EWIDENCJA`.
2. Kliknij komórkę wewnątrz tabeli `SprzedażTygodniowa`.
3. Wybierz `Wstawianie / Tabele / Tabela przestawna`.
4. Kliknij dolną część przycisku i wybierz `Wykres przestawny`.
5. Wybierz `Nowy arkusz`.
6. Kliknij `OK`.
7. Zmień nazwę arkusza na `WykresPrzestawny`.
8. Przeciągnij `Miasto` do `Wiersze`.
9. Przeciągnij `Kwota` do `Wartości`.
10. Wybierz wykres słupkowy.
11. Dodaj tytuł `Sprzedaż według miast`.

Wykres ma przyciski pól.
Możesz zmieniać układ i filtr.
Zmiana wykresu zmienia tabelę pod wykresem.

### 2.12. Odpowiedzi do tematu E8

* Komórki tabeli przestawnej pokazują wartości, nie formuły.
* Tabela nie aktualizuje się sama. Trzeba kliknąć `Odśwież`.
* Wykres przestawny można filtrować i przestawiać.
* Wykres i tabela pod nim korzystają z jednego źródła.
* Raport miesięczny używa kolumny `Miesiąc` zamiast grupowania dat.

---

## 3. Temat E9 – osadzanie obiektów w programie Word

### 3.1. Tekst raportu

1. Otwórz Microsoft Word.
2. Wpisz tytuł `Raport dla zarządu Hog-Dog sp. z o.o.`.
3. Wpisz kilka zdań o sprzedaży.
4. Wstaw tabelę wyników z arkusza `Sprzedaż w miastach`.
5. Wstaw wykres z arkusza `Tabelarycznie`.

Raport w repozytorium ma 4 strony.
Zawiera tabelę i dwa wykresy.

### 3.2. Wstawianie nowego arkusza Excela

1. W Wordzie wybierz `Wstawianie / Tekst / Obiekt`.
2. Wybierz `Arkusz programu Microsoft Excel`.
3. Kliknij `OK`.
4. Wpisz kilka liczb do komórek.
5. Przeciągnij narożnik obiektu, aby go powiększyć.
6. Kliknij poza obiektem.

Word wstawił obiekt OLE.
Obiekt jest osadzony w dokumencie.

### 3.3. Edycja obiektu

1. Kliknij dwukrotnie obiekt arkusza.
2. Word pokaże wstążkę Excela.
3. Wpisz nową liczbę.
4. Kliknij poza obiektem.
5. Wstążka Worda wróci.

Obiekt poddaje się edycji w miejscu.

### 3.4. Dwa wykresy

1. Wybierz `Wstawianie / Ilustracje / Wykres`.
2. Wybierz wykres kolumnowy.
3. Kliknij `OK`.
4. Wpisz dane w oknie arkusza.
5. Zamknij okno danych.
6. Wstaw drugi wykres przez `Wstawianie / Tekst / Obiekt / Wykres`.
7. Porównaj oba wykresy.

Pierwszy wykres jest wykresem Worda.
Drugi jest obiektem OLE programu Excel.
Po dwukliku drugiego wykresu otwiera się Excel.

### 3.5. Arkusz przez tabelę

1. Wybierz `Wstawianie / Tabele / Tabela`.
2. Wybierz `Arkusz kalkulacyjny programu Excel`.
3. Wpisz dane.
4. Kliknij poza tabelą.

To ten sam typ obiektu OLE co w punkcie 3.2.
Różni się tylko przycisk na wstążce.

### 3.6. Zapis i pliki osadzone

1. Zapisz dokument jako `RAPORT-E9-Hog-Dog.docx`.
2. Sprawdź, że powstał jeden plik.
3. Zmień rozszerzenie kopii na `.zip`.
4. Otwórz archiwum.
5. Wejdź do katalogu `word/embeddings`.

Osadzone skoroszyty leżą w środku pliku `.docx`.
Nie powstają osobne pliki `.xlsx`.

### 3.7. Wklejanie tabeli przestawnej

1. W Excelu zaznacz tabelę przestawną `Sprzedaż w miastach`.
2. Skopiuj ją (`Ctrl+C`).
3. W Wordzie wklej zwykle (`Ctrl+V`).
4. Kliknij tabelę jeden raz.
5. Kliknij tabelę dwa razy.

Zwykłe wklejenie daje tabelę Worda albo obraz.
Nie ma połączenia z Excelem.

### 3.8. Wklej specjalnie jako obiekt

1. Ustaw kursor dwa wiersze poniżej wklejonej tabeli.
2. Wybierz `Narzędzia główne / Schowek / Wklej`.
3. Kliknij dolną część przycisku.
4. Wybierz `Wklej specjalnie`.
5. Wybierz `Arkusz programu Microsoft Excel – obiekt`.
6. Kliknij `OK`.

Powstał obiekt OLE.
Dwuklik otwiera Excel z kopią arkusza.

### 3.9. Edycja obiektu osadzonego

1. Kliknij dwukrotnie drugą tabelę.
2. Zmień kilka komórek.
3. Kliknij poza obiektem.
4. Sprawdź, że zmiany są widoczne.
5. Otwórz arkusz źródłowy w Excelu.
6. Sprawdź te same komórki.

Zmiany w obiekcie osadzonym nie zmieniają pliku źródłowego.
Obiekt jest niezależną kopią.

### 3.10. Wklej łącze

1. Skopiuj zakres komórek w Excelu.
2. W Wordzie wybierz `Wklej specjalnie`.
3. Wybierz `Wklej łącze`.
4. Kliknij `OK`.
5. Kliknij dwukrotnie obiekt.
6. Zmień kilka komórek.

Obiekt łączony pokazuje dane z pliku źródłowego.
Zmiana w Excelu aktualizuje obiekt w Wordzie.
Zmiana w obiekcie zmienia plik źródłowy.

### 3.11. Wykres jako obiekt

1. Skopiuj wykres w Excelu.
2. W Wordzie wybierz `Wklej` albo `Wklej specjalnie / Wykres programu MS Excel – obiekt`.
3. Kliknij wykres jeden raz.
4. Kliknij wykres dwa razy.
5. Zmień wartość w danych źródłowych.
6. Sprawdź, co się zaktualizowało.

Obiekt wykresu po dwukliku otwiera Excel.
Zmiana danych źródłowych zmienia wykres.
Przy łączu zmiana pliku źródłowego odświeża wykres.

### 3.12. Inne metody

* `Wstawianie / Obiekt / Utwórz z pliku` – osadzenie albo łącze.
* Przeciągnięcie pliku do dokumentu.
* `Wklej specjalnie` z wybranym formatem.
* `Wstawianie / Obiekt / Utwórz nowy` – nowy obiekt.
* Obiekty OLE można wstawiać także do arkusza Excela.

### 3.13. Odpowiedzi do tematu E9

* Obiekt osadzony to kopia. Nie zmienia pliku źródłowego.
* Obiekt połączony korzysta z pliku źródłowego.
* `Wstaw / Wykres` tworzy wykres Worda.
* `Wstaw / Obiekt / Wykres` tworzy obiekt OLE.
* `Wstaw / Tabela / Arkusz Excela` daje ten sam obiekt co `Wstaw / Obiekt`.
* Po zapisie jest jeden plik `.docx`.
* Osadzone skoroszyty są w `word/embeddings`.

---

## 4. Kontrola wyników

### 4.1. Skoroszyt

| Kontrola | Wartość |
| --- | --- |
| Liczba arkuszy | 10 |
| Karta `EWIDENCJA` | czerwona |
| Tabela `EWIDENCJA` | `SprzedażTygodniowa`, `A1:F386` |
| Suma w `EWIDENCJA` | 2 683 073 |
| `KONSPEKT` – suma końcowa | 2 682 573 |
| `KONSPEKT` – poziomy konspektu | 1, 2, 3 |
| `PRZEFILTROWANE` – widoczne wiersze | 32 |
| `PRZEFILTROWANE` – suma | 333 584 |
| `Sprzedaż w miastach` – suma | 2 683 073 |
| `Kwartalnie w miastach` – Q4 | 841 516 |
| `Tabelarycznie` – wykres | kolumnowy skumulowany |
| `WykresPrzestawny` – wykres | słupkowy |

### 4.2. Raport Word

| Kontrola | Wartość |
| --- | --- |
| Liczba stron | 4 |
| Tabela wyników | 8 miast + razem |
| Wykresy | 2 |
| Opis metod osadzania | tak |
| Tabela porównania metod | tak |

### 4.3. Otwieranie pliku

1. Otwórz skoroszyt w Excelu.
2. Sprawdź, że nie ma komunikatu o naprawie.
3. Otwórz raport w Wordzie.
4. Sprawdź, że tabela i wykresy są widoczne.

---

## 5. Najczęstsze błędy

### 5.1. Konspekt

* Brak sortowania przed sumami częściowymi.
  Skutek: kraj lub miasto występuje kilka razy.
* Zaznaczenie `Zamień bieżące sumy częściowe` przy drugim sumowaniu.
  Skutek: znikają sumy krajów.
* Zaznaczenie całej kolumny zamiast zakresu danych.
  Skutek: Excel dodaje sumy do pustych wierszy.

### 5.2. Tabela

* Spacja w nazwie tabeli.
  Nazwa `Sprzedaż Tygodniowa` jest niepoprawna. Użyj `SprzedażTygodniowa`.
* Data wpisana jako tekst.
  Data musi być prawdziwą datą. Inaczej filtr daty nie działa.
* Brak formuły w nowym wierszu.
  Skopiuj formułę `Kwartał` i `Miesiąc` do wiersza 386.

### 5.3. Tabele przestawne

* Tabela przestawna nie odświeża się sama.
  Kliknij `Analiza / Odśwież`.
* Zły zakres źródła.
  Źródło musi obejmować tabelę `A1:F386` albo nazwę `SprzedażTygodniowa`.
* Brak kolumny `Miesiąc`.
  Raport miesięczny nie pokaże miesięcy.

### 5.4. Word

* Łącze do pliku, który się przeniósł.
  Word nie znajdzie źródła. Wskaż plik ponownie.
* Zwykłe `Wklej` zamiast `Wklej specjalnie`.
  Powstanie tabela albo obraz, a nie obiekt OLE.
* Brak osadzonego pliku w `word/embeddings`.
  Sprawdź, czy wybrano `Obiekt`, a nie `Obraz`.

---

## 6. Pliki

* [Skoroszyt](./TI.1_Excel.E7+E8+E9_PlikRoboczy.xlsx)
* [Raport Word](./RAPORT-E9-Hog-Dog.docx)
* [Treść poleceń PDF](./TI.1_Excel.E7+E8+E9_InstrukcjaLab.pdf)
* [Opis zadań i wyniki](./README.md)
