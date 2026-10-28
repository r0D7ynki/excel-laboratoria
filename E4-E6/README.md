# TI.1 – Arkusz kalkulacyjny: zadania E4, E5 i E6

Kompletne rozwiązanie laboratorium dla spółki Hog-Dog sp. z o.o.
Plik ze skoroszytem: [`TI.1_Excel.E4+E5+E6_PlikRoboczy.xlsx`](./TI.1_Excel.E4+E5+E6_PlikRoboczy.xlsx)
Treść poleceń: [`TI.1_Excel.E4+E5+E6_InstrukcjaLab.pdf`](./TI.1_Excel.E4+E5+E6_InstrukcjaLab.pdf)

---

## 1. Zawartość skoroszytu

| Arkusz | Zawartość |
| --- | --- |
| `Wykonalność` | Model finansowy: tabele 1–6 (E4), tabele 7–8 (E5), menedżer scenariuszy i wykresy (E6). |
| `Podsumowanie scenariuszy` | Raport Excela dla scenariuszy: Bieżący, Optymistyczny i Pesymistyczny. |

**Stan końcowy pliku**

* Waga parówki: **90,386261 g**. To wynik narzędzia Szukaj wyniku z zadania E5.
* Roczny koszt składników: **113 000,00 €**.
* Końcowy ZYSK: **10 000,00 €**.

> Punktem wyjścia zadania E4 była waga 100 g. Przy tej wadze ZYSK wynosił **−2 019 €**.
> Zadanie E5 zmieniło wagę na 90,386 g. Dlatego w gotowym pliku widzisz ZYSK 10 000 €.
> Wartość −2 019 € nadal występuje w Tabeli 7 dla wagi 100 g.
> Komentarz w komórce `I57` przypomina oba wyniki.

Plik zawiera zapisane wyniki formuł. Liczby są widoczne od razu po otwarciu.
Nie trzeba czekać na przeliczenie arkusza.

---

## 2. Temat E4 – adresowanie mieszane i funkcje wyszukiwania

### 2.1. Zadanie 1 – formuły i adresowanie mieszane

| Krok | Komórka | Formuła w pliku | Znaczenie |
| --- | --- | --- | --- |
| 1a | `C7` | `=IF(C31<>0,I48/C31,"Nie można dzielić przez 0")` | Zabezpiecza dzielenie przez zero. Gdy `C31` ma wartość 0, pokazuje komunikat. |
| 1b | `D19` | `=$C19*$C$9/1000*D$13` | Waga smoczego sadła w styczniu w kg. |
| 1c | `D19:H30` | ta sama formuła, wypełniona w prawo i w dół | Waga wszystkich składników w 12 miesiącach. |
| 1d | `C31:H31` | `=SUM(...)` | Sumy roczne pod Tabelą 3. |
| 1e | `D36` | `=D19*D$14` | Koszt smoczego sadła w styczniu w €. |
| 1f | `D36:H47` | ta sama formuła, wypełniona w prawo i w dół | Koszt wszystkich składników w 12 miesiącach. |
| 1g | `I57` | `=I55-I56` | Końcowy ZYSK. |

**Dlaczego adresowanie jest mieszane**

* `$C19` – kolumna `C` jest zablokowana, wiersz zmienia się w dół.
* `$C$9` – adres bezwzględny wagi parówki.
* `D$13` – wiersz 13 jest zablokowany, kolumna zmienia się w prawo.
* Dzięki temu całą tabelę wypełnia się dwoma przeciągnięciami.

**Wynik dla wagi 100 g (punkt wyjścia E4)**

| Pozycja | Wartość |
| --- | --- |
| Liczba parówek w roku (`C31`) | 58 000 |
| Koszt składników na parówkę (`C7`) | 2,1555 € |
| Roczny koszt składników (`I48`) | 125 019,00 € |
| ZYSK (`I57`) | **−2 019,00 €** |

**Wynik po zadaniu E5 (waga 90,386 g)**

| Pozycja | Wartość |
| --- | --- |
| Koszt składników na parówkę (`C7`) | 1,9483 € |
| Roczny koszt składników (`I48`) | 113 000,00 € |
| ZYSK (`I57`) | **10 000,00 €** |

### 2.2. Zadanie 2 – funkcje wyszukiwania

| Podpunkt | Komórka | Formuła w pliku | Wynik |
| --- | --- | --- | --- |
| 2a | `D51` | `=MAX(I36:I47)` | **15 586,21 €** |
| 2b | `D52` | `=INDEX(B36:B47,MATCH(D51,I36:I47,0))` | **Cze** |
| 2c | `D53` | `=VLOOKUP(D52,B19:C30,2,0)` | **8 000** |

**Odpowiedź na pytanie z instrukcji**

Funkcja `WYSZUKAJ` (`LOOKUP`) **nie jest tu odpowiednia**.
Wymaga ona posortowanego wektora rosnąco.
Kolumna `I36:I47` nie jest posortowana.
Wynik funkcji `LOOKUP` jest więc niewiarygodny.

Poprawne rozwiązanie to `INDEKS` + `PODAJ.POZYCJĘ`.
W nowszym Excelu działa też `X.WYSZUKAJ`.
Komentarz w komórce `D52` przypomina tę zasadę.

Przy wadze 100 g wartości Tabeli 5 były większe.
Maksymalny koszt miesięczny wynosił wtedy 17 244,00 €.
Po zmianie wagi w E5 maksimum spadło do 15 586,21 €.

---

## 3. Temat E5 – analiza warunkowa („co-jeśli”)

### 3.1. Zadanie 1 – Szukaj wyniku

**3.1a. Próg rentowności**

Próg rentowności to waga około **98,385 g**.
Przy tej wadze ZYSK wynosi 0 €.

**3.1b. Szukaj wyniku dla ZYSKU 10 000 €**

| Ustawienie | Wartość |
| --- | --- |
| Ustaw komórkę | `I57` (ZYSK) |
| Na wartość | `10000` |
| Zmieniając komórkę | `C9` (Waga parówki) |
| Wynik | **90,386261 g** |

Wynik jest zapisany w komórce `C9`.
Komentarz w `C9` podaje wynik i próg rentowności.
Arkusz nie zawiera już formuły celu.
Narzędzie Szukaj wyniku zapisuje tylko wynik końcowy.

### 3.2. Zadanie 2 – tabele danych

**Tabela 7 – wpływ wagi parówki na zysk**

| Element | Wartość |
| --- | --- |
| Zakres | `B61:C82` |
| Komórka wynikowa | `C61` = `=I57` |
| Komórka wejściowa | `C9` (waga parówki) |
| Typ tabeli | jednoczynnikowa, kolumnowa |

Wybrane wyniki:

| Waga | ZYSK |
| --- | --- |
| 85 g | 16 733,85 € |
| 90 g | 10 482,90 € |
| 95 g | 4 231,95 € |
| 98 g | 481,38 € |
| 99 g | −768,81 € |
| 100 g | **−2 019,00 €** |
| 105 g | −8 269,95 € |

**Tabela 8 – wpływ ceny parówki i kosztu mięsa na zysk**

| Element | Wartość |
| --- | --- |
| Zakres | `E61:N72` |
| Komórka narożna | `E61` = `=I57` |
| Komórka wejściowa wiersza | `E14` (koszt mrówczego mięsa) |
| Komórka wejściowa kolumny | `C6` (cena parówki) |

Wybrane wyniki:

| Cena parówki | Koszt mięsa | ZYSK |
| --- | --- | --- |
| 3,00 € | 80 €/kg | **10 000,00 €** |
| 3,50 € | 70 €/kg | **49 484,81 €** |
| 2,50 € | 110 €/kg | **−50 454,42 €** |

Tabele 7 i 8 są prawdziwymi tabelami danych.
W komórkach `C62` i `F62` są formuły `TABLE(...)`.
Nie edytuj wyników ręcznie.
Excel przelicza je po zmianie `C9`, `C6` lub `E14`.

### 3.3. Zadanie 3 – menedżer scenariuszy

**Nazwy zmienianych komórek**

| Komórka | Nazwa zdefiniowana |
| --- | --- |
| `C6` | `Cena_1_parowki` |
| `D14` | `Koszt_Smoczego_sadla` |
| `E14` | `Koszt_Mrowczego_miesa` |

**Zdefiniowane scenariusze**

| Scenariusz | Cena 1 parówki | Koszt smoczego sadła | Koszt mrówczego mięsa |
| --- | --- | --- | --- |
| Bieżący | 3,00 € | 15,00 € | 80,00 € |
| Optymistyczny | 3,50 € | 10,00 € | 70,00 € |
| Pesymistyczny | 2,50 € | 20,00 € | 90,00 € |

Scenariusze są zapisane w pliku.
Otwórz je przez `Dane` → `Menedżer scenariuszy`.

**Raport `Podsumowanie scenariuszy`**

| Pozycja | Wartości bieżące | Bieżący | Optymistyczny | Pesymistyczny |
| --- | --- | --- | --- | --- |
| Cena 1 parówki | 3,00 € | 3,00 € | 3,50 € | 2,50 € |
| Koszt smoczego sadła | 15,00 € | 15,00 € | 10,00 € | 20,00 € |
| Koszt mrówczego mięsa | 80,00 € | 80,00 € | 70,00 € | 90,00 € |
| Koszty roczne – smocze sadło | 23 590,81 € | 23 590,81 € | 15 727,21 € | 31 454,42 € |
| Koszty roczne – mrówcze mięso | 83 878,45 € | 83 878,45 € | 73 393,64 € | 94 363,26 € |
| Koszty roczne – makulatura | 4 718,16 € | 4 718,16 € | 4 718,16 € | 4 718,16 € |
| Koszty roczne – Magic Stufff | 655,30 € | 655,30 € | 655,30 € | 655,30 € |
| Koszty roczne – woda | 157,27 € | 157,27 € | 157,27 € | 157,27 € |
| Koszty roczne – razem | 113 000,00 € | 113 000,00 € | 94 651,59 € | 131 348,41 € |
| **ZYSK** | **10 000,00 €** | **10 000,00 €** | **57 348,41 €** | **−37 348,41 €** |

W raporcie szare komórki oznaczają zmieniane komórki wejściowe.
Nazwy zdefiniowane zastępują adresy komórek.
Raport jest zestawieniem statycznym, tak jak w Excelu.

---

## 4. Temat E6 – wykresy

### 4.1. Wykres kolumnowy

| Element | Wartość |
| --- | --- |
| Typ | kolumnowy, zgrupowany |
| Dane | `D35:H47` (Tabela 4, bez kolumny `Razem`) |
| Etykiety osi X | `B36:B47` (miesiące) |
| Tytuł | „Koszty składników w czasie (styczeń–grudzień)” |
| Tytuł osi X | „Miesiąc” |
| Tytuł osi Y | „Koszt [€]” |
| Legenda | na dole |
| Wyróżnienie | seria „Makulatura” ma kolor żółty |

Wykres pokazuje 5 serii składników dla 12 miesięcy.
Seria „Makulatura” jest żółta.
Pozostałe serie mają domyślne kolory motywu.

### 4.2. Wykres kołowy

| Element | Wartość |
| --- | --- |
| Typ | kołowy |
| Dane | `D48:H48` (roczne sumy składników) |
| Etykiety kategorii | `D35:H35` |
| Tytuł | „Struktura rocznych kosztów składników” |
| Legenda | po prawej stronie |
| Etykiety danych | nazwa składnika i procent (jedno miejsce po przecinku) |

Wykres pokazuje strukturę kosztów w stanie końcowym.
Największy udział ma mrówcze mięso: 74,2%.
Smocze sadło ma 20,9%.
Makulatura ma 4,2%.
Magic Stufff ma 0,6%, a woda 0,1%.

### 4.3. Wykres przebiegu w czasie (zadanie E6.2)

Instrukcja każe wybrać typ `Wstawianie` / `Wykresy przebiegu w czasie`.
Ten typ powstaje w Excelu tylko na bazie tabeli przestawnej.
Dlatego arkusz zawiera wykres liniowy z markerami po miesiącach.
Pokazuje ten sam przebieg w czasie, a dane bierze z kolumny `Razem`.

| Element | Wartość |
| --- | --- |
| Typ | liniowy z markerami |
| Dane | `I36:I47` (prawa biała kolumna `Razem` z Tabeli 4) |
| Nazwa serii | `Razem` (`I35`) |
| Etykiety osi X | `B36:B47` (miesiące) |
| Tytuł | „Przebieg kosztów całkowitych w czasie (styczeń–grudzień)” |
| Tytuł osi X | „Miesiąc” |
| Tytuł osi Y | „Koszt [€]” |
| Legenda | na dole |

Wykres pokazuje wzrost kosztów od stycznia do czerwca.
Potem koszty spadają w lipcu i sierpniu, a od września znowu rosną.

### 4.4. Układ i wydruk

* Wykres kolumnowy jest zakotwiczony w komórce `A85`.
* Wykres kołowy jest zakotwiczony w komórce `A106`.
* Wykres liniowy jest zakotwiczony w komórce `A132`.
* Arkusz ma orientację poziomą i skalowanie 90%.
* Ręczne podziały stron są po wierszach 32, 58, 83, 104 i 131.
* Każdy wykres zajmuje osobną stronę wydruku.

---

## 5. Kontrola poprawności

| Kontrola | Oczekiwana wartość | Wynik w pliku |
| --- | --- | --- |
| `C31` – liczba parówek | 58 000 | 58 000 |
| `I48` – roczny koszt składników | 113 000,00 € | 113 000,00 € |
| `I57` – ZYSK | 10 000,00 € | 10 000,00 € |
| `D51` – maks. koszt miesięczny | 15 586,21 € | 15 586,21 € |
| `D52` – miesiąc | Cze | Cze |
| `D53` – liczba parówek | 8 000 | 8 000 |
| `C62` – ZYSK dla 85 g | 16 733,85 € | 16 733,85 € |
| `C77` – ZYSK dla 100 g | −2 019,00 € | −2 019,00 € |
| `H67` – ZYSK dla 3,00 € i 80 €/kg | 10 000,00 € | 10 000,00 € |
| `F72` – ZYSK dla 3,50 € i 70 €/kg | 49 484,81 € | 49 484,81 € |
| `N62` – ZYSK dla 2,50 € i 110 €/kg | −50 454,42 € | −50 454,42 € |
| ZYSK – scenariusz Bieżący | 10 000,00 € | 10 000,00 € |
| ZYSK – scenariusz Optymistyczny | 57 348,41 € | 57 348,41 € |
| ZYSK – scenariusz Pesymistyczny | −37 348,41 € | −37 348,41 € |
| Wykres kolumnowy – zakres danych | `D36:H47` | `D36:H47` |
| Wykres kołowy – zakres danych | `D48:H48` | `D48:H48` |
| Wykres liniowy – zakres danych | `I36:I47` | `I36:I47` |
| Liczba wykresów w arkuszu | 3 | 3 |

---

## 6. Uwagi techniczne

* Plik ma zapisane formuły i ich wyniki. Dlatego liczby widać od razu.
* Formuły są w języku angielskim (format pliku). Polski Excel pokazuje je po polsku.
* Tabele danych i scenariusze działają w Excelu 2016 lub nowszym.
* W Excelu tabele danych przeliczają się po zmianie `C9`, `C6` lub `E14`.
* LibreOffice może nie odświeżać tabel danych. Naciśnij `F9` albo otwórz plik w Excelu.
* Scenariusze są zapisane w Menedżerze scenariuszy. Każdy ma trzy zmieniane komórki.
* Po zmianie danych wejściowych naciśnij `F9`, jeśli wykresy się nie odświeżą.
* Nie zmieniaj ręcznie komórek `C62:C82` i `F62:N72`. To wyniki tabel danych.
* Nazwy zdefiniowane: `Cena_1_parowki`, `Koszt_Smoczego_sadla`, `Koszt_Mrowczego_miesa`, `Koszty_roczne_*`, `ZYSK`.
* Arkusz `Podsumowanie scenariuszy` jest raportem statycznym. Nie przelicza się po zmianie scenariuszy.
