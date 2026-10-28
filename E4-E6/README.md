# TI.1 – Arkusz kalkulacyjny: zadania E4, E5 i E6

Repozytorium zawiera plik roboczy oraz dokumentację do zadań laboratoryjnych z tematów **E4**, **E5** oraz **E6** (studium wykonalności spółki Hog-Dog sp. z o.o.).

Plik: `TI.1_Excel.E4+E5+E6_PlikRoboczy.xlsx`

---

## 1. Zawartość skoroszytu

| Arkusz | Opis zawartości |
| --- | --- |
| `Wykonalność` | Główny model finansowy: tabele 1–6 (E4), tabele wrażliwości 7–8 (E5) oraz menedżer scenariuszy. |
| `Podsumowanie scenariuszy` | Oficjalny raport Excela zestawiający scenariusze: Bieżący, Optymistyczny i Pesymistyczny. |

---

## 2. Temat E4 – Adresowanie mieszane i funkcje wyszukiwania

### 2.1. Zadanie 1 – Formuły i adresowanie mieszane

1. **Obsługa błędu dzielenia przez zero (Tab.1, C7)**:
   * Formuła: `=IF(C31 <> 0, I48/C31, "Nie można dzielic przez 0")`
   * Komórka C31 to łączna liczba parówek. Gdy C31 jest pusta lub równa 0, funkcja `JEŻELI` zabezpiecza arkusz przed błędem `#DZIEL/0!`.

2. **Waga składników (Tab.3, D19:H30)**:
   * Formuła w D19: `=$C19*$C$9/1000*D$13`
   * `$C19`: kolumna C zablokowana (liczba parówek), wiersz względny (zmienia się w dół dla kolejnych miesięcy).
   * `$C$9`: adres bezwzględny wagi pojedynczej parówki w gramach.
   * `D$13`: wiersz 13 zablokowany (udział procentowy składnika), kolumna względna (zmienia się w prawo dla kolejnych składników).
   * Dzielnik `/1000`: przeliczenie gramów na kilogramy.
   * Dzięki adresowaniu mieszanemu całą tabelę wypełnia się dwoma przeciągnięciami myszy.

3. **Koszt składników (Tab.4, D36:H47)**:
   * Formuła w D36: `=D19*D$14`
   * `D19`: waga danego składnika w danym miesiącu.
   * `D$14`: wiersz 14 zablokowany (koszt jednostkowy składnika w €/kg z Tabeli 2).

4. **Rachunek wyników (Tab.6)**:
   * Wpływy (I51): `=C31*C6`
   * Koszty zmienne (I52): `=I48`
   * Koszty stałe (I53): `=12*C8`
   * Koszty całkowite (I54): `=I48+I53`
   * Dochód (I55): `=I51-I54`
   * ZYSK (I57): `=I55-I56` (po odjęciu 15 000 € raty kredytu).
   * Wyjściowo przy wadze 100 g zysk wynosił **−2 019 €** (strata).

---

### 2.2. Zadanie 2 – Funkcje wyszukiwania

| Podpunkt | Komórka | Formuła w pliku | Wynik | Wyjaśnienie |
| --- | --- | --- | --- | --- |
| **2a** (Max koszt) | `D51` | `=MAX(I36:I47)` | **15 586,21 €** | Najwyższy miesięczny koszt składników (czerwiec). |
| **2b** (Miesiąc) | `D52` | `=XLOOKUP(D51; I36:I47; B36:B47)` | **Cze** | Nazwa miesiąca o najwyższym koszcie. |
| **2c** (Liczba szt.) | `D53` | `=VLOOKUP(D52; B19:H30; 2; 0)` | **8 000** | Liczba wyprodukowanych parówek w czerwcu. |

#### Odpowiedź na pytanie z instrukcji: Czy funkcja `WYSZUKAJ` jest tu odpowiednia?
* **Nie, funkcja `WYSZUKAJ` (`LOOKUP`) nie jest tu odpowiednia.**
* Wymaga ona posortowania przeszukiwanego wektora rosnąco. Kolumna kosztów miesięcznych nie jest posortowana, więc `WYSZUKAJ` zwraca błędne dane.
* Klasyczne rozwiązanie w Excelu: zestawienie `=INDEKS(B36:B47; PODAJ.POZYCJĘ(D51; I36:I47; 0))`.
* Nowoczesne rozwiązanie (zastosowane w pliku): funkcja `XLOOKUP` (`X.WYSZUKAJ`), która szuka w lewo i nie wymaga sortowania.

---

## 3. Temat E5 – Analiza warunkowa („co-jeśli”)

### 3.1. Zadanie 1 – Szukaj wyniku (Goal Seek)
* **Cel**: Zmiana wagi parówki tak, aby firma osiągnęła zysk równy **10 000 €**.
* **Ścieżka w menu**: `Dane` → `Analiza warunkowa` → `Szukaj wyniku`.
* **Ustawienia**:
  * Ustaw komórkę: `I57` (ZYSK)
  * Na wartość: `10000`
  * Zmieniając komórkę: `C9` (Waga parówki)
* **Wynik**: Excel wyznaczył nową wagę parówki: **`90,386... g`** (ok. 90,39 g).

---

### 3.2. Zadanie 2 – Tabele danych (analiza wrażliwości)

#### Tabela 7: Wpływ wagi parówki na zysk (tabela 1-wymiarowa kolumnowa)
* Zakres tabeli: `B61:C82`.
* Wartości wagi: od 85 g do 105 g w kolumnie B.
* Komórka wynikowa (nagłówek kolumny C61): `=I57`.
* `Dane` → `Analiza warunkowa` → `Tabela danych`:
  * Komórka wejściowa kolumny: `$C$9` (Waga parówki).
* **Wnioski**:
  * Przy wadze 85 g zysk rośnie do **16 733,85 €**.
  * Próg rentowności (zysk = 0) przypada przy wadze ok. **98,4 g**.
  * Przy wyjściowej wadze 100 g zysk wynosi dokładnie **−2 019 €**.

#### Tabela 8: Wpływ ceny i kosztu mięsa na zysk (tabela 2-wymiarowa macierzowa)
* Zakres tabeli: `E61:N72`.
* Wiersze (E62:E72): Cena 1 parówki od 2,50 € do 3,50 € (co 0,10 €).
* Kolumny (F61:N61): Jednostkowy koszt mrówczego mięsa od 70 € do 110 € (co 5 €).
* Komórka narożna (E61): `=I57`.
* `Dane` → `Analiza warunkowa` → `Tabela danych`:
  * Komórka wejściowa wiersza: `$E$14` (Jednostkowy koszt mrówczego mięsa).
  * Komórka wejściowa kolumny: `$C$6` (Cena 1 parówki).
* **Wnioski**:
  * Dla ceny bazowej 3,00 € i kosztu mięsa 80 € zysk wynosi dokładnie **10 000,00 €**.
  * Najlepszy wariant w macierzy (cena 3,50 €, koszt mięsa 70 €) daje zysk **49 484,81 €**.
  * Najgorszy wariant (cena 2,50 €, koszt mięsa 110 €) przynosi stratę **−50 454,42 €**.

---

### 3.3. Zadanie 3 – Menedżer scenariuszy

#### Zdefiniowane nazwy komórek
Aby raport był czytelny zamiast surowych adresów komórek nadano nazwy:
* `$C$6` → `Cena_1_parowki`
* `$D$14` → `Koszt_Smoczego_sadla`
* `$E$14` → `Koszt_Mrowczego_miesa`
* `$D$48`...`$I$48` → `Koszty_roczne_...`
* `$I$57` → `ZYSK`

#### Zestawienie scenariuszy i raport podsumowania

| Parametr / Wynik | Wartości bieżące | Bieżący | Optymistyczny | Pesymistyczny |
| --- | --- | --- | --- | --- |
| **Cena 1 parówki** | 3,00 € | 3,00 € | 3,50 € | 2,50 € |
| **Koszt Smoczego sadła** | 15,00 € | 15,00 € | 10,00 € | 20,00 € |
| **Koszt Mrówczego mięsa** | 80,00 € | 80,00 € | 70,00 € | 90,00 € |
| **Roczny koszt składników** | 113 000,00 € | 113 000,00 € | 94 651,59 € | 131 348,42 € |
| **ZYSK KOŃCOWY** | **10 000,00 €** | **10 000,00 €** | **57 348,41 €** | **−37 348,41 €** |

Raport został wygenerowany w dedykowanym arkuszu `Podsumowanie scenariuszy`.
