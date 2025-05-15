# Leak_test_VBA
Lake_test_VBA

## 📌 Opis

Makro `ImportWszystkieDane` importuje dane z czterech plików CSV:
 (`red1.csv`, `yellow1.csv`, `green1.csv`, `blue1.csv`) do aktywnego arkusza Excela. Dane są następnie przetwarzane (rozdzielane) i kopiowane w odpowiednie miejsce w arkuszu.

Główne funkcje:
- Automatyczny import danych z plików CSV znajdujących się w tym samym folderze co plik `.xlsm`.
- Parsowanie kolumn z wartościami liczbowymi w formacie tekstowym.


## 🚀 Uruchomienie

1. Otwórz plik `1.Wykres.xlsm` w Excelu.
2. Upewnij się, że włączone są makra.

## 🧱 Struktura kodu

- `ImportWszystkieDane`: Główna procedura zarządzająca importem.
- `ImportCSVred`: Specjalna wersja importera dla `red1.csv`, zapisuje dane do kolumn B i C.
- `ImportCSV_Generic`: Importer dla pozostałych plików, zapisuje do kolumn D–F.

## 🛠️ Przykład użycia

## ⚠️ Uwagi
Cały kod jest umieszczony w Module1,
Aby makro automatycznie się uruchamiało po kliknięciu w plik 1.Wykres.xlsm 
Poniższy kod umieszczamy w miejscu: „Ten_skoroszyt”
	Private Sub Workbook_Open()
    		On Error Resume Next
   		 Call ImportWszystkieDane
  		  If Err.Number <> 0 Then MsgBox "Błąd przy imporcie: " & Err.Description
   		 On Error GoTo 0
End Sub
