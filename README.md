Leak_test_VBA
🚀 Uruchomienie

    Upewnij się, że pliki:
    red1.csv, yellow1.csv, green1.csv, blue1.csv
    znajdują się w tym samym folderze co plik 1.Wykres.xlsm.

    Otwórz plik 1.Wykres.xlsm w programie Microsoft Excel.

    Włącz makra, jeśli pojawi się stosowny komunikat.

🛠️ Odblokowanie pliku na nowym komputerze

Jeśli plik został pobrany z internetu lub e-maila, Excel może zablokować makra. Aby je odblokować:

    Zamknij plik Excela.

    Przejdź do folderu, w którym znajduje się 1.Wykres.xlsm.

    Kliknij prawym przyciskiem myszy na plik i wybierz Właściwości.

    Zaznacz pole Odblokuj (jeśli jest dostępne).

    Kliknij Zastosuj, a następnie OK.

    Otwórz ponownie plik w Excelu.

⚙️ Uwagi techniczne

    Cały kod znajduje się w module Module1.

    Aby makro automatycznie uruchamiało się po otwarciu pliku, należy dodać poniższy kod do sekcji „Ten_skoroszyt” (ThisWorkbook):

    	Private Sub Workbook_Open()
   		On Error Resume Next
    		Call ImportWszystkieDane
    		If Err.Number <> 0 Then
       		 MsgBox "Błąd przy imporcie: " & Err.Description
    		End If
    		On Error GoTo 0
	End Sub
