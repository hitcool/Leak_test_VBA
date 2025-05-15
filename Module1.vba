Sub ImportWszystkieDane()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Ustawia referencję do pierwszego arkusza w skoroszycie (można zmienić na konkretną nazwę)

    ' Import danych z plików CSV do arkusza
    ImportCSVred ThisWorkbook.Path & "\red1.csv", ws
    ImportCSV_Generic ThisWorkbook.Path & "\yellow1.csv", ws, "D"
    ImportCSV_Generic ThisWorkbook.Path & "\green1.csv", ws, "E"
    ImportCSV_Generic ThisWorkbook.Path & "\blue1.csv", ws, "F"

    ' Rozdzielenie wartości w kolumnach C–F na podstawie stałej szerokości pól (TextToColumns)
    ws.Columns("C").TextToColumns Destination:=ws.Range("C1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 1), TrailingMinusNumbers:=True
    ws.Columns("D").TextToColumns Destination:=ws.Range("D1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 1), TrailingMinusNumbers:=True
    ws.Columns("E").TextToColumns Destination:=ws.Range("E1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 1), TrailingMinusNumbers:=True
    ws.Columns("F").TextToColumns Destination:=ws.Range("F1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(0, 1), TrailingMinusNumbers:=True

    ' Kopiuje zakres J1:N3 do obszaru zaczynającego się od komórki B1
    ws.Range("J1:N3").Copy Destination:=ws.Range("B1")

    ' Wyświetla komunikat po zakończeniu importu
    MsgBox "Import zakończony!"
End Sub


Sub ImportCSVred(csvPath As String, ws As Worksheet)
    Dim csvLine As String
    Dim fileNum As Integer
    Dim values() As String
    Dim rowNum As Long
    Dim polaczonaLiczba As String

    ' Sprawdza czy plik istnieje, jeśli nie – pokazuje komunikat
    If Dir(csvPath) = "" Then
        MsgBox "Brak pliku: " & csvPath
        Exit Sub
    End If

    fileNum = FreeFile ' Pobiera pierwszy dostępny numer pliku
    Open csvPath For Input As #fileNum ' Otwiera plik CSV do odczytu

    rowNum = 1 ' Zaczyna od pierwszego wiersza w arkuszu
    Do While Not EOF(fileNum) ' Pętla dopóki nie osiągnięto końca pliku
        Line Input #fileNum, csvLine ' Wczytuje pojedynczą linię z pliku
        values = Split(csvLine, ",") ' Dzieli linię na wartości rozdzielone przecinkiem

        If UBound(values) >= 5 Then ' Sprawdza czy linia zawiera co najmniej 6 elementów
            ws.Cells(rowNum, "B").Value = Trim(values(1)) ' Wstawia drugą wartość (indeks 1) do kolumny B
            polaczonaLiczba = Trim(values(4)) & "," & Trim(values(5)) ' Łączy wartości z kolumn 5 i 6 jako tekst liczby dziesiętnej
            ws.Cells(rowNum, "C").Value = polaczonaLiczba ' Wstawia połączoną wartość do kolumny C
        End If

        rowNum = rowNum + 1 ' Przechodzi do następnego wiersza
    Loop
    Close #fileNum ' Zamyka plik
End Sub


Sub ImportCSV_Generic(csvPath As String, ws As Worksheet, targetColumn As String)

' Deklaracja zmiennych

    Dim csvLine As String           ' przechowuje pojedynczy wiersz tekstu wczytany z pliku CSV
    Dim fileNum As Integer          ' unikalny numer deskryptora pliku, który otrzymuje otwarty plik,
    Dim values() As String          ' tablica ciągów znaków utworzona z wiersza CSV rozdzielonego przecinkami,
    Dim rowNum As Long              ' numer wiersza w arkuszu, gdzie będą wpisywane dane,
    Dim polaczonaLiczba As String   ' łączona wartość z kolumn 5 i 6 pliku CSV (czyli indeksy 4 i 5).

    ' Sprawdza czy plik istnieje
    If Dir(csvPath) = "" Then       '  zwraca nazwę pliku, jeśli istnieje
        MsgBox "Brak pliku: " & csvPath
        Exit Sub
    End If

    fileNum = FreeFile ' Pobiera pierwszy dostępny numer pliku
    Open csvPath For Input As #fileNum ' Otwiera plik CSV do odczytu

    rowNum = 1 ' Rozpoczyna od pierwszego wiersza
    Do While Not EOF(fileNum) ' Czyta linie aż do końca pliku
        Line Input #fileNum, csvLine ' Wczytuje jedną linię tekstu
        values = Split(csvLine, ",") ' Dzieli linię na części oddzielone przecinkiem

        If UBound(values) >= 5 Then ' Jeśli jest wystarczająco danych
            polaczonaLiczba = Trim(values(4)) & "," & Trim(values(5)) ' Łączy wartości z kolumn 5 i 6
            ws.Cells(rowNum, targetColumn).Value = polaczonaLiczba ' Wstawia je do odpowiedniej kolumny (D, E lub F)
        End If

        rowNum = rowNum + 1 ' Następny wiersz
    Loop
    Close #fileNum ' Zamyka plik
End Sub



