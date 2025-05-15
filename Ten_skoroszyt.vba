Private Sub Workbook_Open()
    On Error Resume Next
    Call ImportWszystkieDane
    If Err.Number <> 0 Then MsgBox "Błąd przy imporcie: " & Err.Description
    On Error GoTo 0
End Sub
