Attribute VB_Name = "Module1"
Sub findColumnByNumber()



    c = Application.InputBox("Please enter filter column", Type:=2)
    If IsNumeric(c) = True Then
         MsgBox " your column is " & Chr(c + 64)
         Columns(Int(c)).Select
    ElseIf Asc(UCase(c)) >= 65 And Asc(UCase(c)) <= 90 Then
         MsgBox " your column is " & UCase(c)
         Columns(UCase(c)).Select
    Else
        MsgBox "wrong input"
    End If


End Sub
