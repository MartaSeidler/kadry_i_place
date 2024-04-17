Attribute VB_Name = "Module1"
'Arkusz Pracownicy
Option Explicit

Sub ShowDodajPracownika()
    UserForm1.Show
End Sub

Sub UsunPracownika()
    Dim Config As Integer
    Dim Ans As Integer
    Dim Msg As String
    Dim RowNumber As Integer
    
    RowNumber = Selection.Row
    
    Config = vbYesNo + vbQuestion + vbDefaultButton2
    Msg = "Usun¹æ pracownika "
    Msg = Msg & Cells(RowNumber, 3).Value
    Msg = Msg & "?"
    
    Ans = MsgBox(Msg, Config)
    If Ans = vbYes Then
        Rows(RowNumber).EntireRow.Delete
        Worksheets("DaneDodatkowe").Rows(RowNumber).EntireRow.Delete
        Worksheets("ListaP³ac").Rows(RowNumber + 1).EntireRow.Delete
     End If
End Sub

