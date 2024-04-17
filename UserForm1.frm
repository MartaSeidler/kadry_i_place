VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Dodaj pracownika"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5700
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UserForm
Option Explicit

Private Sub CancelButton_Click()
    Unload UserForm1
End Sub


Private Sub OKButton_Click()
    
    'okreœlanie pustego wiersza
    Dim NextRow As Integer
    Sheets("Pracownicy").Activate
    NextRow = Application.WorksheetFunction.CountA(Range("B:B")) + 1
    
    If Not NiepoprawneDane Then
        'imiê i nazwisko
        Cells(NextRow, 3) = TextBox1.Text
        Sheets("DaneDodatkowe").Cells(NextRow, 2) = TextBox1.Text
        Sheets("ListaP³ac").Cells(NextRow + 1, 2) = TextBox1.Text
        
        'PESEL
        Cells(NextRow, 5) = TextBox3.Text
        
        'Data zatrudnienia
        Cells(NextRow, 7) = ComboBox1.Value & "." & ComboBox2.Value & "." & ComboBox3.Value
        
        'stanowisko
        Cells(NextRow, 4) = ListBox1.Value
        
        'wynagrodzenie
        Cells(NextRow, 9) = TextBox5.Text
        
        'ubezpieczenie
        Cells(NextRow, 12) = TextBox6.Text
        
        'czy zwiêkszone koszty
        If CheckBox1.Value = True Then
            Cells(NextRow, 10) = "TAK"
        Else
            Cells(NextRow, 10) = "NIE"
        End If
        
        'czy ulga podatkowa
        If CheckBox2.Value = True Then
            Cells(NextRow, 11) = "TAK"
        Else
            Cells(NextRow, 11) = "NIE"
        End If
        
    Else
        Exit Sub
    End If
    
    'informacja
    Dim Msg As String
    Msg = "Dodano pracownika"
    MsgBox Msg
    
    'sortowanie tabeli wg nazwiska
    Call SortowanieTabeli
    Call SortowanieDanychDodatkowych
    Call SortowanieListyPlac
    
    'czyszczenie okna dialogowego
    Call CzyszczenieOknaDialogowego
    
    'zamkniêcie formularza
    Unload UserForm1
End Sub

Private Function NiepoprawneDane() As Boolean

    'sprawdzanie czy s¹ wype³nione obowi¹zkowe pola
    If TextBox1.Text = "" Then
        MsgBox "WprowadŸ nazwisko i imiê!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If TextBox3.Text = "" Then
        MsgBox "WprowadŸ PESEL!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If ComboBox1.Value = "" Or ComboBox2.Value = "" Or ComboBox3.Value = "" Then
        MsgBox "WprowadŸ datê zatrudnienia!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If ListBox1.Value = "" Then
        MsgBox "Wybierz stanowisko!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If TextBox5.Text = "" Then
        MsgBox "WprowadŸ wynagrodzenie!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    
    'sprawdzanie poprawnoœci danych
    If Not IsNumeric(TextBox3.Value) Or Len(TextBox3.Value) <> 11 Then
        MsgBox "Niepoprawny PESEL!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If Not IsNumeric(ComboBox1.Value) Then
        MsgBox "Niepoprawna data zatrudnienia!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If ComboBox1.Value > 31 Then
        MsgBox "Niepoprawna data zatrudnienia!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If Not IsNumeric(ComboBox2.Value) Then
        MsgBox "Niepoprawna data zatrudnienia!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If ComboBox1.Value > 12 Then
        MsgBox "Niepoprawna data zatrudnienia!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If Not IsNumeric(ComboBox3.Value) Then
        MsgBox "Niepoprawna data zatrudnienia!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If Not IsNumeric(TextBox5.Value) Then
        MsgBox "Niepoprawna wartoœæ wynagrodzenia!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    If Not IsNumeric(TextBox6.Value) Then
        MsgBox "Niepoprawna wartoœæ ubezpieczenia!", vbExclamation
        NiepoprawneDane = True
        Exit Function
    End If
    
    NiepoprawneDane = False
End Function

Private Sub SortowanieTabeli()
    ActiveWorkbook.Worksheets("Pracownicy").ListObjects("Tabela1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Pracownicy").ListObjects("Tabela1").Sort.SortFields. _
        Add2 Key:=Range("Tabela1[[#All],[Pracownik]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Pracownicy").ListObjects("Tabela1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub SortowanieDanychDodatkowych()
    ActiveWorkbook.Worksheets("DaneDodatkowe").ListObjects("Tabela4").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("DaneDodatkowe").ListObjects("Tabela4").Sort.SortFields. _
        Add2 Key:=Range("Tabela4[[#All],[Pracownik]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DaneDodatkowe").ListObjects("Tabela4").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub SortowanieListyPlac()
    ActiveWorkbook.Worksheets("ListaP³ac").ListObjects("Tabela5").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("ListaP³ac").ListObjects("Tabela5").Sort.SortFields. _
        Add2 Key:=Range("Tabela5[[#All],[Pracownik]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ListaP³ac").ListObjects("Tabela5").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub CzyszczenieOknaDialogowego()
    TextBox1.Text = ""
    TextBox3.Text = ""
    TextBox5.Text = ""
    TextBox6.Text = ""
    CheckBox1.Value = False
    CheckBox2.Value = False
    ComboBox1.Value = ""
    ComboBox2.Value = ""
    ComboBox3.Value = ""
End Sub

Private Sub UserForm_Initialize()
    '³adowanie listy stanowisk
    Dim CountRows As Integer
    Sheets("Stanowiska").Activate
    CountRows = Application.WorksheetFunction.CountA(Range("A:A"))
    ListBox1.List = Worksheets("Stanowiska").Range(Cells(2, 1), Cells(CountRows, 1)).Value
    ListBox1.ListIndex = -1
    Sheets("Pracownicy").Activate
    
    '³adowanie dni, miesiêcy i lat do daty zatrudnienia
    ComboBox1.List = Worksheets("Stanowiska").Range("E2:E32").Value
    ComboBox2.List = Worksheets("Stanowiska").Range("F2:F13").Value
    ComboBox3.List = Worksheets("Stanowiska").Range("G2:G36").Value
End Sub


