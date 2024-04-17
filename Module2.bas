Attribute VB_Name = "Module2"
'Arkusz Wykresy
Option Explicit

Sub OdswiezenieWykresow()
Attribute OdswiezenieWykresow.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveSheet.ChartObjects("WykresSrednieWynagrodzenieStanowiska").Activate
    ActiveChart.PivotLayout.PivotTable.PivotCache.Refresh
    Range("B2:B20").Sort , order1:=xlDescending
    
    ActiveSheet.ChartObjects("WykresLiczbaPracownikowWgStanowisk").Activate
    ActiveChart.PivotLayout.PivotTable.PivotCache.Refresh
    Range("B3:B20").Sort , order1:=xlDescending
    
    ActiveSheet.ChartObjects("WykresLiczbaPracownikowWgWieku").Activate
    ActiveChart.PivotLayout.PivotTable.PivotCache.Refresh
    Range("A39:A43").Sort
    
    ActiveSheet.ChartObjects("WykresRodzajeStanowisk").Activate
    ActiveChart.PivotLayout.PivotTable.PivotCache.Refresh
    
    ActiveSheet.ChartObjects("WykresSrednieWynagrodzenieWgRodzajuStanowiska").Activate
    ActiveChart.PivotLayout.PivotTable.PivotCache.Refresh

    Range("R1").Activate
End Sub

Sub WydrukWykresow()

    'ustawienia wydruku
    Application.PrintCommunication = False
    Range("A:B").EntireColumn.Hidden = True
    With ActiveSheet.PageSetup
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 0
    End With
    Application.PrintCommunication = True
    
    'wydruk
    ActiveSheet.PrintOut
 
    'usuniêcie linii podzia³u strony
    ActiveSheet.DisplayPageBreaks = False
    
    Range("A:B").EntireColumn.Hidden = False
    
End Sub

Sub ZapisWykresow()

    Dim sciezkaPliku As String
    Dim wykresDoWydruku As ChartObject
    Dim wysokoscWykresu As Integer
    Dim szerokoscWykresu As Integer
    
    'Set wykresDoWydruku = Sheets("Wykresy").ChartObjects("WykresSrednieWynagrodzenieStanowiska")

    For Each wykresDoWydruku In ActiveSheet.ChartObjects
        wysokoscWykresu = wykresDoWydruku.Height
        szerokoscWykresu = wykresDoWydruku.Width

        'zmiana rozmiaru
        wykresDoWydruku.Height = 400
        wykresDoWydruku.Width = 668
 
        'eksport
        sciezkaPliku = ThisWorkbook.Path & "\" & wykresDoWydruku.Name & ".jpg"
        wykresDoWydruku.Chart.Export sciezkaPliku

        'przywróæenie rozmiaru
        wykresDoWydruku.Height = wysokoscWykresu
        wykresDoWydruku.Width = szerokoscWykresu
    Next wykresDoWydruku
    
    MsgBox "Pliki zosta³y zapisane w folderze " & ThisWorkbook.Path, vbInformation
    
End Sub
