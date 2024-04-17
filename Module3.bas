Attribute VB_Name = "Module3"
'Arkusz Wynagrodzenie
Option Explicit

Sub WydrukWynagrodzenia()
    
    'ustawienia wydruku
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True
    
    'wydruk
    ActiveSheet.PrintOut
 
    'usuniêcie linii podzia³u strony
    ActiveSheet.DisplayPageBreaks = False
    
End Sub
