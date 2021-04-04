Imports System
'This code for doing iteration at the excel file
Sub iterat1()

    'Variable declaration
    Dim i As Integer
    Dim t, t_var, dt As Double

    i = 0
    t = Cells(5, 2) 'Calculation time
    dt = Cells(4, 2) 'Time step

    'cycle for iteration
    Do
        i = i + 1
        t_var = i * dt


        Range("G7:AI7").Select
        Selection.Copy
        Range("G5").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Range("S21").Select
        Application.CutCopyMode = False

        Cells(12, 3) = t_var
    Loop Until t_var > t

End Sub
