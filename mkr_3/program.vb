Imports System
'This code for doing iteration at the excel file

Sub check_2()

'Variable declaration
Dim i As Integer
Dim t, t_var, dt As Double

i = 0
t = Cells(5, 2)
dt = Cells(6, 2)

'cycle for iteration
Do
i = i + 1
t_var = i * dt

    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 41
    Range("H69:AJ99").Select
    Selection.Copy
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 1
    Range("H3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        
Cells(6, 3) = t_var
Loop Until t_var > t

End Sub
