Imports System
'This code for doing iteration at the excel file

Sub p()
  
'Variable declaration
Dim i, n As Integer

i = 0
n = Cells(8, 2)

'cycle for iteration
Do
i = i + 1

    Range("F4:T14").Select
    Selection.Copy
    Range("F18").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        
Cells(8, 3) = i

Loop Until i > n


End Sub
