Sub MF()

Dim i As Integer
Dim t, t_var, dt As Double

Dim j As Integer
Dim tt, tt_var, dtt As Double


i = 0
t = Cells(5, 2)
dt = Cells(4, 2)

j = 0
tt = Cells(5, 2)
dtt = Cells(4, 2)


Do
i = i + 1
t_var = i * dt

j = j + 1
tt_var = j * dtt

   Range("H7:V7").Select
    Selection.Copy
        Range("H5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Application.CutCopyMode = False
         
 Range("W7:AL7").Select
    Selection.Copy
        Range("W5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Application.CutCopyMode = False
        
Cells(6, 3) = tt_var
Cells(6, 2) = t_var
Loop Until t_var > t

End Sub
