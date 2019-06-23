Attribute VB_Name = "Module1"
Sub VBA_HW2b()
'
' VBA_HW2b Macro
'

'
    Application.CutCopyMode = False
    Cells(2, 11).FormulaR1C1 = "=SUMIF(R2C1:R705714C1,RC[-1],R2C7:R705714C7)"
    Range("K2").Select
    Selection.Copy
    Range("J2").Select
    Selection.End(xlDown).Select
    Range("K2836").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
End Sub



