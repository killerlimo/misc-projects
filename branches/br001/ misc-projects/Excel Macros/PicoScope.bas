Attribute VB_Name = "PicoScope"
Sub PicoScopeToCsv()
Attribute PicoScopeToCsv.VB_Description = "Converts PicoScope exported data into a csv file suitable for Octave."
Attribute PicoScopeToCsv.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PicoScopeToCsv Macro
' Converts PicoScope exported data into a csv file suitable for Octave.
'

'
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=HEX2DEC(RC[-3])"
    Range("F3").Select
    Selection.Copy
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveSheet.Paste
    Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C:R[65534]C)"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-R2C6"
    Range("G3").Select
    Selection.Copy
    Range("G4").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveSheet.Paste
    Range("G3").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A1").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
