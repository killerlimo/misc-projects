Attribute VB_Name = "RMS"
Function RMS(values)
'
' This computes the root mean square of entered values
'
    RMS = Sqr(WorksheetFunction.SumSq(values) / WorksheetFunction.Count(values))
End Function

Sub RMSmacro()
'
' RMS Macro
' Macro recorded 13/01/2010 by Administrator
' Calculates RMS values from columns B & C.
' Time steps must be even for true RMS.

    Dim MaxRow, MaxCol As Variant
    
    MaxRow = FindLastRow
    MaxCol = FindLastColumn
    
    Range("2:2").Select
    Selection.Insert
    
    ' Add row title
    Range("A2").Select
    ActiveCell.Formula = "RMS"
    
    ' Insert RMS values directly into RMS cells
    For c = 2 To MaxCol
        Cells(2, c) = RMS(Range(Cells(3, c), Cells(MaxRow, c)))
    Next c
    
    Range("b2", Cells(2, MaxCol)).Select
    Selection.Copy

End Sub

