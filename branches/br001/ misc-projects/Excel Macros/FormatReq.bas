Attribute VB_Name = "FormatReq"
Sub FormatReqSpec()

' Remove any tables from within tables
' Open Word Req Spec and replace all line feeds "^l" with "£" and all "^p" with "£"
' Copy & Paste all requirement boxes from Req Spec into Excel

    Msg = "Open Req Spec in Word" + vbLf + "Remove any tables from within tables" + vbLf + "Replace all line feeds ^l with £ and all ^p with £" + vbLf + "Copy & Paste all requirement boxes from Req Spec into Excel"
    Response = MsgBox(Msg, vbOKOnly, "Instructions for use")
    
' Msgbox lines may need to be commented out if various fields are missing.

    RemoveBlankRows
    Transpose
    
' Copy Headings
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A2").Select
    Selection.Copy
    Range("E1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B2").Select
    Selection.Copy
    Range("F1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A3").Select
    Selection.Copy
    Range("G1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A4").Select
    Selection.Copy
    Range("H1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A5").Select
    Selection.Copy
    Range("I1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A6").Select
    Selection.Copy
    Range("J1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "REQ No."
    
' Delete original columns, remove highlighting & autofit
    Columns("A:D").Select
    Selection.Delete Shift:=xlToLeft
    RemoveBlankRows
    Cells.Select
    Selection.Interior.ColorIndex = xlNone
    Selection.Columns.AutoFit
    
' Remove borders & set column widths
    Columns("A:A").ColumnWidth = 10
    Columns("B:B").ColumnWidth = 15
    Columns("C:C").ColumnWidth = 50
    Columns("D:D").ColumnWidth = 50
    Columns("E:E").ColumnWidth = 50
    Columns("F:F").ColumnWidth = 50
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
' Add new headings & bolden
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Pass/Fail/NYD:"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Test Details:"
    Columns("G:G").ColumnWidth = 15
    Columns("H:H").ColumnWidth = 50
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    
' Reinseret line feeds
    ReplaceTags
    
' Freeze panes
    Range("B3").Select
    ActiveWindow.FreezePanes = True

End Sub
Sub Transpose()
'
' Transpose Macro
' Macro recorded 23/10/2009 by Kevin
   
    Dim i, j As Variant
    
    j = FindLastRow
    i = 0
    
    ' Split joined columns
    
    Columns("B:B").Select
    Selection.UnMerge
    
    Do
        ' Find REQ no. line
        Do
            i = i + 1
        Loop Until Left(Cells(i, 1), 3) = "REQ"
        
        ' Copy & move REQ number
        Cells(i, 1).Select
        Selection.Copy
        Cells(i, 5).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        ' Copy & move Type
        Cells(i, 3).Select
        Selection.Copy
        Cells(i, 6).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        
        ' Copy & move requirement
        If Left(Cells(i + 1, 1), 3) <> "Req" Then
            MsgBox "Row has no requirement: " & i + 1
            End
        Else
            Cells(i + 1, 2).Select
            Selection.Copy
            Cells(i, 7).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
        End If
            
        ' Copy & move rationale
        If Left(Cells(i + 2, 1), 3) <> "Rat" Then
            MsgBox "Row has no rationale: " & i + 2
            End
        Else
            Cells(i + 2, 2).Select
            Selection.Copy
            Cells(i, 8).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
        End If
                
        ' Copy & move fit criteria
        If Left(Cells(i + 3, 1), 3) <> "Fit" Then
            MsgBox "Row has no fit criteria: " & i + 3
            End
        Else
            Cells(i + 3, 2).Select
            Selection.Copy
            Cells(i, 9).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
        End If

        ' Copy & move source
        If Left(Cells(i + 4, 1), 3) <> "Sou" Then
            MsgBox "Row has no source: " & i + 4
        Else
            Cells(i + 4, 2).Select
            Selection.Copy
            Cells(i, 10).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
        End If

    Loop Until i + 5 > j

End Sub
Sub RemoveBlankRows()
'   Removes all non-requirement rows from copied req spec.
 
    Dim i As Variant
    
    i = FindLastRow
 
    Do
        ' Check for duplicate rows and delete
        If Cells(i, 1).Interior.ColorIndex = xlColorIndexNone Then
            Cells(i, 1).Select
            Selection.EntireRow.Delete
        End If
        
        i = i - 1
    Loop Until i = 0
        
End Sub
Function FindLastRow()
'   Find last filled row in sheet

Dim LastRow As Long

    If WorksheetFunction.CountA(Cells) > 0 Then

        'Search for any entry, by searching backwards by Rows.

        LastRow = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
    FindLastRow = LastRow

End Function
Function FindLastColumn()
'   Find last filled column in sheet

Dim LastColumn As Long

    If WorksheetFunction.CountA(Cells) > 0 Then

        'Search for any entry, by searching backwards by Columns.

        LastColumn = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    FindLastColumn = LastColumn

End Function
Sub ReplaceTags()
    ' Select text columns and reinsert line feed chars
    Columns("A:E").Select
    Selection.Replace What:="£", Replacement:=Chr(10), LookAt:=xlPart, SearchOrder:=xlByRows
End Sub
