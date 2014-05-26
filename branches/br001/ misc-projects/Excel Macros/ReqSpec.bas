Attribute VB_Name = "ReqSpec"
Sub UpdateTable()
'
' Copy current data into new free row in table
'
    Const StartRow = 69
    Const MaxRow = 93
    
    ' Copy current data
    Range("C68:AM68").Select
    Selection.Copy
    
    ' Find next free row
    Row = StartRow
    Do
        Row = Row + 1
    Loop Until Cells(Row, 2) = "" Or Row = MaxRow
    
    If Row = MaxRow Then
        MsgBox ("No spare rows!")
    Else
        Cells(Row, 3).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Cells(Row, 2).Select
        Cells(Row, 2) = Date
        Application.CutCopyMode = False
    End If

End Sub
Sub ShowHideSheet()
' Hide/unhide all tabs with a name containing the word 'Link'
' Operation of macro is a toggle.
'
    Dim sheet As Worksheet

    For Each sheet In ActiveWorkbook.Worksheets
        If InStr(sheet.Name, "Link") Then
            If sheet.Visible = True Then
                sheet.Visible = False
            Else
                sheet.Visible = True
            End If
        End If
    Next
    
End Sub
Sub MoveComments()
' Autosize & move comment to align with cell to which they refer.
'
  Dim cmt As Comment
  Dim sht As Worksheet
  
  For Each sht In ActiveWorkbook.Worksheets
  
  For Each cmt In sht.Comments
    With cmt
        .Shape.TextFrame.AutoSize = True
        .Shape.Top = .Parent.Top
        .Shape.Left = .Parent.Offset(0, 1).Left
    End With
  Next
  Next
End Sub
Sub FormatSelection()

Dim Cl As Range
Dim SearchText As String
Dim StartPos As Integer
Dim EndPos As Integer
Dim TestPos As Integer
Dim TotalLen As Integer

On Error Resume Next
Application.DisplayAlerts = False
SearchText = Application.InputBox _
(Prompt:="Enter string.", Title:="Which string to format?", Type:=2)
On Error GoTo 0
Application.DisplayAlerts = True
If SearchText = "" Then
    Exit Sub
Else
    For Each Cl In Selection
      TotalLen = Len(SearchText)
      StartPos = InStr(Cl, SearchText)
      TestPos = 0
      Do While StartPos > TestPos
        With Cl.Characters(StartPos, TotalLen).Font
          .FontStyle = "Bold"
          .ColorIndex = 3
        End With
        EndPos = StartPos + TotalLen
        TestPos = TestPos + EndPos
        StartPos = InStr(TestPos, Cl, SearchText, vbTextCompare)
      Loop
    Next Cl
End If

End Sub
Sub CheckIDnums()
'
' Check that the numberic part of the REQ ID is unique.
'
' Automatically selects all REQ ID cells and looks for duplicates.
'

Dim FirstCl As Range
Dim SecondCl As Range

Duplicates = False

Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp)).Select

For Each FirstCl In Selection
  ' Make a note of current number and reset match counter
  RefNum = Right(FirstCl, 4)
  Found = 0
  ' Compare with all other IDs
    For Each SecondCl In Selection
        ' Make a note of current number (last 4 chars)
        Num = Right(SecondCl, 4)
        If RefNum = Num Then Found = Found + 1
        If Found > 1 Then
            MsgBox ("Duplicate found " & RefNum)
            Duplicates = True
            ' Select duplicate
            FirstCl.Select
            Exit For
        End If
    Next SecondCl
    ' No need to look further
    If Duplicates Then Exit For
Next FirstCl
If Duplicates Then
    MsgBox ("ID Check-" & "Duplicates found")
Else
    MsgBox ("ID Check-" & "No duplicates")
    ' Select top cell to finish
    Cells(1, 1).Select

End If

End Sub

Sub CrossRefGen()
'
' Copy all linked reqs from active sheet to new sheet
' Separate reqs with multiple links onto separate rows to aid comparison
'
' Select a cell containing the ref links to copy

' Turn off screen updates to improve performance
Application.ScreenUpdating = False

' Get active worksheet name
OrigSheetName = ActiveSheet.Name
NewSheetName = OrigSheetName & "-Links"
CreateSheet (NewSheetName)
' Get selected column name for links
ActiveWorkbook.Sheets(OrigSheetName).Activate
LinkCol = ActiveCell.Column
LinkColName = Cells(2, ActiveCell.Column)

' Determine number of active rows.
MaxRows = Sheets(OrigSheetName).UsedRange.Rows.Count

' Copy linked rows
Call CopyRows(OrigSheetName, NewSheetName, LinkCol, MaxRows)
' Delete all the unwanted columns from the new link sheet
Call RemoveUnwantedCols(NewSheetName, 50)
' Determine new link column number on new sheet
LinkCol = FindCol(NewSheetName, LinkColName)
' Split rows with multiple links
Call SplitRows(NewSheetName, LinkCol)
' Tidy up the sheet
Call TidyUp(NewSheetName, LinkCol)

' Turn back on screen updates
Application.ScreenUpdating = True

MsgBox ("Links Copied")

End Sub
Sub CreateSheet(ByVal SheetName As String)
' Create a new sheet unless it exists
' If it exists delete its contents

Dim wsTest As Worksheet
 
Set wsTest = Nothing
On Error Resume Next
Set wsTest = ActiveWorkbook.Worksheets(SheetName)
On Error GoTo 0
 
If wsTest Is Nothing Then
    Worksheets.Add.Name = SheetName
Else
    Application.DisplayAlerts = False
    Sheets(SheetName).Delete
    Application.DisplayAlerts = True
    Worksheets.Add.Name = SheetName
End If

End Sub
Sub CopyRows(ByVal FromSheet As String, ByVal ToSheet As String, ByVal LinkCol As Integer, ByVal MaxRows As Long)
' Copy rows from one sheet to a new sheet where the link cell is not empty

i = 1
For CurrentRow = 2 To MaxRows
    ' Does it have a link?
    'ActiveWorkbook.Sheets(FromSheet).Activate
    Link = Sheets(FromSheet).Cells(CurrentRow, LinkCol)
    If Link <> "" Then
      ' Copy row to new sheet
      Sheets(FromSheet).Cells(CurrentRow, 1).EntireRow.Copy Sheets(ToSheet).Cells(i, 1)
      i = i + 1
    End If
Next CurrentRow

End Sub
Sub RemoveUnwantedCols(ByVal SheetName As String, MaxCols As Integer)
' Remove unwanted columns

ActiveWorkbook.Sheets(SheetName).Activate

For CurrentCol = MaxCols To 1 Step -1
    ' Find required cols
    Header = Cells(1, CurrentCol)
    If InStr(UCase(Header), "REQ") = 0 And InStr(UCase(Header), "LINK") = 0 Then
      ' Delete column
      Cells(1, CurrentCol).EntireColumn.Delete
    End If
Next CurrentCol

End Sub
Function FindCol(ByVal SheetName As String, ByVal LinkColName As String)
' Find row number containing links of interest

Set rng1 = Sheets(SheetName).UsedRange.Find(LinkColName, , xlValues, xlWhole)
If Not rng1 Is Nothing Then
    FindCol = rng1.Column
Else
    MsgBox "Not found", vbCritical
End If

End Function
Sub SplitRows(ByVal SheetName As String, ByVal LinkCol As Long)
' Split rows with multiple links

Dim rng1 As Range

' Determine number of active rows.
MaxRows = Sheets(SheetName).UsedRange.Rows.Count

CurrentRow = 2
Do
    ' Does it have multiple links?
    Links = Cells(CurrentRow, LinkCol)
    ' Split into separate links
    SepLinks = Split(Links, ", ")
    
    NumLinks = UBound(Split(Links, ","))
    MaxRows = MaxRows + NumLinks
    If NumLinks > 0 Then
        For LinkCount = 0 To NumLinks - 1
            ' Insert row
            Cells(CurrentRow + 1, 1).EntireRow.Insert
            ' Copy current row
            Cells(CurrentRow, 1).EntireRow.Copy Cells(CurrentRow + 1, 1)
            ' Write single link to current row & copied
            Cells(CurrentRow, LinkCol) = SepLinks(LinkCount)
            Cells(CurrentRow + 1, LinkCol) = SepLinks(LinkCount + 1)
            CurrentRow = CurrentRow + 1
        Next LinkCount
    End If
    CurrentRow = CurrentRow + 1
Loop Until CurrentRow > MaxRows

End Sub
Sub TidyUp(ByVal SheetName As String, ByVal LinkCol As Long)
' Move Link column & sort

ActiveWorkbook.Sheets(SheetName).Activate

Columns(LinkCol).Select
Selection.Cut
Columns("B:B").Select
Selection.Insert Shift:=xlToRight

' Clear formatting
ActiveSheet.Cells.Select
Selection.ClearFormats
Selection.Columns.AutoFit
' Reduce Requirement width
Columns("C").ColumnWidth = 100
ActiveSheet.Range("C:C").WrapText = True
' Set row height
Columns("B").RowHeight = 15

MaxRows = Sheets(SheetName).UsedRange.Rows.Count

Rows("1:1").Select
Selection.AutoFilter

Range("A1").Select
Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add _
    Key:=Range("B2:B" & MaxRows), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(SheetName).Sort.SortFields.Add _
    Key:=Range("A2:A" & MaxRows), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets(SheetName).Sort
    .SetRange Range("A1:E" & MaxRows)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

End Sub

