Attribute VB_Name = "ReqSpec"
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
Sub CrossRefsMainDB()
'
' Copy all linked reqs from main database to new sheet
' Separate reqs with multiple links onto separate rows

Dim ReqId As Range
Dim SecondCl As Range

'GoTo j2
' Turn off screen updates to improve performance
Application.ScreenUpdating = False
ActiveWorkbook.Sheets("Requirements Database").Activate

' Determin number of active rows.
MaxRows = Sheets("Requirements Database").UsedRange.Rows.Count
i = 2

For CurrentRow = 2 To MaxRows
    ' Does it have a link?
    ActiveWorkbook.Sheets("Requirements Database").Activate
    CustLink = Cells(CurrentRow, 27)
    ET400Link = Cells(CurrentRow, 28)
    If CustLink <> "" Or ET400Link <> "" Then
      ' Copy row to new sheet
      Cells(CurrentRow, 1).EntireRow.Copy Sheets("Cross Ref-DB").Cells(i, 1)
      i = i + 1
    End If
Next CurrentRow

j1:
' Remove unwanted columns
ActiveWorkbook.Sheets("Cross Ref-DB").Activate
' Determin number of active cols.
'MaxCols = ActiveSheet.UsedRange.Columns.Count
MaxCols = 50

For CurrentCol = MaxCols To 1 Step -1
    ' Does it have a link?
    Header = Cells(2, CurrentCol)
    Debug.Print Header
    If Header <> "REQ No." And Header <> "Requirement:" And Header <> "Link to Customer Req:" And Header <> "Link to ET400 Req:" Then
      ' Delete column
      Cells(1, CurrentCol).EntireColumn.Delete
    End If
Next CurrentCol

' Turn back on screen updates
Application.ScreenUpdating = True

j2:
' Split rows with multiple links
' Determin number of active rows.
MaxRows = Sheets("Cross Ref-DB").UsedRange.Rows.Count

CurrentRow = 3
Do
    ' Does it have multiple links?
    CustLink = Cells(CurrentRow, 3)
    ' Split into separate links
    SepLinks = Split(CustLink, ", ")
    'ET400Link = Cells(CurrentRow, 4)
    NumLinks = UBound(Split(CustLink, ","))
    MaxRows = MaxRows + NumLinks
    If NumLinks > 0 Then
        For Links = 0 To NumLinks - 1
            ' Insert row
            Cells(CurrentRow + 1, 1).EntireRow.Insert
            ' Copy current row
            Cells(CurrentRow, 1).EntireRow.Copy Cells(CurrentRow + 1, 1)
            ' Write single link to current row & copied
            Cells(CurrentRow, 3) = SepLinks(Links)
            Cells(CurrentRow + 1, 3) = SepLinks(Links + 1)
            CurrentRow = CurrentRow + 1
        Next Links
    End If
    CurrentRow = CurrentRow + 1
Loop Until CurrentRow > MaxRows

' Move Link column & sort
Columns("C:C").Select
Selection.Cut
Columns("B:B").Select
Selection.Insert Shift:=xlToRight
Rows("2:2").Select
Selection.AutoFilter
ActiveWorkbook.Worksheets("Cross Ref-DB").AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Cross Ref-DB").AutoFilter.Sort.SortFields.Add Key _
    :=Range("B2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets("Cross Ref-DB").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' Clear formatting
ActiveSheet.Cells.Select
Selection.ClearFormats
Selection.Columns.AutoFit
' Reduce Requirement width
Columns("C").ColumnWidth = 100
ActiveSheet.Range("C:C").WrapText = True
' Set row height
Columns("B").RowHeight = 15

' Turn back on screen updates
Application.ScreenUpdating = True

MsgBox ("Main Database Links Copied")

End Sub
Sub CrossRefsCustomerSpec()
'
' Copy all linked reqs from customer requirements to new sheet
'

Dim ReqId As Range
Dim SecondCl As Range

'GoTo j2

' Turn off screen updates to improve performance
Application.ScreenUpdating = False
ActiveWorkbook.Sheets("Customer Requirements").Activate
    
' Determin number of active rows.
MaxRows = Sheets("Customer Requirements").UsedRange.Rows.Count
i = 2
' Filter all reqs containing customer links
'Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp)).AutoFilter Field:=27, Criteria1:="<>"

For CurrentRow = 2 To MaxRows
    ' Does it have a link?
    ActiveWorkbook.Sheets("Customer Requirements").Activate
    ET410Link = Cells(CurrentRow, 4)
    If ET410Link <> "" Then
      ' Copy row to new sheet
      Cells(CurrentRow, 1).EntireRow.Copy Sheets("Cross Ref-Cust").Cells(i, 1)
      i = i + 1
    End If
Next CurrentRow

j1:
' Remove unwanted columns
ActiveWorkbook.Sheets("Cross Ref-Cust").Activate
' Determin number of active cols.
'MaxCols = ActiveSheet.UsedRange.Columns.Count
MaxCols = 50

For CurrentCol = MaxCols To 1 Step -1
    ' Does it have a link?
    Header = Cells(2, CurrentCol)
    If Header <> "" And Header <> "REQ No." And Header <> "Requirement:" And Header <> "Link to ET410 Req:" Then
      ' Delete column
      Cells(1, CurrentCol).EntireColumn.Delete
    End If
Next CurrentCol

j2:
' Split rows with multiple links
' Determin number of active rows.
MaxRows = Sheets("Cross Ref-Cust").UsedRange.Rows.Count

CurrentRow = 3
Do
    ' Does it have multiple links?
    CustLink = Cells(CurrentRow, 3)
    ' Split into separate links
    SepLinks = Split(CustLink, ", ")
    'ET400Link = Cells(CurrentRow, 4)
    NumLinks = UBound(Split(CustLink, ","))
    MaxRows = MaxRows + NumLinks
    If NumLinks > 0 Then
        For Links = 0 To NumLinks - 1
            ' Insert row
            Cells(CurrentRow + 1, 1).EntireRow.Insert
            ' Copy current row
            Cells(CurrentRow, 1).EntireRow.Copy Cells(CurrentRow + 1, 1)
            ' Write single link to current row & copied
            Cells(CurrentRow, 3) = SepLinks(Links)
            Cells(CurrentRow + 1, 3) = SepLinks(Links + 1)
            CurrentRow = CurrentRow + 1
        Next Links
    End If
    CurrentRow = CurrentRow + 1
Loop Until CurrentRow > MaxRows

' Move Link column & sort
Columns("C:C").Select
Selection.Cut
Columns("B:B").Select
Selection.Insert Shift:=xlToRight
Rows("2:2").Select
Selection.AutoFilter
ActiveWorkbook.Worksheets("Cross Ref-Cust").AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Cross Ref-Cust").AutoFilter.Sort.SortFields.Add Key _
    :=Range("B2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets("Cross Ref-Cust").AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' Clear formatting
ActiveSheet.Cells.Select
Selection.ClearFormats
Selection.Columns.AutoFit
' Reduce Requirement width
Columns("C").ColumnWidth = 100
ActiveSheet.Range("C:C").WrapText = True
' Set row height
Columns("B").RowHeight = 15

' Turn back on screen updates
Application.ScreenUpdating = True

MsgBox ("Customer Spec Links Copied")

End Sub
Sub CrossRefsGeneral()
'
' Copy all linked reqs from the current sheet to a new sheet
'

Dim ReqId As Range
Dim SecondCl As Range

' Turn off screen updates to improve performance
Application.ScreenUpdating = False
'ActiveWorkbook.Sheets("Customer Requirements").Activate
    
'GoTo j1
' Determin number of active rows.
MaxRows = ActiveSheet.UsedRange.Rows.Count
i = 2
' Filter all reqs containing customer links
'Range(Cells(2, 1), Cells(Rows.Count, 1).End(xlUp)).AutoFilter Field:=27, Criteria1:="<>"

For CurrentRow = 2 To MaxRows
    ' Does it have a link?
    ActiveWorkbook.Sheets("Customer Requirements").Activate
    ET410Link = Cells(CurrentRow, 4)
    If ET410Link <> "" Then
      ' Copy row to new sheet
      Cells(CurrentRow, 1).EntireRow.Copy Sheets("Cross Ref-Cust").Cells(i, 1)
      i = i + 1
    End If
Next CurrentRow

j1:
' Remove unwanted columns
ActiveWorkbook.Sheets("Cross Ref-Cust").Activate
' Determin number of active cols.
'MaxCols = ActiveSheet.UsedRange.Columns.Count
MaxCols = 50

For CurrentCol = MaxCols To 1 Step -1
    ' Does it have a link?
    Header = Cells(2, CurrentCol)
    If Header <> "" And Header <> "REQ No." And Header <> "Requirement:" And Header <> "Link to ET410 Req:" Then
      ' Delete column
      Cells(1, CurrentCol).EntireColumn.Delete
    End If
Next CurrentCol

' Clear formatting
ActiveSheet.Cells.Select
Selection.ClearFormats
Selection.Columns.AutoFit
' Reduce Requirement width
Columns("B").ColumnWidth = 100
ActiveSheet.Range("B:B").WrapText = True
' Set row height
Columns("B").RowHeight = 15

' Turn back on screen updates
Application.ScreenUpdating = True

' Add filter and freeze pane
Rows("2:2").Select
Selection.AutoFilter
Range("B3").Select
ActiveWindow.FreezePanes = True

MsgBox ("Customer Spec Links Copied")

End Sub







