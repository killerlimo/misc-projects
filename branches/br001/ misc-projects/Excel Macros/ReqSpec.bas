Attribute VB_Name = "ReqSpec"
Sub ClearFilter()
' Clear the autofilter on the current worksheet.
    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
End Sub
Sub DeleteUndeleteRow()
' Mark req as Deleted/Linked.
' Toggle font to/from strikeout on current row.
' Cursor must be in on the row to be formatted.

    ' Get current req status
    Status = Cells(ActiveCell.Row, 3)
    
    If Status = "Deleted" Then
        ' Set status to Linked
        Cells(ActiveCell.Row, 3) = "Linked"
        
        ' Clear Strikeout on row
        StartingCell = ActiveCell.Address
        Rows(ActiveCell.Row).Select
        Selection.Font.Strikethrough = False
    Else
        ' Set status to Deleted
        Cells(ActiveCell.Row, 3) = "Deleted"
        
        ' Strikeout row
        StartingCell = ActiveCell.Address
        Rows(ActiveCell.Row).Select
        Selection.Font.Strikethrough = True
    End If
    
    ' Return cursor to original cell
    Range(StartingCell).Select
    
End Sub
Sub UpdateTable()
'
' Update Progress table
' Copy current data into new free row in table
'
    Const StartRow = 69
    Const MaxRow = 93
    
    ' Copy current data
    Range("CurrentStat").Select
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
Sub FindAllRefs()
' Take ref in current highlighted cell and find all occurences in all sheets and list them.
'
    Dim SearchString As String
    Dim SearchRange As Range, cl As Range
    Dim FirstFound As String
    Dim sh As Worksheet
    Dim results As String
    
    SearchString = ActiveCell.Value
    Debug.Print "++++++++++++++++"
    For Each sh In ActiveWorkbook.Worksheets
        ' Ignore all sheets with Link or Sand in the name
        If InStr(sh.Name, "Link") = 0 And InStr(sh.Name, "Sand") = 0 Then
            Set cl = sh.Cells.Find(What:=SearchString, _
                After:=sh.Cells(1, 1), _
                LookIn:=xlValues, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False, _
                SearchFormat:=False)
            If Not cl Is Nothing Then
                ' if found, remember location
                FirstFound = cl.Address
                ' format found cell
                Do
                    Debug.Print sh.Name, cl, cl.Address
                    results = results & " " & sh.Name & " " & cl & " " & cl.Address & vbLf
                    
                    ' find next instance
                    Set cl = sh.Cells.FindNext(After:=cl)
                    ' repeat until back where we started
                Loop Until FirstFound = cl.Address
            End If
        End If
    Next
    ' Display results
    MsgBox results
    
End Sub
Sub GotoRef()
' Take user to linked reference.
' Cursor must  be in cell containing reference.

    Dim ws As Worksheet
    Dim ReturnSheet As Worksheet
    Dim Refs() As String
    
    ' Variables for initial sheet
    Dim strAFilterRng As String    ' Autofilter range
    Dim varFilterCache()           ' Autofilter cache
    
    ' Variables for each sheet to be searched
    Dim strSearchAFilterRng As String    ' Autofilter range
    Dim varSearchFilterCache()           ' Autofilter cache
    
    SearchRef = ActiveCell.Value
    ' Remove all spaces
    SearchRef = Replace(SearchRef, " ", "")
    ' Look to see if cell contains multiple refs
    If InStr(SearchRef, ",") > 0 Then
        ' Split cell into separate refs
        Refs = Split(SearchRef, ",")
        ' Build menu
        For i = 0 To UBound(Refs)
            Menu = Menu & i + 1 & ". " & Refs(i) & vbLf
        Next i
        ' Show menu and allow choice
        Do Until (Choice > 0) And (Choice <= i)
            Ch = InputBox(Menu, "Choose option:", 1)
            ' Check for Escape key
            If Ch = "" Then Exit Sub Else Choice = Int(Ch)
        Loop
        SearchRef = Refs(Choice - 1)
    End If
    ' Check if blank and then set a string that is impossible to find.
    If SearchRef = "" Then SearchRef = "*BLANK*"
    Set ReturnSheet = ActiveSheet
    ReturnAddress = ActiveCell.Address
    
    ' Make a note of req ID
    ReqId = Cells(ActiveCell.Row, 1)
    
    ' Capture AutoFilter settings
    ' Check for autofilter, turn off if active..
    SaveFilters ReturnSheet, strAFilterRng, varFilterCache
    
    Found = False
    
    For Each ws In ActiveWorkbook.Worksheets
        ' Ignore all sheets with Link or Sand in the name
        If InStr(ws.Name, "Link") = 0 And InStr(ws.Name, "Sand") = 0 Then
            ' Remove any auto filters otherwise results will not be found
            On Error Resume Next
            Worksheets(ws.Name).ShowAllData
            On Error GoTo 0
            
            Set cl = ws.Cells.Find(What:=SearchRef, _
                After:=ws.Cells(1, 1), _
                LookIn:=xlFormulas, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False, _
                SearchFormat:=False)
                
            If Not cl Is Nothing Then
                FirstFound = cl.Address
                Do
                    ' Look for ref in column A only
                    If Left(cl.Address, 3) = "$A$" Then
                        ws.Activate
                        ws.Range(cl.Address).Activate
                        Found = True
                    End If
                    ' find next instance
                    Set cl = ws.Cells.FindNext(After:=cl)
                    ' repeat until back where we started
                Loop Until FirstFound = cl.Address
            End If
        End If
    Next

    If Found Then
        ' Check that source req ID appears in row, i.e. link is good.
        lnRow = ActiveCell.Row
        On Error Resume Next
        lnCol = Cells(lnRow, 1).EntireRow.Find(What:=ReqId, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
        If Err.Number <> 0 Then MsgBox "This requirement is missing its link to " & ReqId
        On Error GoTo 0
        
        Reply = MsgBox("Click OK to return to original requirement" & vbLf & "Cancel to remain here", vbOKCancel)
        If Reply = vbOK Then ReturnSheet.Activate
    Else
        MsgBox "Ref not found"
    End If
    ' Restore Filter settings
    ' Restore original autofilter if present ..
    RestoreFilters ReturnSheet, strAFilterRng, varFilterCache
End Sub
Sub ShowRef()
' Take ref in current highlighted cell and find the source ref and show the requirement.
'
    Dim SearchString As String
    Dim rLastCell As Range, ranrngSearchRange As Range, cl As Range
    Dim FirstFound As String
    Dim ws As Worksheet
    Dim Req As String, results As String
    
    SearchRef = ActiveCell.Value
    Debug.Print "++++++++++++++++"
    For Each ws In ActiveWorkbook.Worksheets
        ' Ignore all sheets with Link or Sand in the name
        If InStr(ws.Name, "Link") = 0 And InStr(ws.Name, "Sand") = 0 Then
        
            Set cl = ws.Cells.Find(What:=SearchRef, _
                After:=ws.Cells(1, 1), _
                LookIn:=xlValues, _
                LookAt:=xlPart, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False, _
                SearchFormat:=False)
                
            If Not cl Is Nothing Then
                ' if found, remember location
                FirstFound = cl.Address
                
                ' Find Req Column
                
                Set rLastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, MatchCase:=False)

                col = 0
                Do
                    col = col + 1
                Loop Until Cells(2, col) = "Requirement:" Or col > rLastCell.Column
                
                If col <= rLastCell.Column Then
                    Do
                        ' Look for ref in column A only
                        If Left(cl.Address, 3) = "$A$" Then
                            Req = ws.Cells(Range(cl.Address).Row, col).Value
                            Debug.Print ws.Name, Ref, Req
                            results = results & ws.Name & " " & SearchRef & " " & Req & vbLf
                        End If
                        ' find next instance
                        Set cl = ws.Cells.FindNext(After:=cl)
                        ' repeat until back where we started
                    Loop Until FirstFound = cl.Address
                End If
            End If
        End If
    Next
    ' Display results
    MsgBox results
    
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

Dim cl As Range
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
    For Each cl In Selection
      TotalLen = Len(SearchText)
      StartPos = InStr(cl, SearchText)
      TestPos = 0
      Do While StartPos > TestPos
        With cl.Characters(StartPos, TotalLen).Font
          .FontStyle = "Bold"
          .ColorIndex = 3
        End With
        EndPos = StartPos + TotalLen
        TestPos = TestPos + EndPos
        StartPos = InStr(TestPos, cl, SearchText, vbTextCompare)
      Loop
    Next cl
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
' Usage example:
'    Dim strAFilterRng As String    ' Autofilter range
'    Dim varFilterCache()           ' Autofilter cache
'    ' [set up code]
'    Set wksAF = Worksheets("Configuration")
'
'    ' Check for autofilter, turn off if active..
'    SaveFilters wksAF, strAFilterRng, varFilterCache
'    [code with filter off]
'    [set up special auto-filter if required]
'    [code with filter on as applicable]
'    ' Restore original autofilter if present ..
'    RestoreFilters wksAF, strAFilterRng, varFilterCache

'~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Sub:      SaveFilters
' Purpose:  Save filter on worksheet
' Returns:  wks.AutoFilterMode when function entered
'
' Arguments:
'   [Name]      [Type]  [Description]
'   wks         I/P     Worksheet that filter may reside on
'   FilterRange O/P     Range on which filter is applied as string; "" if no filter
'   FilterCache O/P     Variant dynamic array in which to save filter
'
' Author:   Based on MS Excel AutoFilter Object help file
'
' Modifications:
' 2006/12/11 Phil Spencer: Adapted as general purpose routine
' 2007/03/23 PJS: Now turns off .AutoFilterMode
' 2013/03/13 PJS: Initial mods for XL14, which has more operators
'
' Comments:
'----------------------------
Function SaveFilters(wks As Worksheet, FilterRange As String, FilterCache()) As Boolean
    Dim ii As Long

    FilterRange = ""    ' Alternative signal for no autofilter active
    SaveFilters = wks.AutoFilterMode
    If SaveFilters Then
        With wks.AutoFilter
            FilterRange = .Range.Address
            With .Filters
                ReDim FilterCache(1 To .Count, 1 To 3)
                For ii = 1 To .Count
                    With .Item(ii)
                        If .On Then
#If False Then ' XL11 code
                            FilterCache(ii, 1) = .Criteria1
                            If .Operator Then
                                FilterCache(ii, 2) = .Operator
                                FilterCache(ii, 3) = .Criteria2
                            End If
#Else   ' first pass XL14
                            Select Case .Operator

                            Case 1, 2   'xlAnd, xlOr
                                FilterCache(ii, 1) = .Criteria1
                                FilterCache(ii, 2) = .Operator
                                FilterCache(ii, 3) = .Criteria2

                            Case 0, 3 To 7 ' no operator, xlTop10Items, _
 xlBottom10Items, xlTop10Percent, xlBottom10Percent, xlFilterValues
                                FilterCache(ii, 1) = .Criteria1
                                FilterCache(ii, 2) = .Operator

                            Case Else    ' These are not correctly restored; there's someting in Criteria1 but can't save it.
                                FilterCache(ii, 2) = .Operator
                                ' FilterCache(ii, 1) = .Criteria1   ' <-- Generates an error
                                ' No error in next statement, but couldn't do restore operation
                                ' Set FilterCache(ii, 1) = .Criteria1

                            End Select
#End If
                        End If
                    End With ' .Item(ii)
                Next
            End With ' .Filters
        End With ' wks.AutoFilter
        wks.AutoFilterMode = False  ' turn off filter
    End If ' wks.AutoFilterMode
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Sub:      RestoreFilters
' Purpose:  Restore filter on worksheet
' Arguments:
'   [Name]      [Type]  [Description]
'   wks         I/P     Worksheet that filter resides on
'   FilterRange I/P     Range on which filter is applied
'   FilterCache I/P     Variant dynamic array containing saved filter
'
' Author:   Based on MS Excel AutoFilter Object help file
'
' Modifications:
' 2006/12/11 Phil Spencer: Adapted as general purpose routine
' 2013/03/13 PJS: Initial mods for XL14, which has more operators
'
' Comments:
'----------------------------
Sub RestoreFilters(wks As Worksheet, FilterRange As String, FilterCache())
    Dim col As Long

    wks.AutoFilterMode = False ' turn off any existing auto-filter
    If FilterRange <> "" Then
        wks.Range(FilterRange).AutoFilter ' Turn on the autofilter
        For col = 1 To UBound(FilterCache(), 1)

#If False Then  ' XL11
            If Not IsEmpty(FilterCache(col, 1)) Then
                If FilterCache(col, 2) Then
                    wks.Range(FilterRange).AutoFilter field:=col, _
                        Criteria1:=FilterCache(col, 1), _
                            Operator:=FilterCache(col, 2), _
                        Criteria2:=FilterCache(col, 3)
                Else
                    wks.Range(FilterRange).AutoFilter field:=col, _
                        Criteria1:=FilterCache(col, 1)
                End If
            End If
#Else

            If Not IsEmpty(FilterCache(col, 2)) Then
                Select Case FilterCache(col, 2)

                Case 0  ' no operator
                    wks.Range(FilterRange).AutoFilter field:=col, _
                        Criteria1:=FilterCache(col, 1) ' Do NOT reload 'Operator'

                Case 1, 2   'xlAnd, xlOr
                    wks.Range(FilterRange).AutoFilter field:=col, _
                        Criteria1:=FilterCache(col, 1), _
                        Operator:=FilterCache(col, 2), _
                        Criteria2:=FilterCache(col, 3)

                Case 3 To 6 ' xlTop10Items, xlBottom10Items, xlTop10Percent, xlBottom10Percent
#If True Then
                    wks.Range(FilterRange).AutoFilter field:=col, _
                        Criteria1:=FilterCache(col, 1) ' Do NOT reload 'Operator' , it doesn't work
                    ' wks.AutoFilter.Filters.Item(col).Operator = FilterCache(col, 2)
#Else ' Trying to restore Operator as well as Criteria ..
                    ' Including the 'Operator:=' arguement leads to error.
                    ' Criteria1 is expressed as if for a FALSE .Operator
                    wks.Range(FilterRange).AutoFilter field:=col, _
                        Criteria1:=FilterCache(col, 1), _
                        Operator:=FilterCache(col, 2)
#End If

                Case 7  'xlFilterValues
                    wks.Range(FilterRange).AutoFilter field:=col, _
                        Criteria1:=FilterCache(col, 1), _
                        Operator:=FilterCache(col, 2)

#If False Then ' Switch on filters on cell formats
' These statements restore the filter, but cannot reset the pass Criteria, so the filter hides all data.
' Leave it off instead.
                Case Else   ' (Various filters on data format)
                    wks.Range(FilterRange).AutoFilter field:=col, _
                        Operator:=FilterCache(col, 2)
#End If ' Switch on filters on cell formats

                End Select
            End If

#End If     ' XL11 / XL14
        Next col
    End If
End Sub

