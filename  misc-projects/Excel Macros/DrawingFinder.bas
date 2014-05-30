Attribute VB_Name = "DrawingFinder"
'Option Explicit
'Must select Tools-Microsoft Runtime
'Use late binding objects to allow for different versions of Excel.

Const Build As String = 23
Const DebugMode = True
Const ForceLocal = False

' Define globals
Public GlobalNum As String
Public GlobalDesc As String
Public GlobalFinderProgramFile As String
Public GlobalProgramPath As String
Public GlobalCurrentIssueFolder As String
Public GlobalOldIssueFolder As String
Public GlobalTransferFolder As String
Public GlobalCurrentIndexFile As String
Public GlobalTransferIndexFile As String
Public GlobalOldIndexFile As String
Public GlobalResultFile As String
Public GlobalBatchFile As String
Public GlobalLogFile As String
Public Globalfilepath As String
Public Globaldrive As String
Public GlobalTutorialFile As String
Public CurrentIndexArray() As String
Public OldIndexArray() As String
Public GlobalTreeRoot As String
Public GlobalLowestBOM As String
Public GlobalWorkbook As Workbook
Public GlobalDoc As Object
Public GlobalMaxLevel As Integer
Public GlobalFileOpener As String

Enum WhatIsIt
    BOM
    DRG
    Mat
    OTH
End Enum

Enum AppType
    Word
    Excel
End Enum

Public Type DrawingType
    Number As String
    Is As WhatIsIt
    Issue As String
    Title As String
End Type

Const ShowDurationSecs As Integer = 5

Const NetProgramPath = "\\atle.bombardier.com\data\uk\pl\dos\Drgstate\"
Const NetTransferPath = "\\atle.bombardier.com\data\uk\pl\dos\"
Const NetDataPath = "\\atle.bombardier.com\data\uk\pl\dos2\"
Const LocalProgramPath = "Drgstate\"
Const LocalTransferPath = ""
Const LocalLogPath = "c:\windows\temp\"
Const LocalDataPath = "\"
Const DesignNoteFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DesignNoteStateSAP.xlsx"
Const DrawingFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DrgStateSAP.xlsx"
Const TreeRoot = "TreeRoot\"

Enum RequestType
    Current = 1
    Old = 2
    ECR = 3
    Tree = 4
End Enum

Enum ActionType
    OpenInApp = 1
    ShowInFolder = 2
End Enum

Dim fso As New FileSystemObject
Dim fld As Folder

Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
ByVal hwnd As Long, _
ByVal lpText As String, _
ByVal lpCaption As String, _
ByVal uType As Long, _
ByVal wLanguageID As Long, _
ByVal lngMilliseconds As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
Public Sub SetGlobals()
' Global variables for use throughout the program
    
    Dim DataPath As String
    Dim TransferPath As String
    Dim Drive As String
    
    On Error GoTo ErrorHandler

    ' Find out whether network drive is connected
    If Not ForceLocal And DirExists(NetDataPath) Then
        DataPath = NetDataPath
        GlobalProgramPath = NetProgramPath
        TransferPath = NetTransferPath
    Else
        Drive = Switch(DirExists("e:\1_current_iss"), "e:", DirExists("f:\1_current_iss"), "f:", DirExists("g:\1_current_iss"), "g:", DirExists("c:\1_current_iss"), "c:\", True, "Not Found")
        DataPath = Drive
        GlobalProgramPath = Drive & LocalProgramPath
        TransferPath = Drive & LocalTransferPath
    End If
    
    ' Test for no suitable directory found
    If Drive = "Not Found" Then
        MsgBox ("Current Issue" & vbLf & "Folder not found")
        End
    End If

    ' Assign Prorgam related variables
    GlobalFinderProgramFile = "DrawingFinder.xls"
    GlobalCurrentIndexFile = GlobalProgramPath & "CurrentIndex.txt"
    GlobalOldIndexFile = GlobalProgramPath & "OldIndex.txt"
    GlobalBatchFile = GlobalProgramPath & "CreateIndex.bat"
    GlobalTutorialFile = GlobalProgramPath & "DrawingFinderTutorial.pdf"
        GlobalFileOpener = GlobalProgramPath & "FileOpener.bat"
    
    ' Assign Data related variables
    GlobalCurrentIssueFolder = DataPath & "1_current_iss"
    GlobalOldIssueFolder = DataPath & "1_old_iss"
    
    GlobalTransferFolder = TransferPath & "1_files for filing"
    GlobalResultFile = LocalLogPath & "DrawingFinderResult.txt"
    GlobalTransferIndexFile = LocalLogPath & "DrawingFinderTransferIndex.txt"
    GlobalTreeRoot = LocalLogPath & TreeRoot
    
    ' Assign Log file path
    ' Select local log file if user doesn't have write access to network log file
    If Not FileExists(GlobalProgramPath & "DrawingFinderLogFile.txt") Then
        GlobalLogFile = LocalLogPath & "DrawingFinderLogFile.txt"
    Else
        GlobalLogFile = GlobalProgramPath & "DrawingFinderLogFile.txt"
    End If

Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: SetGlobal Error number=" & CStr(Err))
    End
End Sub
Private Sub PlantTree()
'Get all the drawings/materials from the the current open drawing.
'Use this list to form a linked list of all the sub level BOMs.
'Look for sub level BOMs and open if Word or Excel
'Create a folder & file structure to represent the data

    Dim WhatItIs As String
    Dim Index As Integer
    Dim fs As New FileSystemObject
    Dim TopLevelBOM As DrawingType
    Dim TopLevel As Integer
    Dim TopLevelFolder As String
    
    On Error GoTo ErrorHandler
    
    If DebugMode Then
        Debug.Print
        Debug.Print "---Start---"
    End If
       
    Call MsgBoxDelay("Building BOM Tree...", "Please Wait", ShowDurationSecs)
    
    Application.ScreenUpdating = False
    SetGlobals
    
    ChDrive "c:"
    MakeDirectory (GlobalTreeRoot)
    ChDir GlobalTreeRoot
    
    'Get top level BOM
    ' Read cell contents
    TopLevelBOM.Number = Cells(ActiveCell.row, 1).Value
    Issue = Cells(ActiveCell.row, 3).Value
    Correction = Cells(ActiveCell.row, 4).Value
    ECRnum = Cells(ActiveCell.row, 6).Value
    Title = Cells(ActiveCell.row, 2).Value
    TopLevel = 0 'Level of the hierarchy
    GlobalMaxLevel = 0
    
    Call LogInformation("PlantTree: TopLevelBOM:" & TopLevelBOM.Number)
    Call FindInfo(TopLevelBOM.Number, TopLevelBOM.Number, Issue:=TopLevelBOM.Issue, Title:=TopLevelBOM.Title)
                    
    'Check that it is a BOM
    If TopLevelBOM.Is = BOM Then
    
        TopLevelFolder = TopLevelBOM.Number & "-" & TopLevelBOM.Issue & " " & TopLevelBOM.Title
        TopLevelFolder = Replace(TopLevelFolder, "/", "-")
        'Clear out old tree
        KillDirs (TopLevelFolder)
        MakeDirectory (TopLevelFolder)
        'Make file version of BOM to allow it to be opened direct from BOM tree.
        ChDir (TopLevelFolder)
        With TopLevelBOM
            Call MakeFile(.Number, Left("~" & .Number & "-" & .Issue & " " & .Title, 44), "BOM")
            ChDir ".."
        End With

        Call BuildTree(fs.GetFolder(TopLevelFolder), TopLevel)
        
        If DebugMode Then
            Debug.Print
            Debug.Print "---Finish---"
        End If
        
        Shell "explorer /e, /root, " & GlobalTreeRoot, vbNormalFocus   'Show root folder
        
        'Release folder
        ChDir LocalLogPath
    '   Stop
    Else
        Call MsgBoxDelay("Drawing is not a BOM...", "DrawingTree", ShowDurationSecs)
    End If
    
    Application.ScreenUpdating = True
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: PlantTree Error number=" & CStr(Err))
    End
End Sub
Sub BuildTree(ByVal SubLevelBOM As Folder, ByRef Level As Integer)

    Dim fs As New FileSystemObject
    Dim FSfolder As Folder
    Dim SubFolder As Folder
    Dim CurrentBOMDoc As String
    Dim Item As String
    Dim IndexFile As String
    Dim DrawingList() As DrawingType
    Dim MaxDrawings As Integer
    Dim WhatApp As AppType
    Dim AllRev As Object
    Dim WhatItIs As String
    
    On Error GoTo ErrorHandler
    
    'Strip BOM name from path
    Item = fs.GetFilename(SubLevelBOM)
    ChDir SubLevelBOM
    
    Level = Level + 1 'Increase level of hierarchy
    
    'Find the BOM, open it and extract the drawings/materials.
    
    'Only strip off issue and title if - appears later on in string.
    i = 4
    Finished = False
    Do
        i = i + 1
        If Mid(Item, i, 1) = "-" Then
            Item = Left(Item, i - 1)
            Finished = True
        End If
    Loop Until i = Len(Item) Or Finished
    
    IndexFile = GlobalCurrentIndexFile
    Call CreateResultFile(Item, IndexFile)
    IndexFile = GlobalResultFile
    CurrentBOMDoc = MsOfficeDoc(IndexFile)
    
    If CurrentBOMDoc = "" Then
        MsgBox ("BOM " & Item & " not found")
    Else
        'Detect Word/Excel and open document
        If InStr(UCase(CurrentBOMDoc), "XLS") Then
            If DebugMode Then Debug.Print "Opening ExcelDoc", fs.GetFilename(CurrentBOMDoc)
            Set DocApp = CreateObject("Excel.Application")
            Set GlobalWorkbook = DocApp.Workbooks.Open(CurrentBOMDoc, ReadOnly:=True)
            DocApp.Visible = False
            WhatApp = Excel
        Else
            If DebugMode Then Debug.Print "Opening WordDoc", fs.GetFilename(CurrentBOMDoc)
            Set DocApp = CreateObject("word.Application")
            Set GlobalDoc = DocApp.Documents.Open(CurrentBOMDoc, ReadOnly:=True)
            DocApp.Visible = False
            WhatApp = Word
            
            'Accept all changes
            With GlobalDoc
                For Each AllRev In .Revisions
                    AllRev.Accept
                Next AllRev
            End With
        End If
        
        'Get a list of all the drawings/materials
        Call GetAllDrawings(WhatApp, Refs:=DrawingList, Occupied:=MaxDrawings)
        ReDim Preserve DrawingList(MaxDrawings) As DrawingType
        'Call QuickSort(DrawingList, LBound(DrawingList), UBound(DrawingList))
        
        For Index = 1 To UBound(DrawingList)
            Item = DrawingList(Index).Number
            WhatItIs = DrawingList(Index).Is
            
            Select Case WhatItIs
                Case 0
                    WhatItIs = "BOM"
                Case 1
                    WhatItIs = "Drawing"
                Case 2
                    WhatItIs = "Material"
            End Select
            
            Select Case DrawingList(Index).Is
                Case BOM
                    Call FindInfo(CurrentBOMDoc, Item, Issue:=DrawingList(Index).Issue, Title:=DrawingList(Index).Title)
                    MakeDirectory (Left(Item & "-" & DrawingList(Index).Issue & " " & DrawingList(Index).Title, 44))
                    'Make file version of BOM to allow it to be opened direct from BOM tree.
                    'ChDir (Left(Item & "-" & DrawingList(Index).Issue & " " & DrawingList(Index).Title, 44)) 'This does not always work correctly
                    Call MakeFile(Item, Left("~" & Item & "-" & DrawingList(Index).Issue & " " & DrawingList(Index).Title, 44), WhatItIs)
                    'ChDir ".."
                    
                    IndexFile = GlobalCurrentIndexFile
                    Call CreateResultFile(Item, IndexFile)
                    IndexFile = GlobalResultFile
                    NewDoc = MsOfficeDoc(IndexFile)
                    If DebugMode Then Debug.Print "BOM", Item, fs.GetFilename(NewDoc)
                Case DRG
                    Call FindInfo(CurrentBOMDoc, Item, Issue:=DrawingList(Index).Issue, Title:=DrawingList(Index).Title)
                    Call MakeFile(Item, Left(Item & "-" & DrawingList(Index).Issue & " " & DrawingList(Index).Title, 44), WhatItIs)
                Case Mat
                    'Create file if not OTH
                    'MakeFile (Item & "." & WhatItIs)
                    Call MakeFile(Item, Item, WhatItIs)
                Case OTH
                    'Nothing to do
            End Select
            
        Next Index
        
        'Detect Word/Excel and close document
        If InStr(UCase(CurrentBOMDoc), "XLS") Then
            If DebugMode Then Debug.Print "Closing ExcelDoc", fs.GetFilename(CurrentBOMDoc)
            GlobalWorkbook.Saved = True   'Prevent do you want to save message
            DocApp.Workbooks.Close
            DocApp.Quit
            Set DocApp = Nothing
        Else
            If DebugMode Then Debug.Print "Closing WordDoc", fs.GetFilename(CurrentBOMDoc)
            GlobalDoc.Saved = True
            DocApp.Documents.Close
            DocApp.Quit wdDoNotSaveChanges
            Set DocApp = Nothing
        End If
        
        'Recursively build the tree
        Set FSfolder = fs.GetFolder(SubLevelBOM)
        For Each SubFolder In FSfolder.SubFolders
            Call BuildTree(SubFolder, Level)
            If GlobalMaxLevel < Level Then
                GlobalMaxLevel = Level      'Set newmax level
                GlobalLowestBOM = SubFolder 'Keep track of lowest level BOM for best Win Explorer view.
            End If
        Next SubFolder
    End If
    Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: BuildTree Error number=" & CStr(Err))
    End
End Sub
Public Sub FindInfo(ByVal BOM As String, ByVal SearchString As String, ByRef Issue As String, ByRef Title As String)
    'Look for the issue and title in the spreadsheet
    
    Dim fs As New FileSystemObject
    
    On Error GoTo ErrorHandler
    
    Set sh = ThisWorkbook.Worksheets(1)
    'Find first instance on sheet
    'c = Application.WorksheetFunction.Match(SearchString, ActiveWorkbook.Sheets(1).Range("A1", "A999"), 1)
        
    'Used to create a new view where all cells are shown to allow FIND to work properly.
    On Error Resume Next
    ThisWorkbook.CustomViews("FindView").Show
    On Error GoTo 0
        
    With sh ' CodeName of filtered sheet
        Application.GoTo .Range("A1")
        ThisWorkbook.CustomViews.Add "FindView", False, True
        .AutoFilterMode = False
        
        Set cl = sh.Cells.Find(What:=SearchString, _
            After:=sh.Cells(1, 1), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False)
         ThisWorkbook.CustomViews("FindView").Show
    End With
    
    If cl Is Nothing Then
        Issue = ""
        Title = ""
        'Strip BOM name from path
        BOM = fs.GetFilename(BOM)
        MsgBox ("Check drawing " & SearchString & vbLf & "in BOM " & BOM)
        Call LogInformation("FindInfo: Drawing number error:" & SearchString & " in BOM:" & BOM)
    Else
        Issue = sh.Cells(Range(cl.Address).row, 3).Value
        Title = sh.Cells(Range(cl.Address).row, 2).Value
    End If
    Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: FindInfo Error number=" & CStr(Err))
    End
End Sub
Public Sub GetAllDrawings(WhatApp As AppType, ByRef Refs() As DrawingType, ByRef Occupied As Integer)
'Compile an array of all the drawing/material numbers

    Dim aTable As Object
    Dim aCell As Object
    Dim aRow As Integer
    Dim DrawingRowStart As Integer
    Dim DrawingColStart As Integer
    Dim ActiveRow As Integer
    Dim RefArray() As String    'Need to use this for Excel to prevent error of using user defined type.

    On Error GoTo ErrorHandler
    
    If WhatApp = Excel Then
        DrawingRowStart = Range(StartOfDrawings).row
        DrawingColStart = Range(StartOfDrawings).Column

        With GlobalWorkbook.Worksheets(1)
            MaxRows = .Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
            ReDim RefArray(MaxRows)
            ReDim Refs(MaxRows)
            
            Occupied = 1
            For ActiveRow = DrawingRowStart To MaxRows
                RefArray(Occupied) = GlobalWorkbook.Worksheets(1).Cells(ActiveRow + 1, DrawingColStart)
                If RefArray(Occupied) <> "" Then
                    RefArray(Occupied) = OnlyAlphaNumericChars(RefArray(Occupied))
                    Occupied = Occupied + 1
                End If
            Next ActiveRow
            
            Occupied = Occupied - 1
            'Copy array into user defined array
            For i = 1 To Occupied
                Refs(i).Number = RefArray(i)
                Refs(i).Is = IsDrawingType(Refs(i).Number)
            Next i
        End With
    Else
    
        With GlobalDoc.Tables(1)
            MaxRows = .Range.Rows.Count
            
            ReDim Refs(MaxRows)
            
            Occupied = 0
            For aRow = 1 To MaxRows - 1
                Set aCell = .cell(aRow + 1, 2)
                Refs(aRow).Number = OnlyAlphaNumericChars(aCell.Range)
                If Refs(aRow).Number <> "" Then
                    Occupied = Occupied + 1
                    Refs(Occupied).Number = Refs(aRow).Number   'Use Occupied index to filter out empty elements
                    Refs(Occupied).Is = IsDrawingType(Refs(aRow).Number)
                End If
            Next aRow
            'Occupied = Occupied - 1
        End With
    End If
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: GetAllDrawings Error number=" & CStr(Err))
    End
End Sub
Function StartOfDrawings() As String
'Find start of drawing/material list

    Dim SearchString As String
    Dim SearchRange As Range, cl As Range
    Dim FirstFound As String
    Dim sh As Worksheet

    On Error GoTo ErrorHandler
    
    'Set Search value
    SearchString = "SAP"
    'Application.FindFormat.Clear
    'loop through all sheets
    For Each sh In GlobalWorkbook.Worksheets
        'Find first instance on sheet
        Set cl = sh.Cells.Find(What:=SearchString, _
            After:=sh.Cells(1, 1), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False)
        If Not cl Is Nothing Then
            'if found, remember location
            FirstFound = cl.Address
        End If
    Next
    StartOfDrawings = FirstFound
Exit Function

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: StartOfDrawings Error number=" & CStr(Err))
    End
End Function
Private Function OnlyAlphaNumericChars(OrigString As String) As String
'Remove unwanted characters

    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    
    On Error GoTo ErrorHandler
    
    OrigString = Trim(OrigString)
    lLen = Len(OrigString)
    For lCtr = 1 To lLen
        sChar = Mid(OrigString, lCtr, 1)
        If IsAlphaNumeric(Mid(OrigString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    DoEvents '(optional, but if processing long string,
    'necessary to prevent program from appearing to hang)
    'if used, write your app so no re-entrancy into this function
    'can occur)
    Next
        
    OnlyAlphaNumericChars = sAns

Exit Function

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: OnlyAlphaNumericChars Error number=" & CStr(Err))
    End
End Function
Private Function IsAlphaNumeric(sChr As String) As Boolean
'Check that charcter is in acceptable list

    IsAlphaNumeric = sChr Like "[0-9A-Za-z,-,/]"
End Function
Function IsDrawingType(Item As String) As WhatIsIt
'Return the type of drawing, BOM, DWG, MAT or OTH
'Determine whether Item is a BOM. Look for new parts lists L52xxxxxxx or old SXL & GXL numbers.

    On Error GoTo ErrorHandler
    
    If (Left(Item, 3) = "L52") Or (Item Like "*SXL*") Or (Item Like "*GXL*") Then
        IsDrawingType = BOM
    ElseIf (Len(Item) = 6 And Left(Item, 1) = "1") Or (Len(Item) = 9 And Left(Item, 2) = "52") Then
            IsDrawingType = Mat
    ElseIf (UCase(Item) Like "*FITTED*") Then
        IsDrawingType = OTH
    Else
        IsDrawingType = DRG
    End If

    'If DebugMode Then Debug.Print "WhatIsIt", Item, IsDrawingType
Exit Function

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: IsDrawingType Error number=" & CStr(Err))
    End
End Function
Sub FilterSheet()


' Set up data filters on Item No. and Description columns.
' Enter up to 2 words in each search box, these will be OR'd

    Dim NumWords() As String
    Dim DescWords() As String
    
    Dim NumWordCount, DescWordCount As Integer
    
    On Error GoTo ErrorHandler
    
    NumWordCount = 4
    DescWordCount = 4
    
    Do Until (NumWordCount < 3)
        ' Get search data from user & convert to upper case
        GlobalNum = StrConv(InputBox("Enter part of the Drawing Number" & vbLf & "Up to 2 words can be entered" & vbLf _
        & "Use & | for AND OR", "Drawing Search", GlobalNum), vbUpperCase)
        GlobalNum = Replace(GlobalNum, " ", "&")
        ' Split the input strings into an array of words.
        If InStr(1, GlobalNum, "&") > 0 Then
            NumWords = Split(GlobalNum, "&")
        ElseIf InStr(1, GlobalNum, "|") > 0 Then
            NumWords = Split(GlobalNum, "|")
        Else: NumWords = Split(GlobalNum)
        End If
        If GlobalNum = "" Then GlobalNum = "*"
        ' Return the number of elements in the arrays.
        NumWordCount = UBound(NumWords) - LBound(NumWords) + 1
    Loop

    Do Until (DescWordCount < 3)
        ' Get search data from user & convert to upper case
        GlobalDesc = StrConv(InputBox("Enter part of the Drawing Description" & vbLf & "Up to 2 words can be entered" & vbLf _
        & "Use & | for AND OR", "Drawing Search", GlobalDesc), vbUpperCase)
        GlobalDesc = Replace(GlobalDesc, " ", "&")
        If InStr(1, GlobalDesc, "&") > 0 Then
            DescWords = Split(GlobalDesc, "&")
        ElseIf InStr(1, GlobalDesc, "|") > 0 Then
            DescWords = Split(GlobalDesc, "|")
        Else: DescWords = Split(GlobalDesc)
        End If
        If GlobalDesc = "" Then GlobalDesc = "*"
        DescWordCount = UBound(DescWords) - LBound(DescWords) + 1
    Loop
        
    SetGlobals
    ' Write search terms to log
    Call LogInformation("Filter: Drawing: " & GlobalNum & " Description: " & GlobalDesc)
        
    Range("A7").Select
    
    If GlobalNum = "*" Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=1, _
        Criteria1:="=*", Operator:=xlAnd
    ElseIf InStr(1, GlobalNum, "&") = 0 And InStr(1, GlobalNum, "|") = 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=1, _
        Criteria1:="=*" & NumWords(0) & "*", Operator:=xlAnd
    ElseIf InStr(1, GlobalNum, "&") > 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=1, _
        Criteria1:="=*" & NumWords(0) & "*", Operator:=xlAnd, Criteria2:="=*" & NumWords(1) & "*"
    ElseIf InStr(1, GlobalNum, "|") > 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=1, _
        Criteria1:="=*" & NumWords(0) & "*", Operator:=xlOr, Criteria2:="=*" & NumWords(1) & "*"
    End If
    
    If GlobalDesc = "*" Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=2, _
        Criteria1:="=*", Operator:=xlAnd
    ElseIf InStr(1, GlobalDesc, "&") = 0 And InStr(1, GlobalDesc, "|") = 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=2, _
        Criteria1:="=*" & DescWords(0) & "*", Operator:=xlAnd
    ElseIf InStr(1, GlobalDesc, "&") > 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=2, _
        Criteria1:="=*" & DescWords(0) & "*", Operator:=xlAnd, Criteria2:="=*" & DescWords(1) & "*"
    ElseIf InStr(1, GlobalDesc, "|") > 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=2, _
        Criteria1:="=*" & DescWords(0) & "*", Operator:=xlOr, Criteria2:="=*" & DescWords(1) & "*"
    End If
    
    ActiveWindow.SmallScroll Down:=-10000
    
    ' Place cursor in top row of results.
    ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 1).Select
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: FilterSheet Error number=" & CStr(Err))
    End
End Sub
Sub ChooseAction()
' Display show menu and allow choice of current issue, old issue or show in folder
    Dim Request As RequestType
    Dim Action As ActionType
    Dim Ch, Index As String
    Dim Choice As Integer
    
    On Error GoTo ErrorHandler
    
    SetGlobals
    
    Choice = -9
    
    Do Until (Choice > 0) And (Choice <= 4)
        Ch = InputBox("1. Current Issue" & vbLf & "2. Old Issue" & vbLf & "3. ECR" & vbLf & "4. BOM Tree", "Choose option:", 1)
        ' Check for Escape key
        If Ch = "" Then Exit Sub Else Choice = Int(Ch)
    Loop
    
    ' Create strings for log entry
        
    Select Case Ch
        Case "1"
            ChoiceStr = "Current"
            Index = GlobalCurrentIndexFile
            Request = Current
        Case "2"
            ChoiceStr = "Old"
            Index = GlobalOldIndexFile
            Request = Old
        Case "3"
            ChoiceStr = "ECR"
            Index = GlobalCurrentIndexFile
            Request = ECR
        Case "4"
            ChoiceStr = "Tree"
            Index = GlobalCurrentIndexFile
            Request = Tree
    End Select
    
    Choice = -9
    
    Do Until ((Choice > 0) And (Choice <= 2)) Or (Request = Tree)
        Ch = InputBox("1. Open document" & vbLf & "2. Show in folder", "Choose action:", 1)
        ' Check for Escape key
        If Ch = "" Then Exit Sub Else Choice = Int(Ch)
    Loop

    Select Case Ch
        Case "1"
            ActionStr = "Open"
            Action = OpenInApp
        Case "2"
            ActionStr = "Show in folder"
            Action = ShowInFolder
    End Select
    
    Call LogInformation("ChooseAction: Choice: " & ChoiceStr & " Action: " & ActionStr)
    
    ' Create new Transfer Index for latest files
    Call CreateIndexFile(GlobalTransferIndexFile, GlobalTransferFolder, True)

    ' Carry out appropriate action
    If Request = Tree Then
        PlantTree
    Else
        If Not ShowItem(Request, Action, Index) Then
            If Request = Current Then
                If Not ShowItem(Request, Action, GlobalTransferIndexFile) Then
                    ' no paths returned from search
                    MsgBox ("File not found")
                    Call LogInformation("ChooseAction: File not found")
                End If
            Else
                ' no paths returned from search
                MsgBox ("File not found")
                Call LogInformation("ChooseAction: File not found")
            End If
        End If
    End If
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: ChooseAction Error number=" & CStr(Err))
    End
End Sub
Function MsOfficeDoc(IndexFile) As String
    'Search for item in index file
    
    Const MaxResults As Integer = 10
    Dim ResultArray(1 To MaxResults) As String
    Dim Index As Integer
    
    On Error GoTo ErrorHandler
    
    FileNum = FreeFile
    MsOfficeDoc = ""
    
    Open IndexFile For Input As #FileNum
    Line = 0
    'While not eof or max array size
    Do Until EOF(FileNum) Or Line = 9
        Input #FileNum, ResultPath
        If InStr(UCase(ResultPath), UCase(Item)) Then
            Line = Line + 1
            ResultArray(Line) = ResultPath
        End If
    Loop
    Close #FileNum

    
    'More than 1 line indicates that at least 1 file has been found
    If Line > MinLines Then
'               ShowItem = True
        For IntIndex = MinLines + 1 To Line
            StrBuf = StrBuf & IntIndex & ". " & Right(GetFilename(ResultArray(IntIndex + MinLines)), 100) & vbLf
        Next
        Choice = -9
        
        'Look for an MS Office document
        Index = 0
        Do
            Index = Index + 1
            If UCase(ResultArray(Index + MinLines)) Like "*DOC*" Then MsOfficeDoc = ResultArray(Index + MinLines)
            If UCase(ResultArray(Index + MinLines)) Like "*XLS*" Then MsOfficeDoc = ResultArray(Index + MinLines)
        Loop Until Index = Line Or MsOfficDoc <> ""
        
    End If
Exit Function

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: MsOfficeDoc Error number=" & CStr(Err))
    End
End Function
Private Sub QuickSort(ByRef Field() As DrawingType, LB As Long, UB As Long)
    Dim P1 As Long, P2 As Long, Ref As DrawingType, TEMP As DrawingType

    On Error GoTo ErrorHandler
    
    P1 = LB
    P2 = UB
    Ref = Field((P1 + P2) / 2)

    Do
        Do While (Field(P1).Number < Ref.Number)
            P1 = P1 + 1
        Loop

        Do While (Field(P2).Number > Ref.Number)
            P2 = P2 - 1
        Loop

        If P1 <= P2 Then
            TEMP = Field(P1)
            Field(P1) = Field(P2)
            Field(P2) = TEMP

            P1 = P1 + 1
            P2 = P2 - 1
        End If
    Loop Until (P1 > P2)

    If LB < P2 Then Call QuickSort(Field, LB, P2)
    If P1 < UB Then Call QuickSort(Field, P1, UB)
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: QuickSort Error number=" & CStr(Err))
    End
End Sub
Function ShowItem(Request As RequestType, Action As ActionType, IndexFile As String) As Boolean
' Search for and then show or open what has been requested

    Const MaxResults As Integer = 10
    Const UseResult As Boolean = True
    
    Dim Cmd As String
    Dim Drawing, Issue, Correction, ECRnum As String
    Dim Reply As Variant
    Dim MinLines As Integer
    Dim Item As String
    
    Dim ResultArray(1 To MaxResults) As String

    Dim StrBuf As String
    Dim IntIndex, Choice As Integer
    Dim TaskId As Long
    Dim Ch As String
    Dim ResultFilePath As String
    Dim FileNum As Integer
    Dim nDirs As Long, nFiles As Long
    Dim lSize As Currency
    
    On Error GoTo ErrorHandler
    
    ShowItem = False
    MinLines = 0
    
    'Report if running on local data
    If ForceLocal Then Call MsgBoxDelay("*** LOCAL ***", "Database Location", ShowDurationSecs)
    
    ' Read cell contents
    Drawing = Cells(ActiveCell.row, 1).Value
    Issue = Cells(ActiveCell.row, 3).Value
    Correction = Cells(ActiveCell.row, 4).Value
    ' Remove non-alpha correction values, such as numeric ones found in SW drawings from the correction number
    If (Correction < "A" Or Correction > "Z") Then Correction = ""
    ECRnum = Cells(ActiveCell.row, 6).Value
    
    ' Find and replace '/' with '-' for file name.
    Drawing = Replace(Drawing, "/", "-")
    
    ' Format ECR number to match filenames
    ECRnum = Replace(ECRnum, "600000000000", "6-")
    ECRnum = Replace(ECRnum, "60000000000", "6-")
    ECRnum = Replace(ECRnum, "6000000000", "6-")
    ECRnum = Replace(ECRnum, "600000000", "6-")
    ECRnum = Replace(ECRnum, "60000000", "6-")
    ECRnum = Replace(ECRnum, "6000000", "6-")
       
    ' Generate full file name for old issue
    ' Create strings for log entry
    Select Case Request
        Case ECR
            Item = ECRnum
            RequestStr = "ECR"
        Case Current
            Item = Drawing
            RequestStr = "Current"
        Case Old
            Item = Drawing & "-" & Issue & Correction
            RequestStr = "Old"
    End Select
    
    ' Check for null string in case the cursor is selecting a non-drawing line
    If Item = "" Then
        MsgBox ("No drawing selected")
        Exit Function
    End If
    
    ' Write seach item to log
    Call LogInformation("ShowItem: Request: " & RequestStr & " : " & Item)

    ' Creating a results file can be quicker using DOS FIND
    If UseResult Then
        Call CreateResultFile(Item, IndexFile)
        IndexFile = GlobalResultFile
    End If
    
    ' Search for item in index file
    FileNum = FreeFile
    
    Open IndexFile For Input As #FileNum
    Line = 0
    ' While not eof or max array size
    Do Until EOF(FileNum) Or Line = 9
        Input #FileNum, ResultPath
        If InStr(UCase(ResultPath), UCase(Item)) Then
            Line = Line + 1
            ResultArray(Line) = ResultPath
        End If
    Loop
    Close #FileNum

    
    ' More than 1 line indicates that at least 1 file has been found
    If Line > MinLines Then
                ShowItem = True
        For IntIndex = MinLines + 1 To Line
            StrBuf = StrBuf & IntIndex & ". " & Right(GetFilename(ResultArray(IntIndex + MinLines)), 100) & vbLf
        Next
        Choice = -9
        
        ' List results to choose from
        Do Until (Choice > MinLines) And (Choice <= Line)
            Ch = InputBox(StrBuf, "Choose File:", 1)
        ' Check for Escape key
        If Ch = "" Then Exit Function Else Choice = Int(Ch)
        Loop
        
        ResultFilePath = ResultArray(Choice)

        If Not FileExists(ResultFilePath) Then
            MsgBox (ResultFilePath & vbLf & "File not found")
            Call LogInformation("ERROR - ShowItem: File not found: " & ResultFilePath)
        Else
            If Action = OpenInApp Then
                Call LogInformation("ShowItem: File path: " & ResultFilePath)
                ' Create full path to file
                ResultFilePath = "file:///" & ResultFilePath
                ' Open file in applicaion
                ActiveWorkbook.FollowHyperlink ResultFilePath
            Else
                Call LogInformation("ShowItem: File path: " & ResultFilePath)
                Shell "explorer /e, /select," & ResultFilePath, vbNormalFocus    ' Use /e to show the folders pane. Opening showing the folders over the network is very slow!
                'Shell "explorer /select," & ResultFilePath, vbNormalFocus
            End If
        End If
    End If
Exit Function

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: ShowItem Error number=" & CStr(Err))
    End
End Function
Private Function FindFile(ByVal sFol As String, sFile As String, _
   nDirs As Long, nFiles As Long) As Currency
   Dim tFld As Folder, tFil As File, FileName As String
   
 '  On Error GoTo Catch
   Set fld = fso.GetFolder(sFol)
   FileName = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or _
                  vbHidden Or vbSystem Or vbReadOnly)
   While Len(FileName) <> 0
      FindFile = FindFile + FileLen(fso.BuildPath(fld.Path, _
      FileName))
      nFiles = nFiles + 1
      List1.AddItem fso.BuildPath(fld.Path, FileName)  ' Load ListBox
      FileName = Dir()  ' Get next file
      DoEvents
   Wend
   Label1 = "Searching " & vbCrLf & fld.Path & "..."
   nDirs = nDirs + 1
   If fld.SubFolders.Count > 0 Then
      For Each tFld In fld.SubFolders
         DoEvents
         FindFile = FindFile + FindFile(tFld.Path, sFile, nDirs, nFiles)
      Next
   End If
   Exit Function
Catch:
    FileName = ""
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: FindFile Error number=" & CStr(Err))
    Resume Next
End Function
Sub Update(Optional UpdateMode As String = "Normal")

    On Error GoTo ErrorHandler
    
    SetGlobals
    
    ' Write build number to cell A1
    Range("A1").Value = "Build: " & Build

    ' Detect why Update has been called and write to log if just started
    Select Case UpdateMode
    Case "Start"
        Call LogInformation("DrawingFinder: +++ Started +++ " & UpdateMode)
        Exit Sub
    
    Case "AutoUpdate"
        Call LogInformation("Update: " & UpdateMode)
        Password = "eng"
    
    Case "Normal"
        ' Password protect process, if running in auto update mode go straight to update.
        Password = InputBox("Enter Password:", "Update Security")
    End Select
    
    
    Select Case Password
        Case "menu"
            RestoreToolbars
    
    ' Import latest SAP data and create a new index
        
        Case "eng"
            ' Write to log file
            Call LogInformation("Update: Starting Scheduled")
            If UpdateMode = "Normal" Then Call LogInformation("Update: Starting Manual")
            ' Provide a method of escaping from the update.
            If MsgBoxDelay("Click OK to cancel", "Update", ShowDurationSecs) = 1 Then
                RestoreToolbars
                ' Write to log file
                Call LogInformation("Update: Aborted by user")
            Else
                Call MsgBoxDelay("Creating Indexes...", "Update", ShowDurationSecs)
                CreateIndexes
                Call MsgBoxDelay("...Indexes Created", "Update", ShowDurationSecs)
                                
                ' Only save workbook if the network drive exists
                If Not ForceLocal And DirExists(NetProgramPath) Then
                    ImportSAP
                    ' Save & set to read-only
                    CheckForArchivedFiles
                    ThisWorkbook.Save
                    SetAttr GlobalProgramPath & GlobalFinderProgramFile, vbReadOnly
                    ' Write to log file
                    Call LogInformation("Update: Finished")
                    CloseSheet
                Else
                    CheckForArchivedFiles
                End If
            End If
            
        Case ""
            Exit Sub
            
        Case "macro"
            ' Show VB editor, Trust access to the VBA project object model check box must be ticked in Excel Options, Trust Centre, Trust Centre Options, Macro Settings
            Application.VBE.MainWindow.Visible = True
        Case Else
            MsgBox "Incorrect password"
    End Select
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: Update Error number=" & CStr(Err))
    End
End Sub
Sub ImportSAP()
' Update Macro
' Merge the Design Note and Drawing State exports from SAP into this spreadsheet and update the index file.
    
    On Error GoTo ErrorHandler
    
' Remove any filters

    Rows("7:7").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
' Select all entries and delete

    Range("A8").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Range("A8").Select
'    ChDir "C:\Documents and Settings\kevin\My Documents\Work\DrgState"
    ChDir "\\atle.bombardier.com\data\UK\PL\DOS\DrgState"
    
' Open Drawings spreadsheet and paste into this sheet
    
    Workbooks.Open FileName:=DrawingFile
    Range("A8").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Windows(GlobalFinderProgramFile).Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
' Move to end of data

    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
' Open Design Notes spreadsheet and paste into this sheet

    Workbooks.Open FileName:=DesignNoteFile
    Range("A8").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Windows(GlobalFinderProgramFile).Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False

' Close Files

    Windows("DesignNoteStateSAP.xlsx").Activate
    ActiveWorkbook.Close False
    Windows("DrgstateSAP.xlsx").Activate
    ActiveWorkbook.Close False
    Range("B8").Select

' Reset CTRL+End
    Reset_Range

' Delete all contents and formatting (by deleting rows) so that CTRL+End works correctly.

    Range("B8").Select
    Selection.End(xlDown).Offset(1, 0).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Delete Shift:=xlUp
    x = ActiveSheet.UsedRange.Rows.Count
    ActiveCell.SpecialCells(xlLastCell).Select

Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: ImportSAP Error number=" & CStr(Err))
    End
End Sub
Sub CreateIndexes()
' Generate new index files
    On Error GoTo ErrorHandler
    
    SetGlobals
    
    Call CreateIndexFile(GlobalCurrentIndexFile, GlobalCurrentIssueFolder, True)
    Call CreateIndexFile(GlobalOldIndexFile, GlobalOldIssueFolder, True)
    
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: CreateIndexes Error number=" & CStr(Err))
    End
End Sub
Sub CreateIndexFile(Index As String, SourcePath As String, UseDosDir As Boolean)
' Creates a text file <Index> containing all the file paths in SourcePath
' The UseDosDir uses the DOS DIR command to generate the index files rather than do it in VB.
' DOS DIR can be faster.

    Dim FileNum As Integer
    
    On Error GoTo ErrorHandler
    
    If UseDosDir Then
        Call WriteIndexUsingDos(Index, SourcePath)
    Else
        FileNum = FreeFile
    
        Open Index For Output As FileNum
        Call WriteIndexFile(FileNum, SourcePath)
        Close FileNum
    End If
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: CreateIndexFile Error number=" & CStr(Err))
    End
End Sub
Sub CreateResultFile(Item As String, IndexFile As String)

    Const Hide As Boolean = True    ' Set to true for normal operation, set to false to allow cmd windows to be seen.
    Dim TaskId As Long
    
    On Error GoTo ErrorHandler

    Set objShell = CreateObject("WScript.Shell")
    
    If Hide Then
        Cmd = Environ$("comspec") & " /c find /i """ & Item & """ " & IndexFile & " > " & GlobalResultFile
        TaskId = objShell.Run(Cmd, 0, True)
    Else
        Cmd = Environ$("comspec") & " /k find /i """ & Item & """ " & IndexFile & " > " & GlobalResultFile
        'Cmd = Environ$("comspec") & " /k find /i """ & Item & """ " & IndexFile
        TaskId = objShell.Run(Cmd, 1, True)
    End If

Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: CreateResultFile Error number=" & CStr(Err))
    End
End Sub
Sub WriteIndexUsingDos(Index As String, SourcePath As String)

    Const Hide As Boolean = True    ' Set to true for normal operation, set to false to allow cmd windows to be seen.
    Dim TaskId As Long
    Dim fso As Object
    Dim stat As Long

    On Error GoTo ErrorHandler
    
    Set objShell = CreateObject("WScript.Shell")
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    ' Create temporary index before overwriting old index if size is not zero
    ' Update current_iss index
    If Hide Then
        Cmd = Environ$("comspec") & " /c " & "dir """ & SourcePath & """ /s/b > " & Index & ".tmp"
        TaskId = objShell.Run(Cmd, 0, True)
    Else
        Cmd = Environ$("comspec") & " /k " & "dir """ & SourcePath & """ /s/b > " & Index & ".tmp"
        TaskId = objShell.Run(Cmd, 1, True)
    End If
    
    ' Check that tmp file is not zero size, overwrite old index & delete tmp file
    If FileLen(Index & ".tmp") > 0 Then
        stat = fso.CopyFile(Index & ".tmp", Index, True)
        fso.DeleteFile (Index & ".tmp")
    Else
        Call LogInformation("Error: WriteIndexUsingDOS: Index has zero length: " & Index)
    End If

Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: WriteIndexUsingDOS Error number=" & CStr(Err))
    End
End Sub
Sub WriteIndexFile(FileNum As Integer, SourcePath As String)
    Set MyObject = New Scripting.FileSystemObject   ' Needs Microsoft Scripting Runtime from Tools - References menu
    Set MySource = MyObject.GetFolder(SourcePath)
    
    Dim PathText As String
    
    On Error GoTo ErrorHandler

    For Each MyFile In MySource.Files
        PathText = MyFile.Path
        Print #FileNum, PathText
    Next

    For Each MySubFolder In MySource.SubFolders
        Call WriteIndexFile(FileNum, MySubFolder.Path)
    Next
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: WriteIndexFile Error number=" & CStr(Err))
    Resume Next
End Sub
Sub Reset_Range()

' Delete all contents and formatting (by deleting rows) so that CTRL+End works correctly.

    On Error GoTo ErrorHandler
    
    Range("B8").Select
    Selection.End(xlDown).Offset(1, 0).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Delete Shift:=xlUp
    x = ActiveSheet.UsedRange.Rows.Count
    ActiveCell.SpecialCells(xlLastCell).Select
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: ResetRange Error number=" & CStr(Err))
    End
End Sub
Public Function GetFilename(Data As String, Optional Delimiter As String = "\") As String
' Returns the filename only from a whole path to the file

  GetFilename = StrReverse(Left(StrReverse(Data), InStr(1, StrReverse(Data), Delimiter) - 1))
   
End Function
Public Function GetPath(Data As String, Optional Delimiter As String = "\") As String
' Returns the path only from a full path to a file

  GetPath = Left(Data, Len(Data) - Len(GetFilename(Data)) - 1)
    
End Function
Public Function DirExists(OrigFile As String)
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    DirExists = fs.FolderExists(OrigFile)
End Function
Function FileExists(ByVal AFileName As String) As Boolean
    On Error GoTo Catch

    FileSystem.FileLen AFileName

    FileExists = True

    GoTo Finally

Catch:
        FileExists = False
Finally:
End Function
Public Function DriveExists(OrigFile As String)
Dim fs, d
Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.DriveExists(OrigFile) = True Then
    Set d = fs.getdrive(OrigFile)
    DExists = 1
        If d.isready = True Then
        DExists = 2
        Exit Function
        End If
    Else
    DExists = 0
    End If
End Function
Sub Tutorial()
    ' Create link to tutorial
    
    On Error GoTo ErrorHandler
    
    link = "file:///" & GlobalTutorialFile
    If FileExists(GlobalTutorialFile) Then
        ' Open tutorial
        ActiveWorkbook.FollowHyperlink link
    Else
        MsgBox "Tutorial File Missing."
    End If
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: Tutorial Error number=" & CStr(Err))
    End
End Sub
Sub RestoreToolbars()
    ' Restore menus
    Application.ScreenUpdating = False
    On Error GoTo 0
'The following line stops copy and paste from working.
'   ActiveWindow.DisplayHeadings = True
    With Application
        .DisplayFullScreen = False
    End With
End Sub
Sub RemoveToolbars()
' Hide the Excel menus
    Application.ScreenUpdating = False
'The following line stops copy and paste from working.
'   ActiveWindow.DisplayHeadings = False
    With Application
        .DisplayFullScreen = True
        .CommandBars("Full Screen").Visible = False
    End With
    On Error GoTo 0
End Sub
Public Function MsgBoxDelay(cMessage, cTitle As String, Timeout As Integer) As Long

    MsgBoxDelay = MessageBoxTimeout(FindWindow(vbNullString, Title), cMessage, cTitle, vbOK, 0, Timeout * 1000)
    
End Function
Sub LogInformation(LogMessage As String)
' Write to log file

Dim FileNum As Integer

    On Error GoTo ErrorHandler
    
    LogMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & " Build: " & Build & " - " & UserNameWindows & " --- " & LogMessage & " ---"
    FileNum = FreeFile ' next file number
    Open GlobalLogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
Exit Sub

ErrorHandler:
    TEMP = MsgBox(LogMessage, vbOKOnly, "Error")
End Sub
Public Function IsFilewriteable(ByVal filePath As String) As Boolean
' Determine whether filePath is writeable.

    Const TestFile As String = "\test.txt"

    On Error GoTo ErrorHandler
        
    Dim nFileNum As Integer
    Dim TestFilePath As String
    
    TestFilePath = filePath & TestFile
    
    nFileNum = FreeFile
    
    Open filePath & TestFile For Output As nFileNum
    Print #nFileNum, " "
    Close nFileNum
    
    ' Delete test file
    If FileExists(TestFilePath) Then SetAttr TestFilePath, vbNormal
    Kill TestFilePath
    
    IsFilewriteable = (Err.Number = 0)
Exit Function

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: IsFileWriteable Error number=" & CStr(Err))
    Resume Next
    Err.Clear
End Function
Sub MakeDirectory(NewDir As String)
    
    On Error GoTo ErrorHandler
    
    'Remove any /
    NewDir = Replace(NewDir, "/", "-")
    
    If DebugMode Then Debug.Print "MkDir:", CurDir, NewDir
    If Not DirExists(NewDir) Then MkDir NewDir
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: MakeDirectory Error number=" & CStr(Err))
    End
End Sub
Sub MakeFile(Item As String, NewFile As String, WhatItIs As String)
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Remove any /
    NewFile = Replace(NewFile, "/", "-")
    Item = Replace(Item, "/", "-")
    
    If DebugMode Then Debug.Print "MkFile:", NewFile
    If Not FileExists(NewFile) Then
        Set oFile = fso.CreateTextFile(NewFile & ".bat")
        oFile.WriteLine GlobalFileOpener & " " & Item & " " & WhatItIs
        oFile.Close
    End If
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: MakeFile Error number=" & CStr(Err))
    End
End Sub
Function UserNameWindows() As String
     UserNameWindows = Environ("USERNAME")
End Function
Sub CloseSheet()
    SetGlobals

    RestoreToolbars
    Workbooks(GlobalFinderProgramFile).Close
End Sub
Function ReadIndex(Index As String) As String()
' Read whole index text file into array for quicker searching
    Dim MyData As String
    
    On Error GoTo ErrorHandler
    
     '~~> Open the file in 1 go to read it into an array
    Open Index For Binary As #1
    MyData = Space$(LOF(1))
    Get #1, , MyData
    Close #1
    
    ReadIndex = Split(MyData, vbCrLf)
Exit Function

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: ReadIndex Error number=" & CStr(Err))
    End
End Function
Sub CheckForPaths(Highlight As Boolean, RecordPath As Boolean)
' Check the drawing number on each row and write its path into column K

    Const Red = 3
    Const DrawingCol = 1    ' Col A
    Const ECRCol = 6        ' Col F
    Const StartRow = 8      ' Fist row after titles rows
    Const PathCol = 11      ' Write results to columns k on
    
    Dim Drawing As String, ECR As String
    Dim row As Range
    Dim cell As Range
    Dim Results() As String
    
    On Error GoTo ErrorHandler
    
    NumRows = Range("A1", Range("A8").End(xlDown)).Rows.Count
    
    ' Clear existing highlights
    Range("A8", Range("F8").End(xlDown)).Select
    Selection.Interior.ColorIndex = 0
    Range("A8").Select
    
    ' Establish "For" loop to loop "numrows" number of times.
    For i = StartRow To NumRows
        Drawing = Cells(i, DrawingCol).Value
        ECR = Cells(i, ECRCol).Value
        ' Find and replace '/' with '-' for file name.
        Drawing = Replace(Drawing, "/", "-")
        ' Format ECR number to match filenames
        ECR = Replace(ECR, "600000000000", "6-")
        ECR = Replace(ECR, "60000000000", "6-")
        ECR = Replace(ECR, "6000000000", "6-")
        ECR = Replace(ECR, "600000000", "6-")
        ECR = Replace(ECR, "60000000", "6-")
        ECR = Replace(ECR, "6000000", "6-")
        
        ' Add the \ to make for a more accurate match
        Drawing = "\" & Drawing
        ECR = "\" & ECR
        
        ' Deal with Drawings, use vbTextCompare to be case insensitive
        If Drawing <> "\" Then
            Results = Filter(CurrentIndexArray, Drawing, True, vbTextCompare)
            If UBound(Results) >= 0 Then
                If RecordPath Then
                    For j = LBound(Results) To UBound(Results)
                        Cells(i, PathCol + j) = Results(j)
                    Next j
                End If
            Else:
                Results = Filter(OldIndexArray, Drawing, True, vbTextCompare)
                If UBound(Results) >= 0 Then
                    If RecordPath Then
                        
                        For j = LBound(Results) To UBound(Results)
                            Cells(i, PathCol + j) = Results(j)
                        Next j
                    End If
                Else
                    ' Highlight the drawing number in red
                    If Highlight Then Cells(i, DrawingCol).Interior.ColorIndex = Red
                End If
            End If
        End If
        ' Deal with ECRs
        If ECR <> "\" And ECR <> "" Then
            Results = Filter(CurrentIndexArray, ECR)
            If UBound(Results) >= 0 Then
                If RecordPath Then
                    ' Write results to columns k on.
                    For j = LBound(Results) To UBound(Results)
                        Cells(i, 11 + j) = Results(j)
                    Next j
                End If
            Else
                ' Highlight the drawing number in red
                If Highlight Then Cells(i, ECRCol).Interior.ColorIndex = Red
            End If
        End If
    Next i
    
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: CheckForPaths Error number=" & CStr(Err))
    End
End Sub
Sub CheckForArchivedFiles()
' Add the indexed path for the drawing to each row.

    Const Highlight As Boolean = True, RecordPath As Boolean = False
    
    On Error GoTo ErrorHandler
    
    SetGlobals
    Call LogInformation("ArchivedFiles: Start Search: Highlight=" & CStr(Highlight) & " Path=" & CStr(RecordPath))
    CurrentIndexArray = ReadIndex(GlobalCurrentIndexFile)
    OldIndexArray = ReadIndex(GlobalOldIndexFile)
    
    Call CheckForPaths(Highlight, RecordPath)
    Call LogInformation("ArchivedFiles: Complete")
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: CheckForArchivedFiles Error number=" & CStr(Err))
    End
End Sub
Public Sub KillDirs(ByVal strFolderPath As String)
'Recursively delete files and folders

   Dim fsoSubFolders As Folders
   Dim fsoFolder As Folder
   Dim fsoSubFolder As Folder
   
   Dim strPaths()
   Dim lngFolder As Long
   Dim lngSubFolder As Long
   
   On Error GoTo ErrorHandler
      
   DoEvents
   
   Set m_fsoObject = New FileSystemObject
   If Not m_fsoObject.FolderExists(strFolderPath) Then Exit Sub
   
   Set fsoFolder = m_fsoObject.GetFolder(strFolderPath)
   
   On Error Resume Next
   
   'Has sub-folders
   If fsoFolder.SubFolders.Count > 0 Then
        lngFolder = 1
        ReDim strPaths(1 To fsoFolder.SubFolders.Count)
        'Get each sub-folders path and add to an array
        For Each fsoSubFolder In fsoFolder.SubFolders
            strPaths(lngFolder) = fsoSubFolder.Path
            lngFolder = lngFolder + 1
        Next fsoSubFolder
        
        lngSubFolder = 1
        'Recursively call the function for each sub-folder
        Do While lngSubFolder < lngFolder
           Call KillDirs(strPaths(lngSubFolder))
           lngSubFolder = lngSubFolder + 1
        Loop
    End If
   
    'Delete files
    If fsoFolder.Files.Count > 0 Then
        Kill strFolderPath & "\*.*"
    End If
   
    'No sub-folders or files
    If fsoFolder.Files.Count = 0 And fsoFolder.SubFolders.Count = 0 Then
        fsoFolder.Delete
    End If
Exit Sub

ErrorHandler:
    Call MsgBoxDelay("Sorry something appears to have gone wrong...", "Error", ShowDurationSecs)
    Call LogInformation("Error: KillDirs Error number=" & CStr(Err))
    End
End Sub


