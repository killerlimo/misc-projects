Attribute VB_Name = "DrawingTree"
' Must select Tools-Microsoft Runtime & Microsoft Excel Objects

Const MaxArray = 1000
Const DebugMode = True
Const Build As String = 1

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

Enum WhatIsIt
    BOM
    DRG
    Mat
End Enum

Enum AppType
    Word
    Excel
End Enum

Public Type DrawingType
    Number As String
    Is As WhatIsIt
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
Private Sub Main()
' Get all the drawings/materials from the the current open drawing.
' Use this list to form a linked list of all the sub level BOMs.
' Look for sub level BOMs and open if Word or Excel
' Create a folder & file structure to represent the data

    Dim WhatItIs As String
    Dim Index As Integer
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim SubFolder As Folder
    
    If DebugMode Then
        Debug.Print
        Debug.Print "---Start---"
    End If
    
   
    SetGlobals

    MakeDirectory (GlobalTreeRoot)
    ChDir GlobalTreeRoot
    
    ' Get top level BOM
'    TopLevelBOM = InputBox("Enter top level BOM:", "Drawing Number")
    TopLevelBOM = "L520002408"
    MakeDirectory (TopLevelBOM)
    
    Set FSfolder = FS.GetFolder(GlobalTreeRoot)
    For Each SubFolder In FSfolder.SubFolders
        Call BuildTree(SubFolder)
    Next SubFolder
    
End Sub
Sub BuildTree(SubLevelBOM As Folder)

    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim SubFolder As Folder
    Dim CurrentBOMDoc As String
    Dim Item As String
    Dim IndexFile As String
    Dim DrawingList() As DrawingType
    Dim MaxDrawings As Integer
    Dim WhatApp As AppType
    
    ' Strip BOM name from path
    Item = FS.GetFilename(SubLevelBOM)
    ChDir SubLevelBOM
    
    ' Find the BOM, open it and extract the drawings/materials.
    IndexFile = GlobalCurrentIndexFile
    Call CreateResultFile(Item, IndexFile)
    IndexFile = GlobalResultFile
    CurrentBOMDoc = MsOfficeDoc(IndexFile)
    
    ' Detect Word/Excel and open document
    If InStr(UCase(CurrentBOMDoc), "XLS") Then
        If DebugMode Then Debug.Print "Opening ExcelDoc", CurrentBOMDoc
        Set App = CreateObject("Excel.Application")
        App.Workbooks.Open CurrentBOMDoc
        App.Visible = True
'        Workbooks.Open(CurrentBOMDoc).Activate
        WhatApp = Excel
    Else
        If DebugMode Then Debug.Print "Opening WordDoc", CurrentBOMDoc
        Set WordApp = CreateObject("word.Application")
        WordApp.Documents.Open CurrentBOMDoc
        WordApp.Visible = True
        Documents.Open(CurrentBOMDoc).Activate
        WhatApp = Word
    End If
    
    ' Get a list of all the drawings/materials
    Call GetAllDrawings(WhatApp, Refs:=DrawingList, Occupied:=MaxDrawings)
    ReDim Preserve DrawingList(MaxDrawings) As DrawingType
    Call QuickSort(DrawingList, LBound(DrawingList), UBound(DrawingList))
    
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
        
        Item = Replace(Item, "/", "-")
        
        If DrawingList(Index).Is = BOM Then
            MakeDirectory (Item)
            IndexFile = GlobalCurrentIndexFile
            Call CreateResultFile(Item, IndexFile)
            IndexFile = GlobalResultFile
            NewDoc = MsOfficeDoc(IndexFile)
            If DebugMode Then Debug.Print "Main", Item, NewDoc
        Else
            ' Create file
            MakeFile (Item & "." & WhatItIs)
        End If
    Next Index
    
    ' Detect Word/Excel and close document
    If InStr(UCase(CurrentBOMDoc), "XLS") Then
        If DebugMode Then Debug.Print "Closing ExcelDoc", CurrentBOMDoc
    Else
        If DebugMode Then Debug.Print "Closing WordDoc", CurrentBOMDoc
        'WordApp.Documents.Close
        WordApp.Quit wdDoNotSaveChanges
        Set WordApp = Nothing
    End If
    
    ' Recursively build the tree
    Set FSfolder = FS.GetFolder(SubLevelBOM)
    For Each SubFolder In FSfolder.SubFolders
        Call BuildTree(SubFolder)
    Next SubFolder
    
End Sub
Public Sub SetGlobals()
' Global variables for use throughout the program
    
    Dim DataPath As String
    Dim TransferPath As String
    Dim Drive As String

    ' Find out whether network drive is connected
    If DirExists(NetDataPath) Then
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
    
    ' Assign Data related variables
    GlobalCurrentIssueFolder = DataPath & "1_current_iss"
    GlobalOldIssueFolder = DataPath & "1_old_iss"
    
    GlobalTransferFolder = TransferPath & "1_files for filing"
    GlobalResultFile = LocalLogPath & "DrawingTreeResult.txt"
    GlobalTransferIndexFile = LocalLogPath & "DrawingTreeTransferIndex.txt"
    GlobalTreeRoot = LocalLogPath & TreeRoot
    
    ' Assign Log file path
    ' Select local log file if user doesn't have write access to network log file
    If Not IsFilewriteable(GlobalProgramPath) Then
        GlobalLogFile = LocalLogPath & "DrawingFinderLogFile.txt"
    Else
        GlobalLogFile = GlobalProgramPath & "DrawingFinderLogFile.txt"
    End If

End Sub
Public Sub GetAllDrawings(WhatApp As AppType, ByRef Refs() As DrawingType, ByRef Occupied As Integer)
' Compile an array of all the drawing/material numbers

    Dim aTable As Table
    Dim aCell As Cell
    Dim aRow As Integer
    Dim DrawingRowStart As Integer
    Dim DrawingColStart As Integer
    Dim ListIndex As Integer
    Dim ActiveRow As Integer

    If WhatApp = Excel Then
        DrawingRowStart = Range(StartOfDrawings).Row
        DrawingColStart = Range(StartOfDrawings).Column

        MaxRows = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        ListIndex = 1
        For ActiveRow = DrawingStartRow To MaxRows
            Refs(ListIndex) = Cells(ActiveRow + DrawingRowStart + 1, DrawingColStart)
            If Refs(ListIndex) <> "" Then
                Refs(ListIndex) = OnlyAlphaNumericChars(Refs(ListIndex))
                ListIndex = ListIndex + 1
            End If
        Next ActiveRow
    Else
    
        For Each aTable In ActiveDocument.Tables
            MaxRows = aTable.Range.Rows.Count
            
            ReDim Refs(MaxRows)
            
            Occupied = 0
            For aRow = 1 To MaxRows - 1
                Set aCell = aTable.Cell(aRow + 1, 2)
                Refs(aRow).Number = OnlyAlphaNumericChars(aCell.Range)
                If Refs(aRow).Number <> "" Then
                    Occupied = Occupied + 1
                    Refs(aRow).Is = IsDrawingType(Refs(aRow).Number)
                End If
            Next aRow
        Next aTable
    End If
End Sub
Function StartOfDrawings() As String
' Find start of drawing/material list

    Dim SearchString As String
    Dim SearchRange As Range, cl As Range
    Dim FirstFound As String
    Dim sh As Worksheet

    ' Set Search value
    SearchString = "SAP"
    Application.FindFormat.Clear
    ' loop through all sheets
    For Each sh In ActiveWorkbook.Worksheets
        ' Find first instance on sheet
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
                cl.Font.Bold = True
                cl.Interior.ColorIndex = 3
                ' find next instance
                Set cl = sh.Cells.FindNext(After:=cl)
                ' repeat until back where we started
            Loop Until FirstFound = cl.Address
        End If
    Next
    StartOfDrawings = FirstFound
End Function
Private Function OnlyAlphaNumericChars(OrigString As String) As String
' Remove unwanted characters

    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    
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

End Function
Private Function IsAlphaNumeric(sChr As String) As Boolean
' Check that charcter is in acceptable list

    IsAlphaNumeric = sChr Like "[0-9A-Za-z,-,/]"
End Function
Function IsDrawingType(Item As String) As WhatIsIt
' Return the type of drawing, BOM, DWG or MAT
' Determine whether Item is a BOM. Look for new parts lists L52xxxxxxx or old SXL & GXL numbers.

    If (Left(Item, 3) = "L52") Or (Item Like "*SXL*") Or (Item Like "*GXL*") Then
        IsDrawingType = BOM
    ElseIf (Len(Item) = 6 And Left(Item, 1) = "1") Or (Len(Item) = 9 And Left(Item, 2) = "52") Then
            IsDrawingType = Mat
    Else
        IsDrawingType = DRG
    End If

    ' If DebugMode Then Debug.Print "WhatIsIt", Item, IsDrawingType
End Function
Sub CreateResultFile(Item As String, IndexFile As String)

    Const Hide As Boolean = True    ' Set to true for normal operation, set to false to allow cmd windows to be seen.
    Dim TaskId As Long

    Set objShell = CreateObject("WScript.Shell")
    
    If Hide Then
        Cmd = Environ$("comspec") & " /c find /i """ & Item & """ " & IndexFile & " > " & GlobalResultFile
        TaskId = objShell.Run(Cmd, 0, True)
    Else
        Cmd = Environ$("comspec") & " /k find /i """ & Item & """ " & IndexFile & " > " & GlobalResultFile
        'Cmd = Environ$("comspec") & " /k find /i """ & Item & """ " & IndexFile
        TaskId = objShell.Run(Cmd, 1, True)
    End If

End Sub
Function MsOfficeDoc(IndexFile) As String
    ' Search for item in index file
    
    Const MaxResults As Integer = 10
    Dim ResultArray(1 To MaxResults) As String
    Dim Index As Integer
    
    FileNum = FreeFile
    MsOfficeDoc = ""
    
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
'                ShowItem = True
        For IntIndex = MinLines + 1 To Line
            StrBuf = StrBuf & IntIndex & ". " & Right(GetFilename(ResultArray(IntIndex + MinLines)), 100) & vbLf
        Next
        Choice = -9
        
        ' Look for an MS Office document
        Index = 0
        Do
            Index = Index + 1
            If UCase(ResultArray(Index + MinLines)) Like "*DOC*" Then MsOfficeDoc = ResultArray(Index + MinLines)
            If UCase(ResultArray(Index + MinLines)) Like "*XLS*" Then MsOfficeDoc = ResultArray(Index + MinLines)
        Loop Until Index = Line Or MsOfficDoc <> ""
        
    End If
End Function
Private Sub QuickSort(ByRef Field() As DrawingType, LB As Long, UB As Long)
    Dim P1 As Long, P2 As Long, Ref As DrawingType, TEMP As DrawingType

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
End Sub
Public Function DirExists(OrigFile As String)
    Dim FS
    Set FS = CreateObject("Scripting.FileSystemObject")
    DirExists = FS.folderexists(OrigFile)
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
Public Function IsFilewriteable(ByVal filePath As String) As Boolean
' Determine whether filePath is writeable.

    Const TestFile As String = "\test.txt"

    On Error Resume Next
    Err.Clear
    
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
End Function
Sub MakeDirectory(NewDir As String)
    If Not DirExists(NewDir) Then MkDir NewDir
End Sub
Sub MakeFile(NewFile As String)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not FileExists(NewFile) Then
        Set oFile = fso.CreateTextFile(NewFile)
        oFile.WriteLine NewFile
        oFile.Close
    End If
End Sub
Public Function GetFilename(Data As String, Optional Delimiter As String = "\") As String
' Returns the filename only from a whole path to the file

  GetFilename = StrReverse(Left(StrReverse(Data), InStr(1, StrReverse(Data), Delimiter) - 1))
   
End Function

