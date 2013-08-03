Attribute VB_Name = "DrawingTree"
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
Private Sub Main()
' Get all the drawings/materials from the the current open drawing.
' Use this list to form a linked list of all the sub level BOMs.
' Look for sub level BOMs and open if Word or Excel

    Dim DrawingList(MaxArray) As String
    Dim Item As String
    Dim Index As Integer
    Dim IndexFile As String
    
    ' Get a list of all the drawings/materials
    GetAllDrawings Refs:=DrawingList
    
    SetGlobals
    
    For Index = 1 To UBound(DrawingList)
        Item = DrawingList(Index)
        If IsBOM(Item) Then
            If DebugMode Then Debug.Print Index, Item
            IndexFile = GlobalCurrentIndexFile
            Call CreateResultFile(Item, IndexFile)
            IndexFile = GlobalResultFile
            NewDoc = MsOfficeDoc(IndexFile)
        End If
    Next Index
    Stop
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
    
    ' Assign Log file path
    ' Select local log file if user doesn't have write access to network log file
    If Not IsFilewriteable(GlobalProgramPath) Then
        GlobalLogFile = LocalLogPath & "DrawingFinderLogFile.txt"
    Else
        GlobalLogFile = GlobalProgramPath & "DrawingFinderLogFile.txt"
    End If
    
End Sub
Private Sub GetAllDrawings(ByRef Refs() As String)
' Compile an array of all the drawing/material numbers

    Dim DrawingRowStart As Integer
    Dim DrawingColStart As Integer
    Dim ListIndex As Integer
    Dim ActiveRow As Integer
    

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

End Sub
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
Function IsBOM(Item As String) As Boolean
' Determine whether Item is a BOM. Look for new parts lists L52xxxxxxx or old SXL & GXL numbers.

'    If (Len(Item) = 6 And Left(Item, 1) = "1") Or (Len(Item) = 9 And Left(Item, 2) = "52") Or Not (Item Like "*#*") Then
    IsBOM = (Left(Item, 3) = "L52") Or (Item Like "*SXL*") Or (Item Like "*GXL*")
    If IsBOM And DebugMode Then Debug.Print "IsBOM", Item, IsBOM
    
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



