Attribute VB_Name = "DrawingLink"
'Option Explicit

' Define globals
Public GlobalNum As String
Public GlobalDesc As String
Public GlobalFinderProgramFile As String
Public GlobalCurrentIssueFolder As String
Public GlobalOldIssueFolder As String
Public GlobalTransferFolder As String
Public GlobalCurrentIndexFile As String
Public GlobalTransferIndexFile As String
Public GlobalOldIndexFile As String
Public GlobalResultFile As String

Public GlobalPartFolder As String
Public GlobalPartTransferFolder As String
Public GlobalPartIndexFile As String
Public GlobalPartTransferIndexFile As String

Public GlobalBatchFile As String
Public GlobalLogFile As String
Public Globalfilepath As String
Public Globaldrive As String
Public GlobalTutorialFile As String

Const Build As String = 4

Const ShowDurationSecs As Integer = 5

Const NetProgramPath = "\\atle.bombardier.com\data\uk\pl\dos\Drgstate"
Const NetTransferPath = "\\atle.bombardier.com\data\uk\pl\dos"
Const NetDataPath = "\\atle.bombardier.com\data\uk\pl\dos2"
Const LocalProgramPath = "\Drgstate"
Const LocalTransferPath = ""
Const LocalLogPath = "c:\windows\temp"
Const LocalDataPath = "\"
Const DesignNoteFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DesignNoteStateSAP.xlsx"
Const DrawingFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DrgStateSAP.xlsx"

Enum RequestType
    Drawing = 1
    Part = 2
End Enum

Enum ActionType
    OpenInApp = 1
    ShowInFolder = 2
End Enum

Type Article
    name As String
    Drawing As Boolean
End Type

Dim fso As New FileSystemObject   ' Needs Microsoft Scripting Runtime from Tools - References menu
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
    Dim ProgPath As String
    Dim TransferPath As String
    Dim Drive As String

    ' Find out whether network drive is connected
    If DirExists(NetDataPath) Then
        DataPath = NetDataPath
        ProgramPath = NetProgramPath
        TransferPath = NetTransferPath
    Else
        Drive = Switch(DirExists("e:\1_current_iss"), "e:", DirExists("f:\1_current_iss"), "f:", DirExists("g:\1_current_iss"), "g:", DirExists("c:\1_current_iss"), "c:", True, "Not Found")
        DataPath = Drive
        ProgramPath = Drive & LocalProgramPath
        TransferPath = Drive & LocalTransferPath
    End If
    
    ' Test for no suitable directory found
    If Drive = "Not Found" Then
        MsgBox ("Current Issue" & vbLf & "Folder not found")
        End
    End If


    ' Assign Prorgam related variables
    GlobalFinderProgramFile = "DrawingFinder.xls"
    GlobalCurrentIndexFile = ProgramPath & "\CurrentIndex.txt"
    GlobalOldIndexFile = ProgramPath & "\OldIndex.txt"
    GlobalBatchFile = ProgramPath & "\CreateIndex.bat"
    GlobalTutorialFile = ProgramPath & "\DrawingFinderTutorial.pdf"
    GlobalPartIndexFile = ProgramPath & "\PartsCurrentIndex.txt"
    
    ' Assign Data related variables
    GlobalCurrentIssueFolder = DataPath & "\1_current_iss"
    GlobalOldIssueFolder = DataPath & "\1_old_iss"
    GlobalPartFolder = DataPath & "\Parts PDF Datasheets"
        
    GlobalTransferFolder = TransferPath & "\1_files for filing"
    GlobalPartTransferFolder = TransferPath & "\1_Parts PDFs for filing"
    GlobalResultFile = LocalLogPath & "\DrawingFinderResult.txt"
    GlobalTransferIndexFile = LocalLogPath & "\DrawingFinderTransferIndex.txt"
    GlobalPartTransferIndexFile = LocalLogPath & "\PartTransferIndex.txt"
    
    ' Assign Log file path
    ' Select local log file if user doesn't have write access to network log file
    If Not IsFilewriteable(ProgramPath) Then
        GlobalLogFile = LocalLogPath & "\DrawingLinkLogFile.txt"
    Else
        GlobalLogFile = ProgramPath & "\DrawingLinkLogFile.txt"
    End If

End Sub
Function IsDrawing() As Article
' Detect whether selection is a drawing (includes ECRs)

Dim name As String

    ' Determine whether anything has already been selected
    name = Selection
    ' Strip out any line feed type chars & replace / with -
    name = Replace(name, ".", "")
    name = Replace(name, vbLf, "")
    name = Replace(name, vbCr, "")
    name = Replace(name, Chr(11), "")
    name = Replace(name, "/", "-")
    If Len(name) = 1 Then
        ' Determine whether number is part or drawing
        Selection.MoveLeft Unit:=wdWord, Count:=1
        Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
        ' Strip off trailing space
        If Selection.Characters.Last = " " Then Selection.End = Selection.End - 1
        name = Selection
    End If
    
    ' Format ECR number to match filenames
    name = Replace(name, "600000000000", "6-")
    name = Replace(name, "60000000000", "6-")
    name = Replace(name, "6000000000", "6-")
    name = Replace(name, "600000000", "6-")
    name = Replace(name, "60000000", "6-")
    name = Replace(name, "6000000", "6-")
    
    ' Convert name to number to check for part number
    nameval = Val(name)
    ' Check that the name is definitely a part number
    If name = LTrim(Str(nameval)) And (nameval > 100000 And nameval < 127000) Or (nameval > 520000000 And nameval < 530000000) Then
        IsDrawing.name = name
        IsDrawing.Drawing = False
    Else
        IsDrawing.name = name
        IsDrawing.Drawing = True
    End If

End Function
Sub ChooseAction()
' Display show menu and allow choice of open or show in folder
    Dim Request As RequestType
    Dim Action As ActionType
    Dim Ch As String
    Dim Index As String
    Dim TransferIndex As String
    Dim Choice As Integer
    
    SetGlobals
        ' Create strings for log entry
        
    If IsDrawing.Drawing Then
        ChoiceStr = "Drawing"
        Index = GlobalCurrentIndexFile
        TransferIndex = GlobalTransferIndexFile
        Request = Drawing
    Else
        ChoiceStr = "Part"
        Index = GlobalPartIndexFile
        TransferIndex = GlobalPartTransferIndexFile
        Request = Part
    End If
    
    Choice = -9
    
    Do Until (Choice > 0) And (Choice <= 2)
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
    
   Call LogInformation("Show: Choice: " & ChoiceStr & " Action: " & ActionStr)
    
    ' Create new Transfer Index for latest files
    If IsDrawing.Drawing Then
        Call CreateIndexFile(GlobalTransferIndexFile, GlobalTransferFolder, True)
    Else
        Call CreateIndexFile(GlobalPartTransferIndexFile, GlobalPartTransferFolder, True)
    End If

    ' Carry out appropriate action
        If Not ShowItem(Request, Action, Index) Then
            If Request = Drawing Then TransferIndex = GlobalTransferIndexFile Else TransferIndex = GlobalPartTransferIndexFile
            If Not ShowItem(Request, Action, TransferIndex) Then
                ' no paths returned from search
                MsgBox ("File not found")
                Call LogInformation("ShowItem: File not found")
            End If
        End If
End Sub
Function ShowItem(Request As RequestType, Action As ActionType, IndexFile As String) As Boolean
' Search for and then show or open what has been requested

    Const MaxResults As Integer = 10
    Const UseResult As Boolean = True
    
    Dim Cmd As String
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
    
        ShowItem = False
        MinLines = 0
    
    ' Get drawing name
    Item = IsDrawing.name
    
    'Generate full file name for old issue
    ' Create strings for log entry
    Select Case Request
        Case Drawing
            RequestStr = "Drawing"
        Case Part
            RequestStr = "Part"
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

        If Action = OpenInApp Then
            Call LogInformation("ShowItem Open in App: File path: " & ResultFilePath)
            ' Create full path to file
            ResultFilePath = "file:///" & ResultFilePath
            ' Open file in applicaion
            'varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & ResultFilePath, WIN_NORMAL)
            On Error Resume Next
               ActiveDocument.FollowHyperlink ResultFilePath   ' Use for Word
               ActiveWorkbook.FollowHyperlink ResultFilePath   ' Use for Excel
            On Error GoTo 0 ' Disable previous On Error
        Else
            Call LogInformation("ShowItem Open in Folder: File path: " & ResultFilePath)
            Shell "explorer /e, /select," & ResultFilePath, vbNormalFocus    ' Use /e to show the folders pane. Opening showing the folders over the network is very slow!
            'Shell "explorer /select," & ResultFilePath, vbNormalFocus
        End If
    End If
End Function
Sub CreateIndexFile(Index As String, SourcePath As String, UseDosDir As Boolean)
' Creates a text file <Index> containing all the file paths in SourcePath
' The UseDosDir uses the DOS DIR command to generate the index files rather than do it in VB.
' DOS DIR can be faster.

    Dim FileNum As Integer
    
    If UseDosDir Then
        Call WriteIndexUsingDos(Index, SourcePath)
    Else
        FileNum = FreeFile
    
        Open Index For Output As FileNum
        Call WriteIndexFile(FileNum, SourcePath)
        Close FileNum
    End If
End Sub
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
Sub WriteIndexUsingDos(Index As String, SourcePath As String)

    Const Hide As Boolean = True    ' Set to true for normal operation, set to false to allow cmd windows to be seen.
    Dim TaskId As Long
    
    Set objShell = CreateObject("WScript.Shell")
    
    ' Update current_iss index
    If Hide Then
        Cmd = Environ$("comspec") & " /c " & "dir """ & SourcePath & """ /s/b > " & Index
        TaskId = objShell.Run(Cmd, 0, True)
    Else
        Cmd = Environ$("comspec") & " /k " & "dir """ & SourcePath & """ /s/b > " & Index
        TaskId = objShell.Run(Cmd, 1, True)
    End If

End Sub
Sub WriteIndexFile(FileNum As Integer, SourcePath As String)
    Set MyObject = New Scripting.FileSystemObject   ' Needs Microsoft Scripting Runtime from Tools - References menu
    Set MySource = MyObject.GetFolder(SourcePath)
    
    Dim PathText As String
    
    On Error Resume Next

    For Each MyFile In MySource.Files
        PathText = MyFile.Path
        Print #FileNum, PathText
    Next

    For Each MySubFolder In MySource.SubFolders
        Call WriteIndexFile(FileNum, MySubFolder.Path)
    Next
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
    DirExists = fs.folderexists(OrigFile)
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
    
    link = "file:///" & GlobalTutorialFile
    If FileExists(GlobalTutorialFile) Then
        ' Open tutorial
        ActiveWorkbook.FollowHyperlink link
    Else
        MsgBox "Tutorial File Missing."
    End If
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

    LogMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & " Build: " & Build & " - " & UserNameWindows & " --- " & LogMessage & " ---"
    FileNum = FreeFile ' next file number
    Open GlobalLogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub
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
Function UserNameWindows() As String
     UserNameWindows = Environ("USERNAME")
End Function
Sub CloseSheet()
' Close sheet to allow for scheduled update
    Dim wBook As Workbook
    
    On Error Resume Next    ' Needed to allow workbook to be set even if workbook is not open
    
    Call LogInformation("CloseSheet: Application Closed")
 
    Set wBook = Workbooks(GlobalFinderProgramFile)
    If Not (wBook Is Nothing) Then
'        RestoreToolbars
'        Application.EnableEvents = False    ' For some reason the close events prevent the workbook from closing
        DoEvents
        Workbooks(GlobalFinderProgramFile).Close
    End If
End Sub

