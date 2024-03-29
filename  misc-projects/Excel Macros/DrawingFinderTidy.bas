Attribute VB_Name = "DrawingFinderTidy"
'Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal lnghProcess As Long, lpExitCode As Long) As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

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
Public GlobalBatchFile As String
Public GlobalLogFile As String
Public Globalfilepath As String
Public Globaldrive As String
Public GlobalTutorialFile As String


Const ShowDurationSecs As Integer = 5

Const NetProgramPath = "\\atle.bombardier.com\data\uk\pl\dos\Drgstate"
Const NetTransferPath = "\\atle.bombardier.com\data\uk\pl\dos"
Const NetDataPath = "\\atle.bombardier.com\data\uk\pl\dos2"
Const LocalProgramPath = "\Drgstate"
Const LocalTransferPath = ""
Const LocalLogPath = "\windows\temp"
Const LocalDataPath = "\"
Const DesignNoteFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DesignNoteStateSAP.xlsx"
Const DrawingFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DrgStateSAP.xlsx"

Enum RequestType
    Current = 1
    Old = 2
    ECR = 3
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
    GlobalTransferIndexFile = ProgramPath & "\TransferIndex.txt"
    GlobalBatchFile = ProgramPath & "\CreateIndex.bat"
    GlobalTutorialFile = ProgramPath & "\DrawingFinderTutorial.pdf"
    
    ' Assign Data related variables
    GlobalCurrentIssueFolder = DataPath & "\1_current_iss"
    GlobalOldIssueFolder = DataPath & "\1_old_iss"
    
    GlobalTransferFolder = TransferPath & "\1_files for filing"
    
    ' Assign Log file path
    GlobalLogFile = ProgramPath & "\DrawingFinderLogFile.txt"
    ' Select local log file if user doesn't have write access to network log file
    If Not IsFileWritable(GlobalLogFile) Then GlobalLogFile = LocalLogPath & "\DrawingFinderLogFile.txt"

End Sub
Sub Filter()

' Set up data filters on Item No. and Description columns.
' Enter up to 2 words in each search box, these will be OR'd

    Dim NumWords() As String
    Dim DescWords() As String
    
    Dim NumWordCount, DescWordCount As Integer
    
    NumWordCount = 4
    DescWordCount = 4
    
    Do Until (NumWordCount < 3)
        ' Get search data from user & convert to upper case
        GlobalNum = StrConv(InputBox("Enter part of the Drawing Number" & vbLf & "Up to 2 words can be entered" & vbLf _
        & "Use & | for AND OR", "Drawing Search", GlobalNum), vbUpperCase)
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
    Call LogInformation("TextFilter: Drawing: " & GlobalNum & " Description: " & GlobalDesc)
        
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
End Sub
Sub ChooseAction()
' Display show menu and allow choice of current issue, old issue or show in folder
    Dim Request As RequestType
    Dim Action As ActionType
    Dim Ch, Index As String
    Dim Choice As Integer
    
    SetGlobals
    
    Choice = -9
    
    Do Until (Choice > 0) And (Choice <= 3)
        Ch = InputBox("1. Current Issue" & vbLf & "2. Old Issue" & vbLf & "3. ECR", "Choose option:", 1)
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
    End Select
    
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
    Call CreateIndexFile(GlobalTransferIndexFile, GlobalTransferFolder, False)

    ' Carry out appropriate action
        If Not ShowItem(Request, Action, Index) Then
            If Request = Current Then
                If Not ShowItem(Request, Action, GlobalTransferIndexFile) Then
                    ' no paths returned from search
                    MsgBox ("File not found")
                    Call LogInformation("ShowItem: File not found")
                End If
            Else
                ' no paths returned from search
                MsgBox ("File not found")
                Call LogInformation("ShowItem: File not found")
            End If
        End If
End Sub
Function ShowItem(Request As RequestType, Action As ActionType, IndexFile As String) As Boolean
' Search for and then show or open what has been requested

    Const MaxResults As Integer = 10
    Dim Cmd As String
    Dim Drawing, Issue, Correction, ECRnum As String
    Dim Reply As Variant
    
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
        
    ' Read cell contents
    Drawing = Cells(ActiveCell.Row, 1).Value
    Issue = Cells(ActiveCell.Row, 3).Value
    Correction = Cells(ActiveCell.Row, 4).Value
    ECRnum = Cells(ActiveCell.Row, 6).Value
    
    ' Find and replace '/' with '-' for file name.
    Drawing = Replace(Drawing, "/", "-")
    
    ' Format ECR number to match filenames
    ECRnum = Replace(ECRnum, "600000000000", "6-")
    ECRnum = Replace(ECRnum, "60000000000", "6-")
    ECRnum = Replace(ECRnum, "6000000000", "6-")
    ECRnum = Replace(ECRnum, "600000000", "6-")
    ECRnum = Replace(ECRnum, "60000000", "6-")
    ECRnum = Replace(ECRnum, "6000000", "6-")
       
    'Generate full file name for old issue
    Select Case Request
        Case ECR
            Item = ECRnum
        Case Current
            Item = Drawing
        Case Old
            Item = Drawing & "-" & Issue & Correction
    End Select
    
    ' Check for null string in case the cursor is selecting a non-drawing line
    If Item = "" Then
        MsgBox ("No drawing selected")
        Exit Function
    End If
    
    ' Create strings for log entry
    Select Case IssueRequest
        Case Current
            RequestStr = "Current"
        Case Old
            RequestStr = "Old"
        Case ECR
            RequestStr = "ECR"
    End Select

    ' Write seach item to log
    Call LogInformation("ShowItem: Request: " & RequestStr & " : " & Item)

'    If UseIndex Then
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
'    Else
'
'        lSize = FindFile("c:\1_current_iss", UCase(Item), nDirs, nFiles)
'    End If
    
    ' More than 1 line indicates that at least 1 file has been found
    If Line > 0 Then
                ShowItem = True
        For IntIndex = 1 To Line
            StrBuf = StrBuf & IntIndex & ". " & Right(GetFilename(ResultArray(IntIndex)), 100) & vbLf
        Next
        Choice = -9
        
        ' List results to choose from
        Do Until (Choice > 0) And (Choice <= Line)
            Ch = InputBox(StrBuf, "Choose File:", 1)
        ' Check for Escape key
        If Ch = "" Then Exit Function Else Choice = Int(Ch)
        Loop
        
        ResultFilePath = ResultArray(Choice)

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
End Function
Private Function FindFile(ByVal sFol As String, sFile As String, _
   nDirs As Long, nFiles As Long) As Currency
   Dim tFld As Folder, tFil As file, FileName As String
   
 '  On Error GoTo Catch
   Set fld = fso.GetFolder(sFol)
   FileName = Dir(fso.BuildPath(fld.path, sFile), vbNormal Or _
                  vbHidden Or vbSystem Or vbReadOnly)
   While Len(FileName) <> 0
      FindFile = FindFile + FileLen(fso.BuildPath(fld.path, _
      FileName))
      nFiles = nFiles + 1
      List1.AddItem fso.BuildPath(fld.path, FileName)  ' Load ListBox
      FileName = Dir()  ' Get next file
      DoEvents
   Wend
   Label1 = "Searching " & vbCrLf & fld.path & "..."
   nDirs = nDirs + 1
   If fld.SubFolders.Count > 0 Then
      For Each tFld In fld.SubFolders
         DoEvents
         FindFile = FindFile + FindFile(tFld.path, sFile, nDirs, nFiles)
      Next
   End If
   Exit Function
Catch:  FileName = ""
       Resume Next
End Function
Sub Update(Optional UpdateMode As String = "Normal")

    SetGlobals

    ' Detect why Update has been called and write to log if just started
    If UpdateMode = "Start" Then
        Call LogInformation("DrawingFinder: +++ Started +++")
        Exit Sub
    End If
    If UpdateMode = "AutoUpdate" Then
        Call LogInformation("DrawingFinder: Started")
    End If
    
    ' Import latest SAP data and create a new index
    
    ' Password protect process, if running in auto update mode go straight to update.
    If UpdateMode = "Normal" Then
        Password = InputBox("Enter Password:", "Update Security")
        Select Case Password
            Case "menu"
                RestoreToolbars
            
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
                    If DirExists(NetProgramPath) Then
                        ImportSAP
                        ' Save & set to read-only
                        ThisWorkbook.Save
                        SetAttr FinderFile, vbReadOnly
                        ' Write to log file
                        Call LogInformation("Update: Finished")
                        Call CloseSheet
                    End If
                End If
                
            Case "macro"
                ' Show VB editor, Trust access to the VBA project object model check box must be ticked in Excel Options, Trust Centre, Trust Centre Options, Macro Settings
                Application.VBE.MainWindow.Visible = True
            Case Else
                MsgBox "Incorrect password"
        End Select
    End If
End Sub
Sub ImportSAP()
' Update Macro
' Merge the Design Note and Drawing State exports from SAP into this spreadsheet and update the index file.
  
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

End Sub
Sub CreateIndexes()
' Generate new index files
    SetGlobals
    
    Call CreateIndexFile(GlobalCurrentIndexFile, GlobalCurrentIssueFolder, True)
    Call CreateIndexFile(GlobalOldIndexFile, GlobalOldIssueFolder, True)
    
End Sub
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
Sub WriteIndexUsingDos(Index As String, SourcePath As String)

    Dim Hide As Boolean
    Dim TaskId As Long

    Hide = True    ' Set to true for normal operation, set to false to allow cmd windows to be seen.
    
    ' Update current_iss index
    If Hide Then
        Cmd = Environ$("comspec") & " /c " & "dir " & SourcePath & " /s/b > " & Index
        TaskId = Shell(Cmd, vbHide)
    Else
        Cmd = Environ$("comspec") & " /k " & "dir " & SourcePath & " /s/b > " & Index
        TaskId = Shell(Cmd, vbNormalFocus)
    End If

    ' Wait for process to finish
    Do While ShlProc_IsRunning(TaskId) = True
        DoEvents
    Loop

End Sub
Sub WriteIndexFile(FileNum As Integer, SourcePath As String)
    Set MyObject = New Scripting.FileSystemObject   ' Needs Microsoft Scripting Runtime from Tools - References menu
    Set MySource = MyObject.GetFolder(SourcePath)
    
    Dim PathText As String
    
    On Error Resume Next

    For Each MyFile In MySource.Files
        PathText = MyFile.path
        Print #FileNum, PathText
    Next

    For Each MySubFolder In MySource.SubFolders
        Call WriteIndexFile(FileNum, MySubFolder.path)
    Next
End Sub
Sub Reset_Range()

' Delete all contents and formatting (by deleting rows) so that CTRL+End works correctly.

    Range("B8").Select
    Selection.End(xlDown).Offset(1, 0).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Delete Shift:=xlUp
    x = ActiveSheet.UsedRange.Rows.Count
    ActiveCell.SpecialCells(xlLastCell).Select
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

    LogMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & UserNameWindows & " --- " & LogMessage & " ---"
    FileNum = FreeFile ' next file number
    Open GlobalLogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub
Public Function IsFileWritable(ByVal filePath As String) As Boolean
    On Error Resume Next
    Err.Clear
    
    Dim nFileNum As Integer
    
    nFileNum = FreeFile
    
    Open filePath For Append As nFileNum
    Print #nFileNum, " "
    Close nFileNum
    IsFileWritable = (Err.Number = 0)
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
Public Function ShlProc_IsRunning(ShellReturnValue As Long) As Boolean

    Dim lnghProcess As Long
    Dim lExitCode As Long
    Dim lRet As Long
    
    '//Get the process handle
    lnghProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, ShellReturnValue)
    If lnghProcess <> 0 Then
        '// The GetExitCodeProcess Function retrieves the
        '// termination status of the specified process.
        GetExitCodeProcess lnghProcess, lExitCode
        If lExitCode <> 0 Then
            '// Process still Running
            ShlProc_IsRunning = True
        Else
            '// Process completed
            ShlProc_IsRunning = False
        End If
    End If
End Function

