Attribute VB_Name = "DrawingFinder"
'Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal lnghProcess As Long, lpExitCode As Long) As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Num As String
Public Desc As String
Public FinderFile As String
Public RepositoryFolder As String
Public TransferFolder As String
Public IndexFile As String
Public TransferIndexFile As String
Public BatchFile As String
Public LogFile As String
Public DataArray(1 To 10) As String
Public filepath As String
Public drive As String

Const ShowDurationSecs As Integer = 5
Const Current = 1
Const Latest = 2
Const ECR = 3

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
Public Sub SetCurrentGlobals()
' Global variables for opening 1_current_iss
    If DirExists("\\atle.bombardier.com\data\uk\pl\dos2") Then
        FinderFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DrawingFinder.xls"
        RepositoryFolder = "\\atle.bombardier.com\data\uk\pl\dos2\1_current_iss"
        TransferFolder = """\\atle.bombardier.com\data\uk\pl\dos\1_files for filing"""
        IndexFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\CurrentIndex.txt"
'        TransferIndexFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\TransferIndex.txt"
        TransferIndexFile = "c:\windows\temp\TransferIndex.txt"
        BatchFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\CreateIndex.bat"
'        LogFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DrawingFinderLogFile.txt"
        LogFile = "c:\windows\temp\DrawingFinderLogFile.txt"
    Else
        drive = Switch(DirExists("e:\1_current_iss"), "e", DirExists("f:\1_current_iss"), "f", DirExists("g:\1_current_iss"), "g", DirExists("c:\1_current_iss"), "c", True, "Not Found")
        If drive = "Not Found" Then
            MsgBox ("Current Issue" & vbLf & "Folder not found")
            End
        Else
                FinderFile = drive & ":\drgstate\DrawingFinder.xls"
                RepositoryFolder = drive & ":\1_current_iss"
                TransferFolder = """:\1_files for filing"""
                IndexFile = drive & ":\drgstate\CurrentIndex.txt"
                TransferIndexFile = ":\drgstate\TransferIndex.txt"
                BatchFile = drive & ":\drgstate\CreateIndex.bat"
                LogFile = drive & ":\drgstate\DrawingFinderLogFile.txt"
        End If
    End If
End Sub
Public Sub SetOldGlobals()
' Global variables for opening 1_old_iss
    If DirExists("\\atle.bombardier.com\data\uk\pl\dos2") Then
        RepositoryFolder = "\\atle.bombardier.com\data\uk\pl\dos2\1_Old_iss"
        IndexFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\OldIndex.txt"
        BatchFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\CreateIndex.bat"
'        LogFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DrawingFinderLogFile.txt"
        LogFile = "c:\windows\temp\DrawingFinderLogFile.txt"
    Else
        drive = Switch(DirExists("e:\1_current_iss"), "e", DirExists("f:\1_current_iss"), "f", DirExists("g:\1_current_iss"), "g", DirExists("c:\1_current_iss"), "c", True, "Not Found")
        If drive = "Not Found" Then
            MsgBox ("Current Issue" & vbLf & "Folder not found")
            End
        Else
            RepositoryFolder = drive & ":\1_Old_iss"
            IndexFile = drive & ":\drgstate\OldIndex.txt"
            BatchFile = drive & ":\drgstate\CreateIndex.bat"
            LogFile = drive & ":\drgstate\DrawingFinderLogFile.txt"
        End If
    
        If Not (DirExists(RepositoryFolder)) Then
            MsgBox (RepositoryFolder & vbLf & "Folder not found")
        End If
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
Sub OpenItem(IssueRequest As Variant)
'
' Use with DrgstateSAP.xlsx to search for and open files directly.
' Periodically when DrgstateSAP.xlsx is updated run the following command to create the index:
' dir \\atle.bombardier.com\data\UK\PL\DOS2\1_current_iss /s/b > \\atle.bombardier.com\data\uk\pl\dos\DrgState\index.txt
' Use DrgstateSAP.xlsx to find the Item number wanted and click in the Item number cell.
' Run the find macro using CTRL+L and if only one file exists it will be opened.
' If there are multiple files then a menu will be presented listing the files available.
' Choose the menu option number and that file will be opened.
'
    Dim Cmd As String
    Dim file As String
    Dim reply As Variant
    
    Dim strBuf As String
    Dim intIndex As Integer
    Dim TaskId As Long
    Dim RepoDate, IndexDate As Variant

    ' Locate file in index and return full path to file (s)
    ' Look in first column only
    file = Cells(ActiveCell.Row, 1).Value
    issue = Cells(ActiveCell.Row, 3).Value
    correction = Cells(ActiveCell.Row, 4).Value
    ECRfile = Cells(ActiveCell.Row, 6).Value
    
    ' Find and replace '/' with '-' for file name.
    file = Replace(file, "/", "-")
    ' Format ECR number to match filenames
    ECRfile = Replace(ECRfile, "600000000000", "6-")
    ECRfile = Replace(ECRfile, "60000000000", "6-")
    ECRfile = Replace(ECRfile, "6000000000", "6-")
    ECRfile = Replace(ECRfile, "600000000", "6-")
    ECRfile = Replace(ECRfile, "60000000", "6-")
    ECRfile = Replace(ECRfile, "6000000", "6-")
       
    'Generate full file name for old issue
    If IssueRequest = ECR Then
        Item = ECRfile
    ElseIf IssueRequest = Current Then
        Item = file
    Else
        Item = file & "-" & issue & correction
    End If
    
    ' Check for null string
    If Item = "" Then
        MsgBox ("No drawing selected")
        Exit Sub
    End If
    
    ' Create strings for log entry
    Select Case IssueRequest
        Case "1"
            RequestStr = "Latest"
        Case "2"
            RequestStr = "Old"
        Case "3"
            RequestStr = "ECR"
    End Select
    ' Write seach item to log
    Call LogInformation("OpenItem: Request: " & RequestStr & " : " & Item)

    ' Search for item in index file
    If CheckIndexes(Item, IndexFile) = False Then
        If CheckIndexes(Item, TransferIndexFile) = False Then
            ' no paths returned from search
            MsgBox ("File not found")
            Call LogInformation("OpenItem: File not found")
        End If
    End If
End Sub
Function CheckIndexes(Item, Index As String) As Boolean
' Allow 1_current_iss and 1_files for filing to be searched separately

    ' Call find and wait for process to finish
        
    Set Sh = CreateObject("WScript.Shell")
    Cmd = Environ$("comspec") & " /c find /i """ & Item & """ " & Index & " > " & ResultFile
    ReturnCode = Sh.Run(Cmd, 1, True)
    Open IndexFile For Input As #1
    Line = 0
    ' while not eof or max array size
    Do Until EOF(1) Or Line = 9
        Input #1, DrawingPath
        If InStr(DrawingPath, Item) Then
            Line = Line + 1
            DataArray(Line) = DrawingPath
        End If
    Loop
    Close #1

    ' More than 1 line indicates that at least 1 file has been found
    If Line > 0 Then
        For intIndex = 1 To Line
            strBuf = strBuf & intIndex & ". " & Right(GetFilename(DataArray(intIndex)), 100) & vbLf
        Next
        Choice = -9
        CheckIndexes = True
        
        ' Indicate that search has been successful
        
        Do Until (Choice > 0) And (Choice <= Line)
            Ch = InputBox(strBuf, "Choose File:", 1)
        ' Check for Escape key
        If Ch = "" Then Exit Function Else Choice = Int(Ch)
        Loop
        If Choice > 0 Then
            filepath = DataArray(Choice)
            ' Create link to file
            link = "file:///" & filepath
            ' Open file in applicaion
            ActiveWorkbook.FollowHyperlink link
        End If
    Else
        CheckIndexes = False

    End If
End Function
Sub TextFilter()
'
' TextFilter Macro
' Set up data filters on Item No. and Description columns.
' Enter up to 2 words in each search box, these will be OR'd

    Dim NumWords() As String
    Dim DescWords() As String
    
    Dim NumWordCount, DescWordCount As Integer
    Dim test() As String
    
    NumWordCount = 4
    DescWordCount = 4
    
    Do Until (NumWordCount < 3)
        ' Get search data from user & convert to upper case
        Num = StrConv(InputBox("Enter part of the Drawing Number" & vbLf & "Up to 2 words can be entered" & vbLf _
        & "Use & | for AND OR", "Drawing Search", Num), vbUpperCase)
        ' Split the input strings into an array of words.
        If InStr(1, Num, "&") > 0 Then
            NumWords = Split(Num, "&")
        ElseIf InStr(1, Num, "|") > 0 Then
            NumWords = Split(Num, "|")
        Else: NumWords = Split(Num)
        End If
        If Num = "" Then Num = "*"
        ' Return the number of elements in the arrays.
        NumWordCount = UBound(NumWords) - LBound(NumWords) + 1
    Loop

    Do Until (DescWordCount < 3)
        ' Get search data from user & convert to upper case
        Desc = StrConv(InputBox("Enter part of the Drawing Description" & vbLf & "Up to 2 words can be entered" & vbLf _
        & "Use & | for AND OR", "Drawing Search", Desc), vbUpperCase)
        If InStr(1, Desc, "&") > 0 Then
            DescWords = Split(Desc, "&")
        ElseIf InStr(1, Desc, "|") > 0 Then
            DescWords = Split(Desc, "|")
        Else: DescWords = Split(Desc)
        End If
        If Desc = "" Then Desc = "*"
        DescWordCount = UBound(DescWords) - LBound(DescWords) + 1
    Loop
        
    SetCurrentGlobals
    ' Write search terms to log
    Call LogInformation("TextFilter: Drawing: " & Num & " Description: " & Desc)
        
    Range("A7").Select
    
    If Num = "*" Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=1, _
        Criteria1:="=*", Operator:=xlAnd
    ElseIf InStr(1, Num, "&") = 0 And InStr(1, Num, "|") = 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=1, _
        Criteria1:="=*" & NumWords(0) & "*", Operator:=xlAnd
    ElseIf InStr(1, Num, "&") > 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=1, _
        Criteria1:="=*" & NumWords(0) & "*", Operator:=xlAnd, Criteria2:="=*" & NumWords(1) & "*"
    ElseIf InStr(1, Num, "|") > 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=1, _
        Criteria1:="=*" & NumWords(0) & "*", Operator:=xlOr, Criteria2:="=*" & NumWords(1) & "*"
    End If
    
    If Desc = "*" Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=2, _
        Criteria1:="=*", Operator:=xlAnd
    ElseIf InStr(1, Desc, "&") = 0 And InStr(1, Desc, "|") = 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=2, _
        Criteria1:="=*" & DescWords(0) & "*", Operator:=xlAnd
    ElseIf InStr(1, Desc, "&") > 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=2, _
        Criteria1:="=*" & DescWords(0) & "*", Operator:=xlAnd, Criteria2:="=*" & DescWords(1) & "*"
    ElseIf InStr(1, Desc, "|") > 0 Then
        ActiveSheet.Range(Selection, Selection.SpecialCells(xlLastCell)).AutoFilter Field:=2, _
        Criteria1:="=*" & DescWords(0) & "*", Operator:=xlOr, Criteria2:="=*" & DescWords(1) & "*"
    End If
    
    ActiveWindow.SmallScroll Down:=-10000
End Sub
Sub ImportSAP()
' Update Macro
' Merge the Design Note and Drawing State exports from SAP into this spreadsheet and update the index file.
'
    DesignNoteFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DesignNoteStateSAP.xlsx"
    DrawingFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DrgStateSAP.xlsx"
  
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
    
    Workbooks.Open Filename:=DrawingFile
    Range("A8").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Windows("DrawingFinderNoDos.xls").Activate
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
' Move to end of data

    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
' Open Design Notes spreadsheet and paste into this sheet

    Workbooks.Open Filename:=DesignNoteFile
    Range("A8").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Windows("DrawingFinderNoDos.xls").Activate
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
    
    Dim TaskId As Long
    Dim Hide As Boolean
    
    Hide = True    ' Set to true for normal operation, set to false to allow cmd windows to be seen.
    
    ' Update current_iss index
    SetCurrentGlobals
    If Hide Then
        Cmd = Environ$("comspec") & " /c " & "dir " & RepositoryFolder & " /s/b > " & IndexFile
        TaskId = Shell(Cmd, vbHide)
    Else
        Cmd = Environ$("comspec") & " /k " & "dir " & RepositoryFolder & " /s/b > " & IndexFile
        TaskId = Shell(Cmd, vbNormalFocus)
    End If

    ' Wait for process to finish
    Do While ShlProc_IsRunning(TaskId) = True
        DoEvents
    Loop

    ' Update old_iss index
    SetOldGlobals
    If Hide Then
        Cmd = Environ$("comspec") & " /c " & "dir " & RepositoryFolder & " /s/b > " & IndexFile
        TaskId = Shell(Cmd, vbHide)
    Else
        Cmd = Environ$("comspec") & " /k " & "dir " & RepositoryFolder & " /s/b > " & IndexFile
        TaskId = Shell(Cmd, vbNormalFocus)
    End If

    ' Wait for process to finish
    Do While ShlProc_IsRunning(TaskId) = True
        DoEvents
    Loop

End Sub
Sub CreateTransferIndex()
' Generate new index files
    
    Dim TaskId As Long
    Dim Hide As Boolean
    
    Hide = True    ' Set to true for normal operation, set to false to allow cmd windows to be seen.
    
    ' Update current_iss index
    ' SetCurrentGlobals
    
    ' Index 1_files for filing contents
    
    If Hide Then
        Cmd = Environ$("comspec") & " /c " & "dir " & TransferFolder & " /s/b > " & TransferIndexFile
        TaskId = Shell(Cmd, vbHide)
    Else
        Cmd = Environ$("comspec") & " /k " & "dir " & TransferFolder & " /s/b > " & TransferIndexFile
        TaskId = Shell(Cmd, vbNormalFocus)
    End If

    ' Wait for process to finish
    Do While ShlProc_IsRunning(TaskId) = True
        DoEvents
    Loop

End Sub
Sub Update(Optional UpdateMode As String = "Normal")

    SetCurrentGlobals

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
    Else
        Password = "eng"
        ' Write to log file
        Call LogInformation("Update: Starting Scheduled")
    End If
    
    If Password = "eng" Then
        If UpdateMode = "Normal" Then Call LogInformation("Update: Starting Manual")
        ' Provide a method of escaping from the update.
        If MsgBoxDelay("Click OK to cancel", "Update", ShowDurationSecs) = 1 Then
            RestoreToolbars
        ' Write to log file
        Call LogInformation("Update: Aborted by user")
    Else
        
        If DirExists("\\atle.bombardier.com\data\uk\pl\dos2") Then
            ImportSAP
        End If
        
        Call MsgBoxDelay("Creating Indexes...", "Update", ShowDurationSecs)
        
        CreateIndexes
        ' Save & set to read-only
        ThisWorkbook.Save
        SetAttr FinderFile, vbReadOnly
        Call MsgBoxDelay("...Indexes Created", "Update", ShowDurationSecs)
        
        ' Write to log file
        Call LogInformation("Update: Finished")
        Call CloseSheet
    End If
    
    ElseIf Password = "menu" Then
        RestoreToolbars
    Else
        MsgBox "Incorrect password"
    End If
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
Sub ShowInFolder(Latest As Variant)

' Show a list of files matching the Item name.
' The Item name is picked up from a cell in the spreadsheet.
' Allow one to be selected and show it in its folder.

    FindFiles (Latest)

End Sub
Sub FindFiles(Latest As Variant)

' Search for the Item in the index and create a file containing a list of paths to the file(s) found.
    
    ' Locate file in index and return full path to file (s)
    ' Look in first column only
    file = Cells(ActiveCell.Row, 1).Value
    issue = Cells(ActiveCell.Row, 3).Value
    correction = Cells(ActiveCell.Row, 4).Value
    ECRfile = Cells(ActiveCell.Row, 6).Value
    
    ' Find and replace '/' with '-' for file name.
    file = Replace(file, "/", "-")
    ' Format ECR number to match filenames
    ECRfile = Replace(ECRfile, "600000000000", "6-")
    ECRfile = Replace(ECRfile, "60000000000", "6-")
    ECRfile = Replace(ECRfile, "6000000000", "6-")
    ECRfile = Replace(ECRfile, "600000000", "6-")
    ECRfile = Replace(ECRfile, "60000000", "6-")
    ECRfile = Replace(ECRfile, "6000000", "6-")
       
    'Generate full file name for old issue
    If Latest = ECR Then
        Item = ECRfile
    ElseIf Latest = Current Then
        Item = file
    Else
        ' Check for null file
        If file <> "" Then Item = file & "-" & issue & correction
    End If
    
    ' Check for null string
    If Item = "" Then
        MsgBox ("No drawing selected")
        Exit Sub
    End If
    
' Display list of files with an option number.

    ' Search for item in index file
    Open IndexFile For Input As #1
    Line = 0
    ' while not eof or max array size
    Do Until EOF(1) Or Line = 9
        Input #1, DrawingPath
        If InStr(DrawingPath, Item) Then
            Line = Line + 1
            DataArray(Line) = DrawingPath
        End If
    Loop
    Close #1


    ' More than 1 line indicates that at least 1 file has been found
    If Line > 0 Then
        If Line > 1 Then
            For intIndex = 1 To Line
                strBuf = strBuf & intIndex & ". " & Right(GetFilename(DataArray(intIndex)), 40) & vbLf
            Next
            Choice = -9
            
            Do Until (Choice > 0) And (Choice <= Line)
                Ch = InputBox(strBuf, "Choose file:", 1)
                ' Protect further code from pressing Escape
                If Ch = "" Then Exit Sub Else Choice = Int(Ch)
            Loop
            filepath = DataArray(Choice)
        Else
            filepath = DataArray(Line)
            Choice = 1
        End If
        ' If a selection has been made then open Windows Explorer showing the correct folder
        If Choice <> 0 Then
            Call LogInformation("DisplayList: File path: " & filepath)
            'Shell "explorer /e, /select," & filepath, vbNormalFocus    ' Opening showing the folders over the network is very slow!
            Shell "explorer /select," & filepath, vbNormalFocus
        End If
    Else
        ' no paths returned from search
        MsgBox ("File not found")
        Call LogInformation("DisplayList: File not found")
    End If
    
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
Sub Show()
' Display show menu and allow choice of current issue, old issue or show in folder

    IssueChoice = -9
    
    Do Until (IssueChoice > 0) And (IssueChoice <= 3)
        Ch = InputBox("1. Latest Issue" & vbLf & "2. Old Issue" & vbLf & "3. ECR", "Choose option:", 1)
        ' Check for Escape key
        If Ch = "" Then Exit Sub Else IssueChoice = Int(Ch)
    Loop
    
    ' Create strings for log entry
        
    Select Case Ch
        Case "1"
            ChoiceStr = "Latest"
        Case "2"
            ChoiceStr = "Old"
        Case "3"
            ChoiceStr = "ECR"
    End Select
    
    ActionChoice = -9
    
    If IssueChoice <> 0 Then
        Do Until (ActionChoice > 0) And (ActionChoice <= 2)
            Ch = InputBox("1. Open document" & vbLf & "2. Show in folder", "Choose action:", 1)
            ' Check for Escape key
            If Ch = "" Then Exit Sub Else ActionChoice = Int(Ch)
        Loop
 
        Select Case Ch
            Case "1"
                ActionStr = "Open"
            Case "2"
                ActionStr = "Show in folder"
        End Select
                
        If ActionChoice <> 0 Then
            ' Set appropriate globals
            If IssueChoice <> "2" Then
                SetCurrentGlobals
            Else
                SetOldGlobals
            End If
            
            Call LogInformation("Show: Choice: " & ChoiceStr & " Action: " & ActionStr)
            ' Index Transfer Directory
            CreateTransferIndex
            
            ' Carry out appropriate action
            If ActionChoice = 1 Then
                OpenItem (IssueChoice)
            Else
                ShowInFolder (IssueChoice)
            End If
        End If
    End If
End Sub
Sub Tutorial()
    ' Create link to tutorial
    TutorialFile = "\\atle.bombardier.com\data\UK\PL\DOS\DrgState\DrawingFinderTutorial.pdf"
    link = "file:///" & TutorialFile
    If FileExists(TutorialFile) Then
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
    Open LogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub
Function UserNameWindows() As String
     UserNameWindows = Environ("USERNAME")
End Function
Sub CloseSheet()
' Close sheet to allow for scheduled update
    Dim wBook As Workbook
    
    On Error Resume Next    ' Needed to allow workbook to be set even if workbook is not open
    
    Call LogInformation("CloseSheet: Application Closed")
 
    Set wBook = Workbooks("DrawingFinder.xls")
    If Not (wBook Is Nothing) Then
'        RestoreToolbars
'        Application.EnableEvents = False    ' For some reason the close events prevent the workbook from closing
        DoEvents
        Workbooks("DrawingFinder.xls").Close
    End If
End Sub

