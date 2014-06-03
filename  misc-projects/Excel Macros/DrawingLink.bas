Attribute VB_Name = "DrawingLink"
'Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal lnghProcess As Long, lpExitCode As Long) As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Num As String
Public Desc As String
Public RepositoryFolder As String
Public IndexFile As String
Public ResultFile As String
Public BatchFile As String
Public DataArray(1 To 10) As String
Public filepath As String
Public drive As String
Public Sub SetCurrentGlobals()
' Global variables for opening 1_current_iss
    If DirExists("\\atle.bombardier.com\data\uk\pl\dos2") Then
        FinderFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\DrawingFinder.xls"
        RepositoryFolder = "\\atle.bombardier.com\data\uk\pl\dos2\1_current_iss"
        IndexFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\CurrentIndex.txt"
        ResultFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\CurrentResult.txt"
        BatchFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\CreateIndex.bat"
    Else 'Look for folder locally, usually for development purposes.
        drive = Switch(DirExists("e:\1_current_iss"), "e", DirExists("f:\1_current_iss"), "f", DirExists("g:\1_current_iss"), "g", DirExists("c:\1_current_iss"), "c", True, "Not Found")
        If drive = "Not Found" Then
            MsgBox ("Current Issue" & vbLf & "Folder not found")
            End
        Else
            FinderFile = drive & ":\drgstate\DrawingFinder.xls"
            RepositoryFolder = drive & ":\1_current_iss"
            IndexFile = drive & ":\drgstate\CurrentIndex.txt"
            ResultFile = drive & ":\drgstate\CurrentResult.txt"
            BatchFile = drive & ":\drgstate\CreateIndex.bat"
        End If
    End If
End Sub
Public Sub SetCurrentPartGlobals()
' Global variables for opening 1_Parts PDF Datasheets
    If DirExists("\\atle.bombardier.com\data\uk\pl\dos2") Then
        RepositoryFolder = """\\atle.bombardier.com\data\uk\pl\dos2\1_Parts PDF Datasheets"""
        IndexFile = "\\atle.bombardier.com\data\uk\pl\dos\Drgstate\PartsCurrentIndex.txt"
        ResultFile = "\\atle.bombardier.com\data\uk\pl\dos\Drgstate\PartsCurrentResult.txt"
        BatchFile = "\\atle.bombardier.com\data\uk\pl\dos\Drgstate\PartsCreateIndex.bat"
    Else 'Look for folder locally, usually for development purposes.
        drive = Switch(DirExists("e:\1_Parts PDF Datasheets"), "e", DirExists("f:\1_Parts PDF Datasheets"), "f", DirExists("g:\1_Parts PDF Datasheets"), "g", DirExists("c:\1_Parts PDF Datasheets"), "c", True, "Not Found")
        If drive = "Not Found" Then
            MsgBox ("Parts Datasheet" & vbLf & "Folder not found")
            End
        Else
            RepositoryFolder = """" & drive & ":\Parts PDF Datasheets" & """"
            IndexFile = """" & drive & ":\Drgstate\PartsCurrentIndex.txt" & """"
            ResultFile = drive & ":\Drgstate\PartsCurrentResult.txt"
            BatchFile = drive & ":\Drgstate\PartsCreateIndex.bat"
        End If
    End If
End Sub
Public Sub SetOldGlobals()
' Global variables for opening 1_old_iss
    If DirExists("\\atle.bombardier.com\data\uk\pl\dos2") Then
        RepositoryFolder = "\\atle.bombardier.com\data\uk\pl\dos2\1_Old_iss"
        IndexFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\OldIndex.txt"
        ResultFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\OldResult.txt"
        BatchFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\CreateIndex.bat"
    Else
        RepositoryFolder = drive & ":\1_Old_iss"
        IndexFile = drive & ":\drgstate\OldIndex.txt"
        ResultFile = drive & ":\drgstate\OldResult.txt"
        BatchFile = drive & ":\drgstate\CreateIndex.bat"
    End If
    If Not (DirExists(RepositoryFolder)) Then
        MsgBox (RepositoryFolder & vbLf & "Folder not found")
        End
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
Sub OpenItem(item As String)
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
    Dim reply As Variant
    
    Dim strBuf As String
    Dim intIndex As Integer
    Dim TaskId As Long
    Dim RepoDate, IndexDate As Variant

    ' Call find and wait for process to finish
        
    Set Sh = CreateObject("WScript.Shell")
    Cmd = Environ$("comspec") & " /c find /i """ & item & """ " & IndexFile & " > " & ResultFile
    ReturnCode = Sh.Run(Cmd, 1, True)

    ' Read in paths to found files
    Open ResultFile For Input As #1
    Line = 0
    ' while not eof or max array size
    Do Until EOF(1) Or Line = 9
        Line = Line + 1
        Input #1, DataArray(Line)
    Loop
    Close #1

    ' More than 2 lines indicates that at least 1 file has been found
    If Line > 2 Then
        For intIndex = 3 To Line
            strBuf = strBuf & intIndex - 2 & ". " & Right(GetFilename(DataArray(intIndex)), 100) & vbLf
        Next
        Choice = -9
        
        Do Until (Choice > 0) And (Choice < Line)
            Ch = InputBox(strBuf, "Choose File:", 1)
        ' Check for Escape key
        If Ch = "" Then Exit Sub Else Choice = Int(Ch)
        Loop
        If Choice > 0 Then
            filepath = DataArray(Choice + 2)
            ' Create link to file
            link = "file:///" & filepath
            ' Open file in applicaion
            ActiveWorkbook.FollowHyperlink link
        End If
    Else
        ' no paths returned from search
        MsgBox ("No datasheet found")
    End If
End Sub
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
Sub ShowInFolder(Latest As Boolean)

' Show a list of files matching the Item name.
' The Item name is picked up from a cell in the spreadsheet.
' Allow one to be selected and show it in its folder.

    FindFiles (Latest)
    DisplayList

End Sub
Sub FindFiles(Latest As Boolean)

' Search for the Item in the index and create a file containing a list of paths to the file(s) found.
    
    ' Locate file in index and return full path to file (s)
    ' Look in first column only
    file = Cells(ActiveCell.Row, 1).Value
    issue = Cells(ActiveCell.Row, 3).Value
    correction = Cells(ActiveCell.Row, 4).Value
    
    ' Find and replace '/' with '-' for file name. SAP uses '/' file system can't.
    file = Replace(file, "/", "-")
    'Generate full file name
    If Latest Then
        item = file
    Else
        item = file & "-" & issue & correction
    End If

    ' Call find and wait for process to finish
    
    Set Sh = CreateObject("WScript.Shell")
    Cmd = Environ$("comspec") & " /c find /i """ & item & """ " & IndexFile & " > " & ResultFile
    ReturnCode = Sh.Run(Cmd, 1, True)

End Sub
Sub DisplayList()

' Open the result file created by the last search.
' Display list of files with an option number.

Open ResultFile For Input As #1
    Line = 0
    ' while not eof or max array size
    Do Until EOF(1) Or Line = 9
        Line = Line + 1
        Input #1, DataArray(Line)
        ' Only store path  names
        ' If Line > 2 Then DataArray(Line) = GetPath(DataArray(Line))
    Loop
    Close #1

    ' More than 2 lines indicates that at least 1 file has been found
    If Line > 2 Then
        If Line > 3 Then
            For intIndex = 3 To Line
                strBuf = strBuf & intIndex - 2 & ". " & Right(GetFilename(DataArray(intIndex)), 40) & vbLf
            Next
            Choice = -9
            
            Do Until (Choice > 0) And (Choice <= Line - 2)
                Ch = InputBox(strBuf, "Choose file:", 1)
                ' Protect further code from pressing Escape
                If Ch = "" Then Exit Sub Else Choice = Int(Ch)
            Loop
            filepath = DataArray(Choice + 2)
        Else
            filepath = DataArray(Line)
            Choice = 1
        End If
        ' If a selection has been made then open Windows Explorer showing the correct folder
        If Choice <> 0 Then
            'Shell "explorer /e, /select," & filepath, vbNormalFocus    ' Opening showing the folders over the network is very slow!
            Shell "explorer /select," & filepath, vbNormalFocus
        End If
    Else
        ' no paths returned from search
        MsgBox ("File not found")
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

Dim name As String

    ' Determine whether anything has already been selected
    name = Selection
    ' Strip out any line feed type chars & replace / with -
    name = Replace(name, vbLf, "")
    name = Replace(name, vbCr, "")
    name = Replace(name, Chr(11), "")
    name = Replace(name, "/", "-")
    If Len(name) = 1 Then
        ' Determine whether number is part or drawing
        Selection.MoveLeft Unit:=wdWord, Count:=1
        Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
        name = Selection
        ' Strip off trailing space
        If Right(name, 1) = " " Then name = Left(name, Len(name) - 1)
    End If
    
    ' Convert name to number to check for part number
    nameval = Val(name)
    ' Check that the name is definitely a part number
    If name = LTrim(Str(nameval)) And (nameval > 100000 And nameval < 127000) Or (nameval > 520000000 And nameval < 530000000) Then
        SetCurrentPartGlobals
        ' Carry out appropriate action
        OpenItem (name)
    Else

        ' Display show menu and allow choice of current issue, old issue or show in folder
    
        IssueChoice = -9
        
        Do Until (IssueChoice > 0) And (IssueChoice <= 2)
            Ch = InputBox("1. Latest Issue" & vbLf & "2. Old Issue", "Choose option:", 1)
            ' Check for Escape key
            If Ch = "" Then Exit Sub Else IssueChoice = Int(Ch)
        Loop
        
        ActionChoice = 1
        
        If IssueChoice <> 0 Then
        '    Do Until (ActionChoice > 0) And (ActionChoice <= 2)
        '        Ch = InputBox("1. Open drawing" & vbLf & "2. Show in folder", "Choose action:", 1)
                ' Check for Escape key
        '        If Ch = "" Then Exit Sub Else ActionChoice = Int(Ch)
        '    Loop
        
            If ActionChoice <> 0 Then
                ' Set appropriate globals
                If IssueChoice = "1" Then
                    SetCurrentGlobals
                Else
                    SetOldGlobals
                End If
                
                ' Carry out appropriate action
                If ActionChoice = 1 Then
                    OpenItem (name)
                Else
                    ShowInFolder (IssueChoice = "1")
                End If
            End If
        End If
    End If
End Sub



