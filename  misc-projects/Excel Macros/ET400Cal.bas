Attribute VB_Name = "ET400Cal"
Sub CalCreate()
Attribute CalCreate.VB_Description = "Create cal values from DtiMonitor decode."
Attribute CalCreate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CalCreate Macro
' Create cal values from DtiMonitor decode.
'
    Dim Address(999) As String
    Dim Data(999) As String
    Dim DataLine As String
    
    SerialAddr = 4
    ModAddr = 6
    GAaddr = 8
    
    FileToOpen = Application.GetOpenFilename(Title:="Please choose decode file to process", FileFilter:="Decode Files *.csv (*.csv),")

    If FileToOpen = False Then
        MsgBox "No file specified.", vbExclamation, "Exiting"
        Exit Sub
    End If
    
    Open FileToOpen For Input As #1
    Line = 0
    
    ' while not eof
    Do Until EOF(1)
        Input #1, DataLine
        
        'Check for config memory entry
        If Mid(DataLine, 3, 1) = ":" Then
            Address(Line) = Left(DataLine, 2)
            Data(Line) = Mid(DataLine, 5, 6)
            Line = Line + 1
        End If
    Loop
    Close #1

    ' Create special values for serial no, m/s & GA part no.
    SerialNo = Val("&H" & Right(Data(SerialAddr + 1), 4) & Right(Data(SerialAddr), 4))
    ModState = Val("&H" & Right(Data(ModAddr), 4))
    GApart = Val("&H" & Right(Data(GAaddr + 1), 4) & Right(Data(GAaddr), 4))

    ' Write cal values to text file
    FileToSave = Replace(FileToOpen, ".csv", "Cal.txt")
    
    Open FileToSave For Output As #1
    
    Print #1, "sn 6238 " & SerialNo
    Print #1, "ms 3794 " & ModState
    Print #1, "ga 4424 " & GApart
        
    For Line = 32 To 51
        CommandLine = "cw " & Address(Line) & " " & Data(Line)
        Print #1, CommandLine
    Next
    
    Close #1
    
    ' Open finished file
    ActiveWorkbook.FollowHyperlink FileToSave
    
End Sub
