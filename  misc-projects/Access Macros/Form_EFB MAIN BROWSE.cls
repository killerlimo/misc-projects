VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EFB MAIN BROWSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons
Const Build As String = 1

Private Sub Button57_Click()
On Error GoTo Err_Button57_Click


    DoCmd.Close

Exit_Button57_Click:
    Exit Sub

Err_Button57_Click:
    MsgBox Error$
    Resume Exit_Button57_Click
    
End Sub

Private Sub Button58_Click()
On Error GoTo Err_Button58_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Button58_Click:
    Exit Sub

Err_Button58_Click:
    MsgBox Error$
    Resume Exit_Button58_Click
    
End Sub

Private Sub Button59_Click()
On Error GoTo Err_Button59_Click


    DoCmd.DoMenuItem A_FORMBAR, A_EDITMENU, 10, , A_MENU_VER20

Exit_Button59_Click:
    Exit Sub

Err_Button59_Click:
    MsgBox Error$
    Resume Exit_Button59_Click
    
End Sub

Private Sub Button95_Click()
On Error GoTo Err_Button95_Click


    DoCmd.Close

Exit_Button95_Click:
    Exit Sub

Err_Button95_Click:
    MsgBox Error$
    Resume Exit_Button95_Click
    
End Sub

Private Sub SAVE_AND_EXIT_Click()
On Error GoTo Err_SAVE_AND_EXIT_Click

    DoCmd.Close A_FORM, "efb main"

   
Exit_SAVE_AND_EXIT_Click:
    Exit Sub

Err_SAVE_AND_EXIT_Click:
    MsgBox Error$
    Resume Exit_SAVE_AND_EXIT_Click
    
End Sub
Private Sub ViewEFN_Click()
On Error GoTo Err_ViewEFN_Click

Dim Command As String

Rem Local path
Rem Command = """C:\Users\kevin\Google Drive\Work\EFN Finder\EFN Finder""" & " " & REF
Command = """\\atle.bombardier.com\data\uk\pl\dos\drgstate\EFN Finder\EFN Finder""" & " " & REF

Rem Write log message
Call LogInformation("Starting View EFN:" & REF)

Rem use the /k option to leave the shell window open for debug
Rem Call Shell("cmd /k" & Command, 1)
Call Shell("cmd /c" & Command, 1)

Call LogInformation("Ending View EFN")

Exit_ViewEFN_Click:
    Exit Sub

Err_ViewEFN_Click:
    MsgBox Error$
    Resume Exit_ViewEFN_Click
    
End Sub
Sub LogInformation(LogMessage As String)
' Write to log file

Dim FileNum As Integer
Dim GlobalLogFile As String

    Rem On Error GoTo ErrorHandler
    GlobalLogFile = "\\atle.bombardier.com\data\uk\pl\dos\drgstate\EFN Finder\EFNFinderLog.txt"
    UserNameWindows = Environ("USERNAME")
    
    LogMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & " Build: " & Build & " - " & UserNameWindows & " --- " & LogMessage & " ---"
    FileNum = FreeFile ' next file number
    Open GlobalLogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub
