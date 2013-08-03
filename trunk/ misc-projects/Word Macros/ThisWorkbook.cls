VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' This must match the value stored in the DrawingFinder module.
Const AutoUpdateHour As String = "23"   ' Hour of time to start update
Private Sub Workbook_Open()
' Automatically update on opening spreadsheet.
    Dim StartMode As String
    StartMode = "Start"
    ' If the program has been started at the defined time then auto update.
    HourNow = Left(Time, InStr(Time, ":") - 1)
    If HourNow = AutoUpdateHour Then StartMode = "AutoUpdate"

    Call Update(StartMode)
End Sub

'this code needs to under ThisWorkbook object, when the workbook is activated it will remove all the toolbars.

Private Sub Workbook_Activate()
    RemoveToolbars
End Sub

'this code needs to under ThisWorkbook object agin, when the workbook is de-activated it will restore all the toolbars.

Private Sub Workbook_Deactivate()
    RestoreToolbars
End Sub

'below two sub-routines should be under module

Sub RemoveToolbars()
    Application.ScreenUpdating = False
'The following line stops copy and paste from working.
'   ActiveWindow.DisplayHeadings = False
    With Application
        .DisplayFullScreen = True
        .CommandBars("Full Screen").Visible = False
    End With
    On Error GoTo 0
End Sub
Sub RestoreToolbars()
    Application.ScreenUpdating = False
    On Error GoTo 0
'The following line stops copy and paste from working.
'    ActiveWindow.DisplayHeadings = True
    With Application
        .DisplayFullScreen = False
    End With
End Sub
