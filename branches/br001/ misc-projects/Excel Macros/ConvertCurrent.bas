Attribute VB_Name = "ConvertCurrent"
Const MinLen = 5            ' Set minimum number of digits that cell must contain
Const Multiplier = 1000     ' Set muliplier
Const ShowDurationSecs = 5  ' Message delay

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
Sub ConvertCurrent()
'
' Multiply each cell in every sheet that contains a decimal with more than 5 digits by 1000
'
    Dim rng As Range
    Dim s As Worksheet
    
    Call MsgBoxDelay("This may take a minute...", "Please Wait", ShowDurationSecs)
  
    For Each s In ActiveWorkbook.Worksheets
        Set rng = s.UsedRange
        
        For Each c In rng
            If (c < 1) And (Len(c) > MinLen) Then
                c.Value = Application.WorksheetFunction.Product(c, Multiplier)
            End If
        Next c
    Next s
    
    MsgBox ("All Done!")
End Sub
Public Function MsgBoxDelay(cMessage, cTitle As String, Timeout As Integer) As Long

    MsgBoxDelay = MessageBoxTimeout(FindWindow(vbNullString, Title), cMessage, cTitle, vbOK, 0, Timeout * 1000)
    
End Function

