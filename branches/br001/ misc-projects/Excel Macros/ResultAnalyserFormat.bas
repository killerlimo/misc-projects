Attribute VB_Name = "ResultAnalyserFormat"
'
' Format test data from John Hytche's Result Analyser tool so that each test is on a new line.
'
' Instructions
' 1. Export results as csv in Result Analyser
' 2. Open csv file in Excel
' 3. Import macro
' 4. Run macro
' 5. Converstion takes a while if the csv file is large
'
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
Sub ResultAnalyserFormat()
'
' Put each test result on a new line and copy the serial and attempt numbers to new row.
'
    Dim i As Long
    Dim MaxRows As Long
    
    MaxRows = ActiveSheet.Rows.CountLarge
    
    'Write row headers
    Cells(1, 1) = "Serial No."
    Cells(1, 2) = "Attempt"
    Cells(1, 3) = "Time"
    Cells(1, 4) = "Test"
    Cells(1, 5) = "Low Limit"
    Cells(1, 6) = "High Limit"
    Cells(1, 7) = "Value"
    Cells(1, 8) = "Units"
    Cells(1, 9) = "Result"
    Rows(1).Font.Bold = True
    
    For i = 2 To MaxRows
        SerialNum = Cells(i, 1)
        Attempt = Cells(i, 2)
        TimeStamp = Cells(i, 3)
        
        ' Look for more results
        While (Cells(i, 10) <> "")
            ' Create new row, add serial num, attempt, time & increase max row count
            Cells(i + 1, 1).Select
            Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i, 1).Select
            Cells(i + 1, 1) = SerialNum
            Cells(i + 1, 2) = Attempt
            Cells(i + 1, 3) = TimeStamp
            
            MaxRows = MaxRows + 1
            ' Copy rest of row to next row
            MaxCols = ActiveSheet.Cells(i, ActiveSheet.Columns.Count).End(xlToLeft).Column
            For j = 10 To MaxCols
                temp = Cells(i, j)
                Cells(i + 1, j - 6) = temp
                ' Delete original cell
                Cells(i, j) = ""
            Next j
        Wend
        
    Next i
    
    MsgBox ("All Done!")
End Sub
Public Function MsgBoxDelay(cMessage, cTitle As String, Timeout As Integer) As Long

    MsgBoxDelay = MessageBoxTimeout(FindWindow(vbNullString, Title), cMessage, cTitle, vbOK, 0, Timeout * 1000)
    
End Function

