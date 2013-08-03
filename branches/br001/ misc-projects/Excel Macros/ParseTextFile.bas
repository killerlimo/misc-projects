Attribute VB_Name = "ParseTextFile"
'~~> Change this to the relevant path
Const strPath As String = "C:\Temp\ET200 Zollner results\"

Sub ParseTextFiles()
    Dim ws As Worksheet
    Dim MyData As String, strData() As String
    Dim txtLine() As String
    Dim WriteToRow As Long, i As Long
    Dim strCurrentTxtFile As String

    Set ws = Sheets("Sheet1")

    '~~> Start from Row 1
    WriteToRow = 1

    strCurrentTxtFile = Dir(strPath & "*.Txt")

    '~~> Looping through all text files in a folder
    Do While strCurrentTxtFile <> ""

        '~~> Open the file in 1 go to read it into an array
        Open strPath & strCurrentTxtFile For Binary As #1
        MyData = Space$(LOF(1))
        Get #1, , MyData
        Close #1

        strData() = Split(MyData, vbCrLf)
        
        '~~> Read from the array and write to Excel
        For i = LBound(strData) To UBound(strData)
            If Left(strData(i), 4) = "VOUT" Then
                txtLine = Split(strData(i), " ")
                WriteToCol = 1
                For j = LBound(txtLine) To UBound(txtLine)
                    If txtLine(j) <> "" Then
                        ws.Cells(WriteToRow, WriteToCol).Value = txtLine(j)
                        WriteToCol = WriteToCol + 1
                    End If
                Next j
                
            End If
        Next i
        
        WriteToRow = WriteToRow + 1
        strCurrentTxtFile = Dir
    Loop

    MsgBox "Done"
End Sub
