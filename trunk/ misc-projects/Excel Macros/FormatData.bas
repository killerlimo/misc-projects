Attribute VB_Name = "FormatData"
Sub Format_Data()
Attribute Format_Data.VB_Description = "Format Logged Data & Graph"
Attribute Format_Data.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Format_Data Macro
' Format Logged Data & Graph
'

'
    ActiveCell.SpecialCells(xlLastCell).Select
    Range(Selection, Cells(1)).Select
    Columns("A:A").EntireColumn.AutoFit
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLines
    ActiveChart.SetSourceData Source:=Sheets("log6").Range("A1:D159")
    ActiveChart.Location Where:=xlLocationAsObject, name:="log6"
    With ActiveChart
        .DisplayBlanksAs = xlInterpolated
        .PlotVisibleOnly = True
        .SizeWithWindow = True
    End With
    Application.ShowChartTipNames = True
    Application.ShowChartTipValues = True
End Sub
