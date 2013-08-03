Attribute VB_Name = "LoadLog"
Sub Load_Log()
Attribute Load_Log.VB_Description = "Macro recorded 14/05/2008 by Administrator"
Attribute Load_Log.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Load_Log Macro
' Macro recorded 14/05/2008 by Administrator
'

'
    ChDir "R:\Testing\Logging"
    Workbooks.Open Filename:="R:\testing\logging\log.csv"
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "m/d/yyyy h:mm"
    Columns("A:A").EntireColumn.AutoFit
    Cells.Replace What:="mA", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="V", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
   Charts.Add
    ActiveChart.ChartType = xlXYScatterLines
    ActiveChart.SetSourceData Source:=Sheets("log").Range("A1:D19")
    ActiveChart.Location Where:=xlLocationAsObject, name:="log"
 
 With ActiveChart
        .DisplayBlanksAs = xlInterpolated
        .PlotVisibleOnly = True
        .SizeWithWindow = True
    End With
    Application.ShowChartTipNames = True
    Application.ShowChartTipValues = True
End Sub


