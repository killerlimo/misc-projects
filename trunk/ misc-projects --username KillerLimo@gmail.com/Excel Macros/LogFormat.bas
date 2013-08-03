Attribute VB_Name = "LogFormat"
Sub Log_Format()
Attribute Log_Format.VB_Description = "Formats loaded csv file."
Attribute Log_Format.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Log_Format Macro
' Formats loaded csv file.
'
'
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "m/d/yyyy h:mm:ss"
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
