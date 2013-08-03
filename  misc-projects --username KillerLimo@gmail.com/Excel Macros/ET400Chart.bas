Attribute VB_Name = "ET400Chart"
Sub ET400_Chart()
Attribute ET400_Chart.VB_Description = "Convert RX key log data into chart."
Attribute ET400_Chart.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ET400_Chart Macro
' Convert RX key log data into chart.
'

'
    ActiveCell.SpecialCells(xlLastCell).Select
    Range(Selection, Cells(1)).Select
    ActiveSheet.Shapes.AddChart.Select
    'ActiveChart.SetSourceData Source:=Range( _
        "'ET400RX_476951_keylog_20110804'!$A$1:$E$197")
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    ActiveChart.PlotArea.Select
    ActiveChart.DisplayBlanksAs = xlInterpolated
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).Select
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).AxisGroup = 2
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).Select
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.Axes(xlCategory).Select
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.ChartArea.Select
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(2).Select
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.PlotArea.Select
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(1).Select
'   ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.SeriesCollection(3).Select
'    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 10
    
End Sub

