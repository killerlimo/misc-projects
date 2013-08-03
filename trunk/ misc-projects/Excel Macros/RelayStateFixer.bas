Attribute VB_Name = "RelayStateFixer"
Sub ChartOnly()
Attribute ChartOnly.VB_ProcData.VB_Invoke_Func = "q\n14"
    'Use this macro for data that has already had additional relay states added. i.e. version 2 Key Manager
    
    ' Format time column to include seconds
    Columns("A:A").EntireColumn.AutoFit
    Columns("A:A").Select
    Selection.NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
    Call RemoveDuplicates
    Call AddMissingTimes
'    Call Sort              Do not use as it causes problems with the relay missing states.
    Call SelectAllData
    Call AddChart
    Call FormatChart
    
End Sub
Sub RelayStateFixer()
' Use this macro with data that is missing the additional relay data points. i.e. version 1 of Key Manager.

With Application
 
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
  
        ' Format time column to include seconds
        Columns("A:A").EntireColumn.AutoFit
        Columns("A:A").Select
        Selection.NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
        Call FindData
 '       LastLine = SelectAllData
        Call AddChart
        
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    
    Call FormatChart
 
End With
 
End Sub
Sub Sort()
'
' Sort Macro
'

'
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("A2:A9999") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A2:F9999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub FindData()
'   Starts at the third row (ignore header) and moves downwards looking for a non-blank relay entry
'   On finding this it inserts a blank row, copies the time stamp from the row above and the relay value from the next value down the list
 
    Dim i As Variant
 
    i = 3
 
    Do
        ' Check for duplicate rows and delete
        If Cells(i, 1) = Cells(i + 1, 1) And Cells(i, 2) = Cells(i + 1, 2) And Cells(i, 3) = Cells(i + 1, 3) And Cells(i, 4) = Cells(i + 1, 4) And Cells(i, 5) = Cells(i + 1, 5) Then
            Cells(i, 1).Select
            Selection.EntireRow.Delete
        End If
        
        RelayValue = Cells(i - 1, 2)
        
        If RelayValue <> "" Then
            Cells(i, 2).Select
            Call InsertBlankRow
            Cells(i, 2) = NextRelayValue(i)
            i = i + 1
        End If
        i = i + 1
    Loop Until Cells(i, 1) = ""
        
End Sub
Sub RemoveDuplicates()
' Check for duplicate rows and delete
    i = 3
    Do
        If Cells(i, 1) = Cells(i + 1, 1) And Cells(i, 2) = Cells(i + 1, 2) And Cells(i, 3) = Cells(i + 1, 3) And Cells(i, 4) = Cells(i + 1, 4) And Cells(i, 5) = Cells(i + 1, 5) Then
            Cells(i, 1).Select
            Selection.EntireRow.Delete
        End If
        i = i + 1
    Loop Until Cells(i, 1) = "" And Cells(i, 2) = "" And Cells(i, 3) = "" And Cells(i, 4) = "" And Cells(i, 5) = ""
        
End Sub
Sub AddMissingTimes()
' Add missing time stamps
    i = 3
    Do
        If Cells(i, 1) = "" Then
            Cells(i, 1) = Cells(i - 1, 1)
        End If
        i = i + 1
    Loop Until Cells(i, 1) = "" And Cells(i, 2) = "" And Cells(i, 3) = "" And Cells(i, 4) = "" And Cells(i, 5) = ""
        
End Sub
Sub InsertBlankRow()
'   Creates a blank line above selected data.
'   Copies time stamp from row above into new line.

    Selection.EntireRow.Insert
    ActiveCell.Offset(-1, -1).Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
            
End Sub
Function NextRelayValue(k As Variant) As Variant
'   Search for next relay value and return it

    j = k
    Do
        j = j + 1
        If Cells(j, 2) <> "" Then NextRelayValue = Cells(j, 2)
    Loop Until Cells(j, 1) = "" Or Cells(j, 2) <> ""
 
End Function
Sub SelectAllData()
' Find last line in list and select very last cell

    l = 1
    Do
    l = l + 1
    Loop Until Cells(l, 1) = ""
    Cells(l - 1, 6).Select
    
    ' Select entire list
    Range(Selection, Cells(1)).Select
    
End Sub
Sub AddChart()
    ' Add chart
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    ' ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet5"
End Sub
Sub FormatChart()
    With ActiveChart
        .DisplayBlanksAs = xlInterpolated
        .PlotVisibleOnly = True
        .SizeWithWindow = True
    End With
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).AxisGroup = 2
   
    With Selection.Border
        .Weight = xlThin
        .LineStyle = xlAutomatic
    End With
    With Selection
        .MarkerBackgroundColorIndex = xlAutomatic
        .MarkerForegroundColorIndex = xlAutomatic
        .MarkerStyle = xlNone
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.Axes(xlValue, xlPrimary).Select
    With ActiveChart.Axes(xlValue, xlPrimary)
        .MinimumScale = -10
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlValue, xlSecondary).Select
    With ActiveChart.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .MaximumScale = 30
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlValue).Select
    ActiveChart.PlotArea.Select
    With ActiveChart
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Current:mA"
        .Axes(xlValue, xlSecondary).HasTitle = True
        .Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = _
        "Relay State:1=picked"
        .DisplayBlanksAs = xlInterpolated
        .PlotVisibleOnly = True
        .SizeWithWindow = True
    
        .SeriesCollection(1).Format.Line.Weight = 1
        .SeriesCollection(2).Format.Line.Weight = 2
        .SeriesCollection(3).Format.Line.DashStyle = 4
        .SeriesCollection(3).Format.Line.Weight = 2
        .SeriesCollection(4).Format.Line.Weight = 2
        .SeriesCollection(5).MarkerStyle = 3
        .SeriesCollection(5).MarkerSize = 10
        .SeriesCollection(5).ApplyDataLabels
        .SeriesCollection(5).Format.Line.Visible = False
        
    End With
End Sub




Sub ResetRange()
    ActiveSheet.UsedRange
End Sub
