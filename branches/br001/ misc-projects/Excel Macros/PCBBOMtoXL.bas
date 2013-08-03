Attribute VB_Name = "PCBBOMtoXL"
Sub PCBBOMtoXL()
' Attribute PCBBOMtoXL.VB_Description = "Macro recorded 16/10/2008 by Administrator"
' Attribute PCBBOMtoXL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PCB Proteus (Isis) BOM format for Excel Parts List Macro
' Macro recorded 16/10/2008 by Administrator
'

'
    Columns("A:F").Select
    Columns("A:F").EntireColumn.AutoFit
    Columns("A:B").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    
    Cells.Select
    Selection.Sort Key1:=Range("F1"), Order1:=xlAscending, Key2:=Range("E1") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2 _
        :=xlSortNormal
    
    Columns("E:E").Select
    Selection.Replace What:="-", Replacement:="--", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=",", Replacement:=", ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("E:E").EntireColumn.AutoFit
    
    Range("A1").Select
End Sub


