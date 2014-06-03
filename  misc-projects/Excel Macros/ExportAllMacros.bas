Attribute VB_Name = "ExportAllMacros"
Sub ExportAllMacros()
     
     ' reference to extensibility library
     ' Export all .bas files to c:\temp in one go
     
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
     
    Set objMyProj = Application.VBE.ActiveVBProject
     
    For Each objVBComp In objMyProj.VBComponents
        If objVBComp.Type = vbext_ct_StdModule Then
            objVBComp.Export "C:\temp\" & objVBComp.name & ".bas"
        End If
    Next
     
End Sub
