Attribute VB_Name = "RemoveTemplate"
Sub RemoveTemplate()
Attribute RemoveTemplate.VB_Description = "Remove a document's template to make opening faster."
Attribute RemoveTemplate.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.RemoveTemplate"
'
' RemoveTemplate Macro
' Remove a document's template to make opening faster.
'
    With ActiveDocument
        .UpdateStylesOnOpen = False
        .AttachedTemplate = ""
        .XMLSchemaReferences.AutomaticValidation = True
        .XMLSchemaReferences.AllowSaveAsXMLWithoutValidation = False
    End With
End Sub
