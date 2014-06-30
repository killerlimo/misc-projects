Attribute VB_Name = "UpdateFields"
Sub UpdateFields()
' Update all fields in the active document

    Dim oStory As Range
    For Each oStory In ActiveDocument.StoryRanges
        oStory.Fields.Update
        If oStory.StoryType <> wdMainTextStory Then
            While Not (oStory.NextStoryRange Is Nothing)
                Set oStory = oStory.NextStoryRange
                oStory.Fields.Update
            Wend
        End If
    Next oStory
    Set oStory = Nothing
End Sub
