Function shapeExists(shapeName As String, slideIdx As Integer) As Boolean
'ashleedawg https://stackoverflow.com/questions/37179927/how-to-check-whether-any-shape-exists'
'returns TRUE if a shape named [ShapeName] exists on the active worksheet'
    Dim sh As Shape
    For Each sh In ActivePresentation.Slides.Range(slideIdx).Shapes
         If sh.Name = shapeName Then shapeExists = True
    Next sh
End Function

Sub addToc()

upperLimit = ActivePresentation.Slides.Count - 1

presTitle = ActivePresentation.Slides.Range(1).Shapes("Title 1").TextFrame.TextRange.Text

For ii = 2 To upperLimit
    'delete ToC text box, if it already exists'
    If shapeExists("tracking_id_99", (ii)) Then
        ActivePresentation.Slides.Range(ii).Shapes("tracking_id_99").Delete
    End If
Next ii

For ii = 2 To upperLimit
    Set newTextBox = ActivePresentation.Slides.Range(ii).Shapes.AddTextbox(msoTextOrientationHorizontal, _
        Left:=7, Top:=5, Width:=200, Height:=10)
    newTextBox.Name = "tracking_id_99"
    With newTextBox.TextFrame
        On Error GoTo errHandler
            .TextRange.Text = presTitle & "  >  " & ActivePresentation.SectionProperties.Name(ActivePresentation.Slides.Range(ii).sectionIndex)
        .TextRange.Font.Size = 7
        .TextRange.Font.Name = "candara"
        .TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
        .MarginLeft = 0
        .MarginTop = 0
    End With
Next ii

exitSub:
    Exit Sub

'Print only presentation title, if slide belongs to no section'
errHandler:
    errMsg = "No section found: Only presentation title added to the headline."
    MsgBox errMsg
    With newTextBox.TextFrame
        .TextRange.Text = presTitle
        .TextRange.Font.Size = 7
        .TextRange.Font.Name = "candara"
        .TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
        .MarginLeft = 0
        .MarginTop = 0
    End With
    Resume Next

End Sub
