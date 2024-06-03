Sub addToc()

upperLimit = ActivePresentation.Slides.Count - 1

presTitle = ActivePresentation.Slides.Range(1).Shapes("Title 1").TextFrame.TextRange.Text

For ii = 2 To upperLimit
    ActivePresentation.Slides.Range(ii).Shapes("tracking_id_99").Delete
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

errHandler:
    errMsg = "No sections found: No sections were added to the headline."
    MsgBox errMsg
    GoTo exitSub

End Sub
