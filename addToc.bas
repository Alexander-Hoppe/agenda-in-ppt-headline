Sub addToc()

upperLimit = ActivePresentation.Slides.Count - 1

presTitle = ActivePresentation.Slides.Range(1).Shapes("Title 1").TextFrame.TextRange.Text

For ii = 2 To upperLimit
    For Each mySlide In ActivePresentation.Slides.Range(ii)
        Set newTextBox = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            Left:=7, Top:=5, Width:=200, Height:=10)
        With newTextBox.TextFrame
            .TextRange.Text = presTitle & "  >  " & ActivePresentation.SectionProperties.Name(mySlide.sectionIndex)
            .TextRange.Font.Size = 7
            .TextRange.Font.Name = "candara"
            .MarginLeft = 0
            .MarginTop = 0
            .TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
        End With
    Next
Next ii
    
End Sub
