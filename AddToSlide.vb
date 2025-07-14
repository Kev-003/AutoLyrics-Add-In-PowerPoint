Imports Microsoft.Office.Core
Public Class AddToSlide
    ' if "lyrics" layout exists in slide master, create that slide
    Public Function AddLyricsToSlide(lyrics As List(Of String)) As String
        Dim slideLayoutName As String = "Lyrics"
        Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim presentation As PowerPoint.Presentation = pptApp.ActivePresentation
        Dim slideMaster As PowerPoint.Master = presentation.Designs(1).SlideMaster
        Dim slideLayout As PowerPoint.CustomLayout = Nothing

        ' Find the custom layout with the name "lyrics"
        For Each layout As PowerPoint.CustomLayout In slideMaster.CustomLayouts
            If layout.Name.Equals(slideLayoutName, StringComparison.OrdinalIgnoreCase) Then
                slideLayout = layout
                Exit For
            End If
        Next

        If slideLayout Is Nothing Then
            Return "No 'lyrics' layout found in the slide master."
        End If

        ' Add a new slide with the "lyrics" layout
        For Each block As String In lyrics
            Dim newSlide As PowerPoint.Slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, slideLayout)

            ' Try multiple methods to find the lyrics text box
            Dim textBox As PowerPoint.Shape = FindLyricsTextBox(newSlide, slideLayout)

            If textBox IsNot Nothing Then
                textBox.TextFrame.TextRange.Text = block.Trim()
            Else
                ' Fallback: add a textbox if named one not found
                Dim fallback As PowerPoint.Shape = newSlide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    100, 100, 500, 300
                )
                fallback.TextFrame.TextRange.Text = block.Trim()
            End If
        Next

        Return "Lyrics added to the slides successfully."
    End Function

    Private Function FindLyricsTextBox(slide As PowerPoint.Slide, layout As PowerPoint.CustomLayout) As PowerPoint.Shape
        ' Find text placeholders that are NOT title or footer
        For Each shape As PowerPoint.Shape In slide.Shapes
            If shape.Type = MsoShapeType.msoPlaceholder AndAlso shape.HasTextFrame Then
                Dim placeholderType As PowerPoint.PpPlaceholderType = shape.PlaceholderFormat.Type

                ' Check if it's NOT a title or footer placeholder
                If placeholderType <> PowerPoint.PpPlaceholderType.ppPlaceholderTitle AndAlso
                   placeholderType <> PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle AndAlso
                   placeholderType <> PowerPoint.PpPlaceholderType.ppPlaceholderFooter AndAlso
                   placeholderType <> PowerPoint.PpPlaceholderType.ppPlaceholderHeader AndAlso
                   placeholderType <> PowerPoint.PpPlaceholderType.ppPlaceholderSlideNumber AndAlso
                   placeholderType <> PowerPoint.PpPlaceholderType.ppPlaceholderDate Then
                    Return shape
                End If
            End If
        Next

        ' Fallback: try to find by exact name (in case it's a regular text box)
        For Each shape As PowerPoint.Shape In slide.Shapes
            If shape.Name.Equals("!!Lyrics", StringComparison.OrdinalIgnoreCase) Then
                Return shape
            End If
        Next

        ' Final fallback: find any text box that's not a placeholder
        For Each shape As PowerPoint.Shape In slide.Shapes
            If shape.Type = MsoShapeType.msoTextBox AndAlso shape.HasTextFrame Then
                Return shape
            End If
        Next

        Return Nothing
    End Function
End Class