Module Module1
    '   Sub Check_Shape_CustomShow()
    '       Dim NS As NamedSlideShow
    '       Dim i As Long
    '       Dim S As String
    '       Dim msg As String

    '       With ActivePresentation.SlideShowSettings
    '           For Each NS In .NamedSlideShows
    '               If NS.Name = ActiveWindow.Selection.ShapeRange(1).ActionSettings(ppMouseClick).Run Then
    '                   msg = NS.Name & ":" & vbNewLine
    '                   For i = 1 To NS.count
    '                       msg = msg & ActivePresentation.Slides.FindBySlideID(CStr(NS.SlideIDs(i))).SlideIndex & ": " & _
    '                             ActivePresentation.Slides.FindBySlideID(CStr(NS.SlideIDs(i))).Shapes.Title.TextFrame.TextRange.Text & vbCrLf
    '                       'MsgBox CStr(NS.SlideIDs(I)) & vbNewLine
    '                   Next
    '                   MsgBox msg
    '               End If
    '           Next
    '       End With
    '   End Sub
    '   Sub Check_if_slide_CustomShow()
    '       Dim NS As NamedSlideShow
    '       Dim i As Long
    '       Dim S As String
    '       Dim msg As String
    '       Dim sldid As Integer
    '       Dim sld As Slide

    '       With ActiveWindow
    '           If .View.Type = ppViewNotesPage Then
    '               .ViewType = ppViewSlide
    '           End If
    '       End With

    '       sld = Application.ActiveWindow.View.Slide
    '       sldid = sld.SlideID
    '       msg = ""

    '       With ActivePresentation.SlideShowSettings
    '           For Each NS In .NamedSlideShows
    '               For i = 1 To NS.count
    '                   If NS.SlideIDs(i) = sldid Then msg = msg & NS.Name & vbCrLf
    '               Next
    '           Next

    '           If Not msg = "" Then
    '               MsgBox("This slide exists in these custom shows:" & vbCrLf & vbCrLf & msg, vbInformation, "Custom Shows List")
    '           Else
    '               MsgBox("This slide doesn't appear in any custom show", vbInformation, "Custom Shows List")
    '           End If

    '       End With

    '   End Sub

    '   Private Function Check_slide_CustomShow(sld As Slide) As String
    '       Dim NS As NamedSlideShow
    '       Dim i As Long
    '       Dim count As Long
    '       Dim S As String
    '       Dim msg As String
    '       Dim sldid As Integer

    '       sldid = sld.SlideID
    '       msg = ""
    '       count = 0

    '       With ActivePresentation.SlideShowSettings

    '           For Each NS In .NamedSlideShows
    '               For i = 1 To NS.count
    '                   If NS.SlideIDs(i) = sldid Then
    '                       If count = 0 Then
    '                           msg = msg & NS.Name
    '                       Else
    '                           msg = msg & ", " & NS.Name
    '                       End If
    '                       count = 1
    '                   End If
    '               Next
    '           Next

    '           Check_slide_CustomShow = msg
    '           '         If Not Msg = "" Then
    '           '           Check_slide_CustomShow = "This slide exists in these custom shows:" & vbCrLf & vbCrLf & Msg
    '           '          Else
    '           '           Check_slide_CustomShow = "This slide doesn't appear in any custom show"
    '           '          End If

    '       End With

    '   End Function

    '   Sub Slides_CustomShow_Report()

    '       Dim sld As Slide
    '       Dim msg As String

    '       Dim i As Integer

    '       msg = ""

    '       For Each sld In ActivePresentation.Slides
    '           msg = msg & sld.SlideIndex & ": " & Check_slide_CustomShow(sld) & vbCrLf
    '       Next

    '       CustomShowsRepForm.TextBox1.Text = msg
    '       CustomShowsRepForm.TextBox1.SelStart = 0
    '       CustomShowsRepForm.Show()

    '   End Sub


    '   Sub Delete_Empty_CustomShows()
    '       Dim NS As NamedSlideShow
    '       Dim msg As String
    '       Dim emptyCount As Long

    '       emptyCount = 0
    '       With ActivePresentation.SlideShowSettings
    '           For Each NS In .NamedSlideShows
    '               If NS.count = 0 Then
    '                   msg = msg & NS.Name & vbCrLf
    '                   emptyCount = emptyCount + 1
    '               End If
    '           Next
    '       End With

    '       msg = "Empty Custom Slideshows:" & vbCrLf & _
    '               "Total: " & emptyCount & " Out of " & ActivePresentation.SlideShowSettings.NamedSlideShows.count & vbCrLf & _
    '               "------------------------" & vbCrLf _
    '               & msg

    '       EmptySlidesForm.TextBox1.Text = msg
    '       EmptySlidesForm.TextBox1.SelStart = 0
    '       EmptySlidesForm.Show()


    '   End Sub

    '   Sub Jump_To_Hyperlink()
    '       Dim sldNumber As Long

    'Debug.Print ActivePresentation.SlideShowSettings.NamedSlideShows/..ActiveWindow.Selection.ShapeRange(1).ActionSettings(ppMouseClick).Run

    '   End Sub


    'Private Sub CommandButton2_Click()
    '    Dim i As Long
    '    Dim deletedCount As Long
    '    Dim NS As NamedSlideShow

    '    deletedCount = 0
    '    With ActivePresentation.SlideShowSettings
    '        For i = 1 To .NamedSlideShows.count ' Cycling through all named slide shows
    '            If .NamedSlideShows(i).count = 0 Then 'Checking if namedslideshow is empty
    '                .NamedSlideShows(i).Delete()
    '                i = i - 1 'Deletion casued all namedslideshows indexs to decrease by 1 so this line makes sure the for loop doesn't miss any namedslideshow
    '                deletedCount = deletedCount + 1
    '            End If
    '            If i = .NamedSlideShows.count Then Exit For 'A Patch, for some reason the for loop kept running with an i bigger than the namedslideshow.count
    '        Next
    '    End With

    '    MsgBox("Delete was successful. " & deletedCount & " Custom-Slide-Shows were deleted")

    '    Unload UserForm2
    'End Sub



End Module
