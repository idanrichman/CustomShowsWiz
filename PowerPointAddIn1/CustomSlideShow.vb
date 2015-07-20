Public Class CustomSlideShow

    Private _count As Integer
    Private _slides As List(Of PowerPoint.Slide)
    Private _slidesIDs As Object
    Private _namedshow As PowerPoint.NamedSlideShow

    Public ReadOnly Property Count() As Integer
        Get
            Return _count
        End Get
    End Property

    Public Sub New()
    End Sub

    Public Sub New(namedShow As PowerPoint.NamedSlideShow)
        Init(namedShow)
    End Sub

    Public Sub Init(namedShow As PowerPoint.NamedSlideShow)
        _count = namedShow.Count
        _namedshow = namedShow
        _slidesIDs = namedShow.SlideIDs
        GetSlides()
    End Sub

    Private Sub GetSlides() '' Puts all slides from the namedshow to the _slides array
        Try
            Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application
            Dim oPresentation As PowerPoint.Presentation = oPowerpoint.ActivePresentation

            If _slidesIDs IsNot Nothing Then
                For i As Integer = 1 To _count
                    _slides.Add(oPresentation.Slides.FindBySlideID(CType(_slidesIDs(i), Integer)))
                Next
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function IsSlideInside(sld As PowerPoint.Slide) As Boolean
        Dim sldID As Integer

        For Each sldID In _slidesIDs
            If sldID = sld.SlideID Then Return True
        Next

        Return False
    End Function
End Class
