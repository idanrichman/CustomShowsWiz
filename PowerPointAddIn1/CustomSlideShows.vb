Public Class CustomSlideShows

    Structure SlideCustomShowMeta
        Dim RefersTo As List(Of Integer) ''all customshows links within the slide
        Dim ContainedIn As List(Of String) '' naming all the custom shows that contain this slide
        Dim SlideID As Integer
    End Structure

    Private _count As Integer
    Private _namedshows As PowerPoint.NamedSlideShows
    Private _slides As List(Of SlideCustomShowMeta) '' Holding all of the presentation slides
    Private _customshows As List(Of CustomSlideShow) '' Holding all of the presenation customshows

    Public ReadOnly Property Count() As Integer
        Get
            Return _count
        End Get
    End Property

    Public Sub New()
    End Sub

    Public Sub New(namedshows As PowerPoint.NamedSlideShows)
        _count = namedshows.Count
        _namedshows = namedshows
    End Sub

    Public Sub Init(presentation As PowerPoint.Presentation)

    End Sub

End Class
