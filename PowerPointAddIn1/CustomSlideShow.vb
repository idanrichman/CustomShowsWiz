Public Class CustomSlideShow

    Private _count As Integer
    Private _slides As List(Of PowerPoint.Slide)
    Private _slidesIDs() As Integer
    Private _namedshow As PowerPoint.NamedSlideShow



    Private newPropertyValue As String
    Public Property NewProperty() As String
        Get
            Return newPropertyValue
        End Get
        Set(ByVal value As String)
            newPropertyValue = value
        End Set
    End Property

End Class
