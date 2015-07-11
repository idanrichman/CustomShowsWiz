Public Class SwapSlidesForm

    Private Sub SwapSlidesForm_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim SldRng As PowerPoint.SlideRange
        Dim SLabelTitle As String

        SldRng = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange
        SLabelTitle = Label_Title.Text
        SLabelTitle = SLabelTitle.Replace("XXXX", SldRng(2).SlideIndex.ToString)
        SLabelTitle = SLabelTitle.Replace("YYYY", SldRng(1).SlideIndex.ToString)
        Label_Title.Text = SLabelTitle

        '' Pasting Old Slide to be replaced
        Dim grabpicture As System.Drawing.Image
        SldRng(2).Copy()
        grabpicture = My.Computer.Clipboard.GetImage()
        PictureBox_OldSld.Image = grabpicture
        PictureBox_OldSld.Tag = SldRng(2)

        '' Pasting New Slide to replace with
        SldRng(1).Copy()
        grabpicture = My.Computer.Clipboard.GetImage()
        PictureBox_NewSld.Image = grabpicture
        PictureBox_NewSld.Tag = SldRng(1)

        If LocalizeM.isRTL Then
            PictureBox_Arrow.Image.RotateFlip(Drawing.RotateFlipType.RotateNoneFlipX)
        End If

    End Sub

    Private Sub ButtonOK_Click(sender As Object, e As EventArgs) Handles ButtonOK.Click
        Dim OldSld As PowerPoint.Slide
        Dim NewSld As PowerPoint.Slide

        OldSld = CType(PictureBox_OldSld.Tag, PowerPoint.Slide)
        NewSld = CType(PictureBox_NewSld.Tag, PowerPoint.Slide)

        SwapSlides_AllCustomShows(OldSld, NewSld)
        Close()
    End Sub

    Private Sub PictureBox_NewSld_Click(sender As Object, e As EventArgs) Handles PictureBox_NewSld.Click

    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub PictureBox_Arrow_Click(sender As Object, e As EventArgs) Handles PictureBox_Arrow.Click

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        '' Localize (RTL in case of hebrew UI) the form and all it's controls
        LocalizeM.Localize_Forms(Me)
    End Sub
End Class