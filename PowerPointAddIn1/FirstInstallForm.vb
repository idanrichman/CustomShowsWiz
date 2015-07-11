Public Class FirstInstallForm
    Private Structure Feature
        Dim image As System.Drawing.Bitmap
        Dim desc As String
        Dim title As String
    End Structure

    '    Dim myFeatures(0 To 1) As Feature
    Dim myFeatures As New List(Of Feature)
    Dim iFeatureNum As Integer = 0

    Private Sub UpdateControls()
        PictureBox1.Image = myFeatures(iFeatureNum).image
        Desc_Label.Text = myFeatures(iFeatureNum).desc
        Featuring_Label.Text = LocalizeM.String_FirstInstall_Featuring + myFeatures(iFeatureNum).title
    End Sub

    Private Sub AddFeature(image As Drawing.Bitmap, desc As String, title As String)
        Dim oFeature As Feature
        oFeature.image = image
        oFeature.desc = desc
        oFeature.title = title
        myFeatures.Add(oFeature)
    End Sub

    Private Sub FirstInstallForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        '' Defenitions for the images and descriptions.
        'myFeatures(0).image = My.Resources.CreateNewPng
        'myFeatures(0).desc = LocalizeM.String_FirstInstall_Feature0_Desc
        'myFeatures(0).title = LocalizeM.String_FirstInstall_Feature0_Title
        'myFeatures(1).image = My.Resources.FindEmptyPng
        'myFeatures(1).desc = LocalizeM.String_FirstInstall_Feature1_Desc
        'myFeatures(1).title = LocalizeM.String_FirstInstall_Feature1_Title
        AddFeature(image:=My.Resources.CreateNewPng, desc:=LocalizeM.String_FirstInstall_Feature0_Desc, title:=LocalizeM.String_FirstInstall_Feature0_Title)
        AddFeature(image:=My.Resources.FindEmptyPng, desc:=LocalizeM.String_FirstInstall_Feature1_Desc, title:=LocalizeM.String_FirstInstall_Feature1_Title)
        AddFeature(image:=My.Resources.SwapSlidesPng, desc:=My.Resources.FirstInstall_Feature_SwapSlides_Desc, title:=My.Resources.FirstInstall_Feature_SwapSlides_Title)

        ''Changing title according to localization.
        Label1.Text = LocalizeM.String_FirstInstall_MainMsg
        Me.ButtonOK.Text = LocalizeM.String_Button_OK
        Me.Button_Next.Text = LocalizeM.String_Button_Next

        '' Loading the current feature
        UpdateControls()
    End Sub

    Private Sub ButtonOK_Click(sender As Object, e As EventArgs) Handles ButtonOK.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Button_Next_Click(sender As Object, e As EventArgs) Handles Button_Next.Click
        iFeatureNum += 1
        If iFeatureNum = myFeatures.Count - 1 Then
            Me.Button_Next.Visible = False
            Me.ButtonOK.Visible = True
        End If

        UpdateControls()

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        '' Localize (RTL in case of hebrew UI) the form and all it's controls
        LocalizeM.Localize_Forms(Me)
    End Sub
End Class