Public Class WhatsNewForm

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        LocalizeM.Localize_Forms(Me)

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub WhatsNewForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.LabelVersion.Text = "Version: " & ThisAddIn.sCurrentVersion
        Me.TextBox1.Text = "Whats new?" + vbNewLine +
                           "------------------" + vbNewLine +
                           "* New feature: Swap slides in all custom shows" + vbNewLine +
                           "* Bug fixes"

        If ThisAddIn.OfficeLanguage = Office.MsoLanguageID.msoLanguageIDHebrew Then
            Me.Text = "מה חדש?"
            Me.LabelVersion.Text = "גרסא: " & ThisAddIn.sCurrentVersion
            Me.LabelMsg.Text = "ברכות! תוסף Custom Shows Wiz בדיוק התעדכן."
            Me.ButtonOK.Text = LocalizeM.String_Button_OK
            Me.TextBox1.Text = "מה חדש?" + vbNewLine +
                               "---------------" + vbNewLine +
                               "* פונקציה חדשה: החלפת שקף ישן בשקף חדש בכל התצוגות האישיות" + vbNewLine +
                               "* תיקוני באגים"


        End If

        Me.TextBox1.SelectionStart() = 0
    End Sub

    Private Sub ButtonOK_Click(sender As Object, e As EventArgs) Handles ButtonOK.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class