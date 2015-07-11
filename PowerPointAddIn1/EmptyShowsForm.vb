Imports System.Windows.Forms

Public Class EmptyShowsForm
    Private isAllChecked As Boolean = False


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        '' Localize (RTL in case of hebrew UI) the form and all it's controls
        LocalizeM.Localize_Forms(Me)

    End Sub

    Private Sub EmptyShowsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim nCustomShow As PowerPoint.NamedSlideShow
        Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim iShowNum As Integer
        Dim iEmptyShownum As Integer = 0
        Dim sBufferString As String

        Me.Text = LocalizeM.String_Form_Emptyshows_Title
        Me.Button1.Text = String_Button_SelectAll
        Me.Button2.Text = String_Button_Delete
        Me.Button3.Text = String_Button_Cancel
        Me.Label1.Text = String_Form_Emptyshows_label1

        '' For Debuging purposes only
        'Me.CheckedListBox1.Items.Add("Test1")
        'Me.CheckedListBox1.Items.Add("Test2")
        'Me.CheckedListBox1.Items.Add("Test3")

        '' Populating the checkedlistbox with empty shows names
        For Each nCustomShow In oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows
            If nCustomShow.Count = 0 Then
                Me.CheckedListBox1.Items.Add(nCustomShow.Name)
                iEmptyShownum += 1
            End If
        Next

        '' Updating label2 text
        iShowNum = oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows.Count
        sBufferString = String_Form_Emptyshows_label2.Replace("XXXX", iEmptyShownum.ToString)
        sBufferString = sBufferString.Replace("YYYY", iShowNum.ToString)
        Me.Label2.Text = sBufferString

    End Sub

    Private Sub EmptyShowsForm_Close()
        Me.Close()
        Me.Dispose()

        Globals.ThisAddIn.InvalidateRibbonCustomShows() '' To update the ribbon list in case changed has been made to the custom shows list
    End Sub

    '' Button 1 = Select All
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i As Integer

        If isAllChecked = False Then
            For i = 0 To Me.CheckedListBox1.Items.Count - 1
                Me.CheckedListBox1.SetItemChecked(i, True)
                isAllChecked = True
            Next
        Else
            For i = 0 To Me.CheckedListBox1.Items.Count - 1
                Me.CheckedListBox1.SetItemChecked(i, False)
                isAllChecked = False
            Next
        End If

    End Sub

    '' Button 3 = Cancel button
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        EmptyShowsForm_Close()
    End Sub

    '' Button 2 = Delete button
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim iDeleteShowNum As Integer
        Dim iDeletedShowNum As Integer = 0
        Dim i As Integer = 1
        Dim iUsrInput As Windows.Forms.DialogResult
        Dim oCheckedItem As String
        Dim sBufferString As String

        iDeleteShowNum = Me.CheckedListBox1.CheckedItems.Count
        If iDeleteShowNum > 0 Then
            iUsrInput = Windows.Forms.MessageBox.Show(LocalizeM.String_DeleteCustomShows_Msg.Replace("XXXX", iDeleteShowNum.ToString),
                                                      LocalizeM.String_DeleteCustomShows_Title,
                                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                      LocalizeM.RTL_Setting, False)

            Select Case iUsrInput
                Case Windows.Forms.DialogResult.Yes
                    For Each oCheckedItem In Me.CheckedListBox1.CheckedItems
                        Try
                            '' Deleting the selected custom shows
                            oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows(oCheckedItem).Delete()
                            iDeletedShowNum += 1
                        Catch ex As Exception
                            '' prompting the user that a specific customshow could not be deleted for any reason
                            Windows.Forms.MessageBox.Show(String_DeleteCustomShows_Err.Replace("XXXX", oCheckedItem),
                                            LocalizeM.String_DeleteCustomShows_Title,
                                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation,
                                            MessageBoxDefaultButton.Button1, LocalizeM.RTL_Setting, False)
                        End Try
                    Next

                    '' Prompting the user how many custom shows were successfully deleted out the selected ones
                    sBufferString = String_DeleteCustomShows_Done_Msg.Replace("XXXX", iDeletedShowNum.ToString)
                    sBufferString = sBufferString.Replace("YYYY", iDeleteShowNum.ToString)

                    Windows.Forms.MessageBox.Show(sBufferString,
                                            LocalizeM.String_DeleteCustomShows_Title,
                                            MessageBoxButtons.OK, MessageBoxIcon.Information,
                                            MessageBoxDefaultButton.Button1, LocalizeM.RTL_Setting, False)
                    EmptyShowsForm_Close()
            End Select
        Else
            '' Prompting the user that no custom show was checked
            Windows.Forms.MessageBox.Show(LocalizeM.String_DeleteCustomShows_NoChecked_Msg,
                                            LocalizeM.String_DeleteCustomShows_Title,
                                            MessageBoxButtons.OK, MessageBoxIcon.Information,
                                            MessageBoxDefaultButton.Button1, LocalizeM.RTL_Setting, False)
        End If
    End Sub

End Class