Imports Microsoft.Office.Interop
Imports System.Globalization
Imports System.Deployment
Imports System.Windows.Forms




Public Class ThisAddIn
    Private Const WHATS_NEW As Boolean = False '' only if TRUE then the published version should show the whats new box.

    Private myUserControl1 As MyUserControl
    Private myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public customRibbon As Office.IRibbonUI
    Public Shared OfficeLanguage As Integer
    Public Shared sCurrentVersion As String = vbNullString

   
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ''      can be done in order to test the application in different languages.
        ' Threading.Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")
       
        Try
            '' Checking if the application is running from VS2013 or deployed, then checking if it has been just updated (.IsFirstRun) - if so then show whats new form
            If Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
                With Deployment.Application.ApplicationDeployment.CurrentDeployment
                    If .IsFirstRun Then
                        If My.Settings.FirstRun = True Then ''Settings.FirstRun is registery value indicating if it's the add-in very first run
                            AddHandler Application.WindowActivate, AddressOf EWindowActivateJustInstalled_Event
                            My.Settings.FirstRun = False
                            My.Settings.Save()
                        Else
                            If WHATS_NEW Then
                                ''Adding event handler to show the Whats New Form when the window is activated and the application is running. otherwise, the form would draw before the application started running
                                AddHandler Application.WindowActivate, AddressOf EWindowActivateWhatsNew_Event
                            End If
                        End If
                        sCurrentVersion = .CurrentVersion.ToString()
                    End If
                End With
            End If

            '''' EVENTUALLY I'M NOT USING IT. INSTEAD I'M USING THE .FIRSTRUN METHOD
            '' if the registery version value is different from the application version (hence, an update has been made)
            '' then pop up the Whats new window (if defined for it)
            'If Not My.Settings.Version = sCurrentVersion Then
            '    MsgBox("Updated to version: " & sCurrentVersion)
            '    '' updating the registery version value
            '    My.Settings.Version = sCurrentVersion
            '    My.Settings.Save()
            'End If
            '''' -------------------------------------------------------------------

            'myUserControl1 = New MyUserControl
            'myCustomTaskPane = Me.CustomTaskPanes.Add(myUserControl1, "My Task Pane")
            'myCustomTaskPane.Visible = True

            '  Load_Context_Menus() '' for now it's not needed as i used dynamic menu

            '' Get office UI language and localize strings if in hebrew
            OfficeLanguage = Me.Application.LanguageSettings.LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI)
            LocalizeM.InitializeStrings()

            '' Adding the event handler to the event when the selection is changeing so the ribbon will be updated
            AddHandler Application.WindowSelectionChange, AddressOf EWindowSelectionChange_Event

        Catch e As Exception
            MsgBox("An unexpected error occured while loading addin, some functions may not work", CType(MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, MsgBoxStyle), "Custom Shows Wiz Addin")
        End Try

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        '    Unload_Context_Menus()
    End Sub
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

#Region "Events handlers"

    Private Sub EWindowSelectionChange_Event(Sel As PowerPoint.Selection)
        InvalidateRibbonCustomShows()
    End Sub

    Private Sub EWindowActivateJustInstalled_Event(ByVal Pres As PowerPoint.Presentation, ByVal Wn As PowerPoint.DocumentWindow)
        Dim fFirstInstallForm As New FirstInstallForm
        fFirstInstallForm.ShowDialog()
        RemoveHandler Application.WindowActivate, AddressOf EWindowActivateJustInstalled_Event
    End Sub

    Private Sub EWindowActivateWhatsNew_Event(ByVal Pres As PowerPoint.Presentation, ByVal Wn As PowerPoint.DocumentWindow)
        Dim fWhatsNewForm As New WhatsNewForm
        fWhatsNewForm.ShowDialog()
        RemoveHandler Application.WindowActivate, AddressOf EWindowActivateWhatsNew_Event
    End Sub
#End Region

#Region "Various routines"

    Public Sub InvalidateRibbonCustomShows()
        '' Making sure that the ribbon is updated.
        customRibbon.InvalidateControl("CustomSlideShowCustomMenu")
    End Sub
#End Region
End Class
