Imports System.Windows.Forms
Imports System.Xml

'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Public ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText(LocalizeM.Ribbon_XML_FileName)
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        Globals.ThisAddIn.customRibbon = Me.ribbon
    End Sub

    Public Sub ReplaceAppendShow_CallBack(control As Office.IRibbonControl)
        Dim iUsrInput As Windows.Forms.DialogResult

        iUsrInput = Windows.Forms.MessageBox.Show(LocalizeM.String_AppendReplace_Msg1 & """ " + control.Tag + """ " + LocalizeM.String_AppendReplace_Msg2,
                                      LocalizeM.String_AppendReplace_Msg_Title + " " + control.Tag,
                                      MessageBoxButtons.YesNoCancel,
                                      MessageBoxIcon.Question,
                                      MessageBoxDefaultButton.Button1,
                                      LocalizeM.RTL_Setting,
                                      False)
        Select Case iUsrInput
            Case Windows.Forms.DialogResult.Yes
                CustomShowModule.AppendShow(control.Tag)
            Case Windows.Forms.DialogResult.No
                CustomShowModule.ReplaceShow(control.Tag)
        End Select

    End Sub

    Public Sub CreateNewShow_CallBack(control As Office.IRibbonControl)

        CustomShowModule.CreateNewShow()
        '' Update the ribbon dynamic menu so it will show the new added custom show
        Globals.ThisAddIn.InvalidateRibbonCustomShows()

    End Sub

    Public Sub RunCustomShow_CallBack(control As Office.IRibbonControl)
        Try
            With Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings
                .RangeType = PowerPoint.PpSlideShowRangeType.ppShowNamedSlideShow
                .SlideShowName = control.Tag
                .Run()
            End With
        Catch ex As Exception

        Finally
            Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings _
             .RangeType = PowerPoint.PpSlideShowRangeType.ppShowSlideRange '' Returning to default settings. otherwise when pressing the slideshow button it will always run only the last namedslideshow shown
        End Try

    End Sub

    Public Sub FindEmptyShows_CallBack(control As Office.IRibbonControl)
        Dim fEmptyShowForm As New EmptyShowsForm

        fEmptyShowForm.ShowDialog() ''.ShowDialog is for showing in modal mode. .Show is for non-modal mode.
    End Sub

    Public Sub SwapSlides_CallBack(control As Office.IRibbonControl)
        SwapSlides_LaunchForm()
    End Sub

    Public Function SwapSlides_getEnabled(control As Office.IRibbonControl)

        Try
            Dim SldRng As PowerPoint.SlideRange

            SldRng = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange
            If SldRng.Count = 2 Then
                Return vbTrue
            Else
                Return vbFalse
            End If
        Catch
            Return vbFalse
        End Try
    End Function

    Public Sub Jump2Hyperlink_CallBack(control As Office.IRibbonControl)
        Jump2Hyperlink()
    End Sub

    Public Function Jump2Hyperlink_getVisible(control As Office.IRibbonControl)

        Try
            Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application
            Dim oWindow As PowerPoint.DocumentWindow = oPowerpoint.ActiveWindow
            Dim oSelection As PowerPoint.Selection = oWindow.Selection
            Dim oShape As PowerPoint.Shape
            Dim oHyper1 As PowerPoint.Hyperlink
            Dim oHyper2 As PowerPoint.Hyperlink


            oShape = oSelection.ShapeRange(1)
            oHyper1 = oShape.ActionSettings(PowerPoint.PpMouseActivation.ppMouseClick).Hyperlink
            oHyper2 = oShape.TextFrame.TextRange.ActionSettings(PowerPoint.PpMouseActivation.ppMouseClick).Hyperlink

            If Not IsNothing(oHyper1.SubAddress) Or Not IsNothing(oHyper2.SubAddress) Then
                Return vbTrue
            Else
                Return vbFalse
            End If
        Catch
            Return vbFalse
        End Try

    End Function


    Function GetContent_NewShow(control As Office.IRibbonControl) As String
        Dim xml As String
        Dim aShows() As PowerPoint.NamedSlideShow
        Dim i As Integer
        Dim sShowName As String
        Dim sIdNum As String

        aShows = GetAllCustomShows()

        xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" + _
              "<button id=""CreateNewShow_Button"" imageMso=""ListSetNumberingValue"" label=""" + LocalizeM.String_Context_CreateNewShow_Button + """ onAction=""CreateNewShow_CallBack""/>" + _
              "<menuSeparator id=""CreateNewShowPopupSeperator""/>"


        '' Adding a button for each custom show. the "security.securityelement.escape" method is for encoding the inputted string to xml style (eg. double quotes are &quot;  etc')
        For i = 1 To aShows.Length - 1
            sShowName = aShows(i).Name
            sShowName = Security.SecurityElement.Escape(sShowName)
            Dim sDisplayShowName As String = sShowName.Replace("&amp;", "&amp;&amp;")
            sIdNum = CStr(i)
            xml += "<button id=""ContextcustomshowButton" + sIdNum + """ label=""" + sDisplayShowName + """ onAction=""ReplaceAppendShow_CallBack"" tag=""" + sShowName + """/>"
        Next

        xml += "</menu>"

        Return xml
    End Function


    Function GetContent_RibbonCustomShow(control As Office.IRibbonControl) As String
        Dim xml As String
        Dim aShows() As PowerPoint.NamedSlideShow
        Dim i As Integer
        Dim sShowName As String
        Dim sIdNum As String

        '' Puts in aShows all the custom shows that exists in the slide
        aShows = GetAllCustomShows()

        xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" + _
              "<button id=""RibbonCreateNewShow_Button"" imageMso=""ListSetNumberingValue"" label=""" + LocalizeM.String_Context_CreateNewShow_Button + """ onAction=""CreateNewShow_CallBack""/>" + _
              "<button id=""RibbonFindEmpty_Button"" imageMso=""InkEraseMode"" label=""" + LocalizeM.String_RibbonCustomShows_FindEmpty_Button + """ onAction=""FindEmptyShows_CallBack""/>" + _
              "<button id=""RibbonSwapSlides_Button"" imageMso=""SmartArtRightToLeft"" label=""" + My.Resources.Ribbon_CustomShows_SwapSlides_Button + """ onAction=""SwapSlides_CallBack""/>" + _
              "<button idMso=""SlideShowCustom"" label=""" + LocalizeM.String_RibbonBuiltInCustomShows_Button + """ />" + _
              "<menuSeparator id=""RibbonCustomShowsSeperator"" title=""" + LocalizeM.String_RibbonCustomShows_MenuSeparatorTitle + """/>"


        '' Adding a button for each custom show. the "security.securityelement.escape" method is for encoding the inputted string to xml style (eg. double quotes are &quot;  etc')
        For i = 1 To aShows.Length - 1
            sShowName = aShows(i).Name
            sShowName = Security.SecurityElement.Escape(sShowName)
            Dim sDisplayShowName As String = sShowName.Replace("&amp;", "&amp;&amp;")
            sIdNum = CStr(i)
            xml += "<button id=""RibboncustomshowButton" + sIdNum + """ label=""" + sDisplayShowName + """ onAction=""RunCustomShow_CallBack"" tag=""" + sShowName + """/>"
        Next

        xml += "</menu>"

        Return xml
    End Function


#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
