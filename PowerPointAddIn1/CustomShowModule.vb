Imports System.Windows.Forms

Module CustomShowModule

    Function GetAllCustomShows() As PowerPoint.NamedSlideShow()
        Dim aShows() As PowerPoint.NamedSlideShow
        Dim iShowNum As Integer
        Dim i As Integer = 1
        Dim oNamedSlideShow As PowerPoint.NamedSlideShow


        iShowNum = Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.NamedSlideShows.Count
        ReDim aShows(iShowNum)

        For Each oNamedSlideShow In Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.NamedSlideShows
            aShows(i) = oNamedSlideShow
            i += 1
        Next

        Return aShows
    End Function

    '' Putting the selected slides ID's into an array. used for adding a new custom slide show.
    '' SmartOrder: if TRUE then the routine checks for the active slide, if it's a lower slide index then most probably user selected multiple slide
    ''             with the Shift Key, hence a reverse order should be make.
    Private Function SelectionToSafeArrayOfSlideIDs(Optional SmartOrder As Boolean = False) As Integer()
        Dim safeArrayOfSlideIDs() As Integer
        Dim iSelectionCount As Integer
        Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim bReverseOrder As Boolean '' Accepts TRUE when shift key is used to select multiple slides

        iSelectionCount = oPowerpoint.ActiveWindow.Selection.SlideRange.Count
        ReDim safeArrayOfSlideIDs(iSelectionCount)

        If SmartOrder = True Then
            '' I assume that the user selects from top to bottom (lower index first), otherwise this test won't really work correctly
            If oPowerpoint.ActiveWindow.Selection.SlideRange(1).SlideIndex < oPowerpoint.ActiveWindow.Selection.SlideRange(iSelectionCount).SlideIndex Then
                bReverseOrder = True
            End If
        End If

        '' when selecting manualy slide by slide then slides should be entered backwards to the array so that the NamedSlideShow will show in proper order
        For i = 1 To iSelectionCount
            If bReverseOrder = True Then
                safeArrayOfSlideIDs(i) = oPowerpoint.ActiveWindow.Selection.SlideRange(i).SlideID
            Else
                safeArrayOfSlideIDs(safeArrayOfSlideIDs.Length - 1 - i) = oPowerpoint.ActiveWindow.Selection.SlideRange(i).SlideID
            End If

        Next

        Return safeArrayOfSlideIDs
    End Function

    Private Function ListToSafeArrayOfSlideIDs(lListOfSlidesIDS As List(Of Integer)) As Integer()
        Dim safeArrayOfSlideIDs(lListOfSlidesIDS.Count) As Integer
        Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application

        For i = 1 To lListOfSlidesIDS.Count
            safeArrayOfSlideIDs(i - 1) = lListOfSlidesIDS(i - 1)
        Next

        Return safeArrayOfSlideIDs
    End Function

    Sub CreateNewShow()
        Dim sShowName As String = ""
        Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim safeArrayOfSlideIDs() As Integer

        '' Prompting the user for a name for the custom show
        Do
            sShowName = InputBox(LocalizeM.String_CreateNewShow_Msg, LocalizeM.String_CreateNewShow_Msg_Title, DefaultResponse:=sShowName)
            ''Custom Show name cannot be longer than 31 characters
            If sShowName.Length > 31 Then
                Windows.Forms.MessageBox.Show(text:=LocalizeM.String_CreateNewShow_LongName_Msg, caption:=LocalizeM.String_Error,
                                                      buttons:=Windows.Forms.MessageBoxButtons.OK, icon:=MessageBoxIcon.Exclamation, defaultButton:=MessageBoxDefaultButton.Button1,
                                                      options:=LocalizeM.RTL_Setting)
            End If
        Loop While sShowName.Length > 31

        '' Checking if any name was entered.
        If Not String.IsNullOrWhiteSpace(sShowName) Then
            safeArrayOfSlideIDs = SelectionToSafeArrayOfSlideIDs(SmartOrder:=True)

            Try
                oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows.Add(sShowName, safeArrayOfSlideIDs)
            Catch ex As System.Runtime.InteropServices.COMException
                If ex.ErrorCode = -2147188160 Then '' Error code for "another show with that name already exists"
                    Windows.Forms.MessageBox.Show(text:=LocalizeM.String_CannotAddShowDuplicate_Msg, caption:=LocalizeM.String_Error,
                                                      buttons:=Windows.Forms.MessageBoxButtons.OK, icon:=MessageBoxIcon.Information, defaultButton:=MessageBoxDefaultButton.Button1,
                                                      options:=LocalizeM.RTL_Setting)
                    CreateNewShow() '' Re-Prompting the user to enter a name
                Else
                    MsgBox(LocalizeM.String_CannotAddShow_Msg & vbNewLine & LocalizeM.String_Error & ": " & ex.Message, vbCritical, Title:=LocalizeM.String_Error)
                End If

            End Try

        Else
            Windows.Forms.MessageBox.Show(text:=LocalizeM.String_CreateNewShow_NoName_Msg, caption:=LocalizeM.String_Error,
                                                      buttons:=Windows.Forms.MessageBoxButtons.OK, icon:=MessageBoxIcon.Information, defaultButton:=MessageBoxDefaultButton.Button1,
                                                      options:=LocalizeM.RTL_Setting)
        End If

    End Sub

    Sub ReplaceShow(sShowName As String)
        Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim safeArrayOfSlideIDs() As Integer

        safeArrayOfSlideIDs = SelectionToSafeArrayOfSlideIDs(SmartOrder:=True)

        Try
            oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows(sShowName).Delete()
            oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows.Add(sShowName, safeArrayOfSlideIDs)
        Catch ex As Exception
            MsgBox(LocalizeM.String_CannotReplaceShow_Msg & vbNewLine & LocalizeM.String_Error & ": " & ex.Message, vbCritical, Title:=LocalizeM.String_Error)
        End Try

    End Sub

    Private Function ConvertObjectArrayToInteger(obj As Object) As Integer()
        Dim iReturnArray() As Integer
        Dim i As Integer
        Dim objCount As Integer

        objCount = CType(obj, Array).Length - 1
        ReDim iReturnArray(objCount)
        For i = 1 To objCount
            iReturnArray(i) = CType(obj(i), Integer)
        Next

        Return iReturnArray
    End Function

    Sub AppendShow(sShowName As String)
        Dim oPowerpoint As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim oOldSafeArrayOfSlideIDs() As Integer
        Dim oNewSafeArrayOfSlideIDs() As Integer
        Dim oMixSafeArrayOfSlideIDs() As Integer

        oOldSafeArrayOfSlideIDs = ConvertObjectArrayToInteger(oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows(sShowName).SlideIDs)
        oNewSafeArrayOfSlideIDs = SelectionToSafeArrayOfSlideIDs(SmartOrder:=True)
        '' Merging the two arrays. duplicates are left out.
        oMixSafeArrayOfSlideIDs = oOldSafeArrayOfSlideIDs.Union(oNewSafeArrayOfSlideIDs).ToArray

        Try
            oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows(sShowName).Delete()
            oPowerpoint.ActivePresentation.SlideShowSettings.NamedSlideShows.Add(sShowName, oMixSafeArrayOfSlideIDs)
        Catch ex As Exception
            MsgBox(LocalizeM.String_CannotAppendShow_Msg & vbNewLine & LocalizeM.String_Error & ": " & ex.Message, vbCritical, Title:=LocalizeM.String_Error)
        End Try


        '' +++ TODO: check if i need to release array's memory
    End Sub

    Sub SwapSlides_LaunchForm()
        Dim fSwapSlidesForm As New SwapSlidesForm
        Dim SldRng As PowerPoint.SlideRange

        Try
            SldRng = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange
            If SldRng.Count = 2 Then
                fSwapSlidesForm.ShowDialog() ''.ShowDialog is for showing in modal mode. .Show is for non-modal mode.
            Else
                Windows.Forms.MessageBox.Show(text:=My.Resources.Form_SwapSlides_BadSelection_ErrMsg, caption:=LocalizeM.String_Error,
                                              buttons:=Windows.Forms.MessageBoxButtons.OK, icon:=MessageBoxIcon.Exclamation, defaultButton:=MessageBoxDefaultButton.Button1,
                                              options:=LocalizeM.RTL_Setting)
            End If
        Catch ex As Exception
            MsgBox(LocalizeM.String_CannotAppendShow_Msg & vbNewLine & LocalizeM.String_Error & ": " & ex.Message, vbCritical, Title:=LocalizeM.String_Error)
        End Try
        
    End Sub

    Sub SwapSlides_AllCustomShows(OldSld As PowerPoint.Slide, NewSld As PowerPoint.Slide)
        Dim fSwapSlidesForm As New SwapSlidesForm
        Dim oNamedSlideShow As PowerPoint.NamedSlideShow
        Dim lTempSlidesIDList As List(Of Integer) = New List(Of Integer)
        Dim lTempSlideID As Integer
        Dim i As Integer
        Dim j As Integer
        Dim bHasChanged As Boolean = False
        Dim iShowsChangedCount As Integer
        Dim sSummaryMsg As String

        Try

            For j = 1 To Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.NamedSlideShows.Count
                oNamedSlideShow = Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.NamedSlideShows(j)
                For i = 1 To oNamedSlideShow.Count
                    lTempSlideID = CType(oNamedSlideShow.SlideIDs(i), Integer)
                    If lTempSlideID = OldSld.SlideID Then '' swap the needed slides
                        lTempSlideID = NewSld.SlideID
                        bHasChanged = True '' inform that a change been made and there is a need to rebuild this custom show
                    End If
                    lTempSlidesIDList.Add(lTempSlideID)
                Next

                If bHasChanged = True Then
                    Dim sNewShowName As String = oNamedSlideShow.Name
                    oNamedSlideShow.Delete()
                    j = j - 1
                    Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.NamedSlideShows.Add(sNewShowName, ListToSafeArrayOfSlideIDs(lTempSlidesIDList))
                    iShowsChangedCount += 1
                End If

                bHasChanged = False '' Clears the flag
                lTempSlidesIDList.Clear() '' Clears the temporary list
            Next

            If Not iShowsChangedCount = 0 Then
                sSummaryMsg = My.Resources.Form_SwapSlides_Summary_Msg.Replace("{{1}}", iShowsChangedCount.ToString)
            Else
                sSummaryMsg = My.Resources.Form_SwapSlides_Summary_NoChange_Msg
            End If

            Windows.Forms.MessageBox.Show(text:=sSummaryMsg, caption:=My.Resources.Form_SwapSlides_Summary_TItle,
                                                      buttons:=Windows.Forms.MessageBoxButtons.OK, icon:=MessageBoxIcon.Asterisk, defaultButton:=MessageBoxDefaultButton.Button1,
                                                      options:=LocalizeM.RTL_Setting)

        Catch ex As Exception
            MsgBox(LocalizeM.String_CannotAppendShow_Msg & vbNewLine & LocalizeM.String_Error & ": " & ex.Message, vbCritical, Title:=LocalizeM.String_Error)

        End Try
    End Sub
End Module
