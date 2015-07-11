Module LocalizeM
    Public isRTL As Boolean = False
    Public RTL_Setting As Windows.Forms.MessageBoxOptions = 0
    Public Ribbon_XML_FileName As String = "PowerPointAddIn1.Ribbon1_en.xml"
    Public String_Context_CreateNewShow_Button As String = "Create New..."
    Public String_CreateNewShow_Msg As String = "Enter New Custom Show Name"
    Public String_CreateNewShow_Msg_Title As String = "Create New Custom Show"
    Public String_Error As String = "Error"
    Public String_Action_Canceled As String = "Action Canceled"
    Public String_CreateNewShow_NoName_Msg As String = "Action canceled. No name was entered."
    Public String_CreateNewShow_LongName_Msg As String = "Please enter a name no longer than 31 characters. Please try a different name."
    Public String_CannotAddShow_Msg As String = "Sorry, Can not add custom show!"
    Public String_CannotAddShowDuplicate_Msg As String = "Another custom show with this name already exists. Please try a different name."
    Public String_CannotAppendShow_Msg As String = "Sorry, Can not append custom show!"
    Public String_CannotReplaceShow_Msg As String = "Sorry, Can not replace custom show!"
    Public String_AppendReplace_Msg_Title As String = "Append/Replace "
    Public String_AppendReplace_Msg1 As String = "Would you like to append selected slides to this show?" & vbNewLine & vbNewLine & "Press [YES] to append" & vbNewLine & "Press [NO] to replace"
    Public String_AppendReplace_Msg2 As String = "with selected slides"
    Public String_RibbonCustomShows_MenuSeparatorTitle As String = "Custom Shows List:"
    Public String_RibbonBuiltInCustomShows_Button As String = "Custom Shows..."
    Public String_RibbonCustomShows_FindEmpty_Button As String = "Find Empty Shows..."
    Public String_DeleteCustomShows_Msg As String = "Are you sure you want to delete XXXX Custom shows?"
    Public String_DeleteCustomShows_Title As String = "Deleting Custom Shows"
    Public String_DeleteCustomShows_Err As String = "Error in deleting ""XXXX"" custom show"
    Public String_DeleteCustomShows_NoChecked_Msg As String = "No custom show was selected. Please select at least one show."
    Public String_DeleteCustomShows_Done_Msg As String = "Successfully deleted XXXX shows out of YYYY selected"
    Public String_Form_Emptyshows_Title As String = "Empty Custom Shows"
    Public String_Form_Emptyshows_label1 As String = "* Empty shows are Custom slide shows without any slide in it. Deleting it won't affect your presentation."
    Public String_Form_Emptyshows_label2 As String = "Found XXXX empty custom shows out of total YYYY"
    Public String_Button_Cancel As String = "Cancel"
    Public String_Button_Delete As String = "Delete"
    Public String_Button_SelectAll As String = "Select All"
    Public String_Button_OK As String = "OK"
    Public String_Button_Next As String = "Next"
    Public String_FirstInstall_Featuring As String = "Featuring: "
    Public String_FirstInstall_Feature0_Desc As String = "Create new shows directly from the slides pane. just select the slides, right click and choose ""Create custom show"""
    Public String_FirstInstall_Feature0_Title As String = "Instant Create new custom shows"
    Public String_FirstInstall_Feature1_Desc As String = "Find and erase all empty custom shows left overs from previous file versions"
    Public String_FirstInstall_Feature1_Title As String = "Erase empty custom shows"
    Public String_FirstInstall_MainMsg As String = "Congratulations! you've just installed Custom Shows Wiz!"


    Sub InitializeStrings()
        If ThisAddIn.OfficeLanguage = Office.MsoLanguageID.msoLanguageIDHebrew Then
            isRTL = True
            RTL_Setting = Windows.Forms.MessageBoxOptions.RtlReading Or Windows.Forms.MessageBoxOptions.RightAlign
            Ribbon_XML_FileName = "PowerPointAddIn1.Ribbon1_he.xml"
            String_Context_CreateNewShow_Button = "צור חדש..."
            String_CreateNewShow_Msg = "הכנס שם לתצוגה האישית החדשה"
            String_CreateNewShow_Msg_Title = "יצירת תצוגה אישית"
            String_Error = "שגיאה"
            String_Action_Canceled = "הפעולה בוטלה"
            String_CreateNewShow_NoName_Msg = "הפעולה בוטלה. לא הוזן שם לתצוגה."
            String_CreateNewShow_LongName_Msg = "שם של תצוגה אישית אינו יכול להיות ארוך יותר מ-31 תווים. נא להכניס שם אחר."
            String_CannotAddShow_Msg = "מצטער, לא מצליח ליצור תצוגה אישית"
            String_CannotAddShowDuplicate_Msg = "קיימת כבר תצוגה אישית בשם זה. אנא נסו שם אחר."
            String_CannotAppendShow_Msg = "מצטערת, לא מצליחה להוסיף שקפים לתצוגה האישית"
            String_CannotReplaceShow_Msg = "מצטערת, לא מצליחה להחליף תצוגה אישית"
            String_AppendReplace_Msg_Title = "החלפה/הוספה"
            String_AppendReplace_Msg1 = "האם תרצה להוסיף את השקפים הנבחרים לתצוגה האישית?" & vbNewLine & vbNewLine & "לחץ [כן] כדי להוסיף" & vbNewLine & "לחץ [לא] כדי להחליף את"
            String_AppendReplace_Msg2 = "עם השקפים הנבחרים"
            String_RibbonCustomShows_MenuSeparatorTitle = "רשימת תצוגות אישיות:"
            String_RibbonBuiltInCustomShows_Button = "הצגות מותאמות אישית..."
            String_RibbonCustomShows_FindEmpty_Button = "מצא תצוגות ריקות..."
            String_DeleteCustomShows_Msg = "למחוק XXXX תצוגות מיוחדות?"
            String_DeleteCustomShows_Title = "מחיקת תצוגות מיוחדות"
            String_DeleteCustomShows_Err = "לא הצלחתי למחוק את התצוגה המיוחדת ""XXXX"""
            String_DeleteCustomShows_NoChecked_Msg = "לא נבחרה אף תצוגה מיוחדת למחיקה"
            String_DeleteCustomShows_Done_Msg = "נמחקו בהצלחה XXXX תצוגות מיוחדות מתוך YYYY שנבחרו"
            String_Form_Emptyshows_Title = "תצוגות מיוחדות ריקות"
            String_Form_Emptyshows_label1 = "* תצוגות ריקות הן תצוגות שלא מכילות שום שקף. מחיקתן לא תשפיע על המצגת."
            String_Form_Emptyshows_label2 = "נמצאו XXXX תצוגות ריקות מתוך סך של YYYY תצוגות"
            String_Button_Cancel = "ביטול"
            String_Button_Delete = "מחק"
            String_Button_SelectAll = "בחר הכל"
            String_Button_OK = "אישור"
            String_Button_Next = "הבא"
            String_FirstInstall_Featuring = "הכלים: "
            String_FirstInstall_Feature0_Desc = "צור הצגות חדשות במהירות! פשוט בחר את השקפים הרצויים, לחץ על הכפתור הימני ובחר ב""צור הצגה מותאמת אישית"""
            String_FirstInstall_Feature0_Title = "יצירה בקליק של הצגות מותאמות אישית"
            String_FirstInstall_Feature1_Desc = "מצא ומחק את כל אותן תצוגות מיוחדות ישנות וריקות מגרסאות קודמות של המצגת"
            String_FirstInstall_Feature1_Title = "מצא ומחק תצוגות ריקות"
            String_FirstInstall_MainMsg = "ברכות! כרגע התקנת את תוסף Custom Shows Wiz!"
        End If
    End Sub

    '' Localize (RTL in case of hebrew UI) the form and all it's controls
    Sub Localize_Forms(ByRef fForm As Windows.Forms.Form)
        Dim cFormControl As Windows.Forms.Control

        Try
            fForm.RightToLeftLayout = LocalizeM.isRTL
            If LocalizeM.isRTL Then
                fForm.RightToLeft = Windows.Forms.RightToLeft.Yes
                For Each cFormControl In fForm.Controls
                    cFormControl.RightToLeft = Windows.Forms.RightToLeft.Yes
                Next
            End If
        Catch ex As Exception
            '' Exists only for the unknown error that may occur when calling the 
            '' cFormControl.RightToLeft.maybe it's not valid for a some kind of 
            '' control i'm not aware of.
            '' Program should not terminate for a problem in localization.
            System.Diagnostics.Debug.Print("Error in Localize_Forms subroutine:")
            System.Diagnostics.Debug.Print(ex.Message)
        End Try
        
    End Sub
End Module
