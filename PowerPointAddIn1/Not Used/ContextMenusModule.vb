Module ContextMenusModule

    Private WithEvents cCtrl As Office.CommandBarButton
    Private cCtrl2 As Office.CommandBarControl
    Private WithEvents cmdBar As Office.CommandBar

    Sub Load_Context_Menus()

        Dim cCtrl3 As Office.CommandBarControl

        cmdBar = Globals.ThisAddIn.Application.CommandBars("Thumbnails") '' also "Slide Sorter" and "Slider Sorter"

        cCtrl = cmdBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Id:=1, Temporary:=True)
        With cCtrl
            .Caption = "Create CustomShow"
            .BeginGroup = True
            .FaceId = 581
        End With

        cCtrl2 = cmdBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlPopup, Temporary:=True)

        cCtrl3 = cCtrl2.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Temporary:=True)
        cCtrl3.Caption = "try"



    End Sub

    Sub Unload_Context_Menus()

    End Sub

    Private Sub AddShow_Click() Handles cCtrl.Click
        MsgBox("addshow")
    End Sub

  
End Module

