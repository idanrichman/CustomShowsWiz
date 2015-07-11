<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SwapSlidesForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SwapSlidesForm))
        Me.Button3 = New System.Windows.Forms.Button()
        Me.ButtonOK = New System.Windows.Forms.Button()
        Me.Label_Title = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PictureBox_Arrow = New System.Windows.Forms.PictureBox()
        Me.PictureBox_NewSld = New System.Windows.Forms.PictureBox()
        Me.PictureBox_OldSld = New System.Windows.Forms.PictureBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        CType(Me.PictureBox_Arrow, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox_NewSld, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox_OldSld, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button3
        '
        resources.ApplyResources(Me.Button3, "Button3")
        Me.Button3.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button3.Name = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ButtonOK
        '
        resources.ApplyResources(Me.ButtonOK, "ButtonOK")
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.UseVisualStyleBackColor = True
        '
        'Label_Title
        '
        resources.ApplyResources(Me.Label_Title, "Label_Title")
        Me.Label_Title.Name = "Label_Title"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'PictureBox_Arrow
        '
        resources.ApplyResources(Me.PictureBox_Arrow, "PictureBox_Arrow")
        Me.PictureBox_Arrow.Name = "PictureBox_Arrow"
        Me.PictureBox_Arrow.TabStop = False
        '
        'PictureBox_NewSld
        '
        resources.ApplyResources(Me.PictureBox_NewSld, "PictureBox_NewSld")
        Me.PictureBox_NewSld.Name = "PictureBox_NewSld"
        Me.PictureBox_NewSld.TabStop = False
        '
        'PictureBox_OldSld
        '
        resources.ApplyResources(Me.PictureBox_OldSld, "PictureBox_OldSld")
        Me.PictureBox_OldSld.Name = "PictureBox_OldSld"
        Me.PictureBox_OldSld.TabStop = False
        '
        'TableLayoutPanel1
        '
        resources.ApplyResources(Me.TableLayoutPanel1, "TableLayoutPanel1")
        Me.TableLayoutPanel1.Controls.Add(Me.PictureBox_OldSld, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.PictureBox_NewSld, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.PictureBox_Arrow, 1, 1)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        '
        'SwapSlidesForm
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button3
        Me.Controls.Add(Me.Label_Title)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.ButtonOK)
        Me.Controls.Add(Me.Button3)
        Me.Name = "SwapSlidesForm"
        Me.ShowIcon = False
        CType(Me.PictureBox_Arrow, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox_NewSld, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox_OldSld, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

End Sub
    Friend WithEvents PictureBox_NewSld As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox_Arrow As System.Windows.Forms.PictureBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents Label_Title As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox_OldSld As System.Windows.Forms.PictureBox
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
End Class
