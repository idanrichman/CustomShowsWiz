<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FirstInstallForm
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
        Me.ButtonOK = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Featuring_Label = New System.Windows.Forms.Label()
        Me.Desc_Label = New System.Windows.Forms.Label()
        Me.Button_Next = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.SuspendLayout
        '
        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(230, 374)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(69, 24)
        Me.ButtonOK.TabIndex = 4
        Me.ButtonOK.Text = "OK"
        Me.ButtonOK.UseVisualStyleBackColor = True
        Me.ButtonOK.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(483, 24)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Congratulations! you've just installed Custom Shows Wiz!" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Featuring_Label
        '
        Me.Featuring_Label.AutoSize = True
        Me.Featuring_Label.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Featuring_Label.Location = New System.Drawing.Point(51, 60)
        Me.Featuring_Label.Name = "Featuring_Label"
        Me.Featuring_Label.Size = New System.Drawing.Size(122, 24)
        Me.Featuring_Label.TabIndex = 7
        Me.Featuring_Label.Text = "Featuring: {1}"
        '
        'Desc_Label
        '
        Me.Desc_Label.AutoSize = True
        Me.Desc_Label.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Desc_Label.Location = New System.Drawing.Point(37, 305)
        Me.Desc_Label.MaximumSize = New System.Drawing.Size(500, 0)
        Me.Desc_Label.Name = "Desc_Label"
        Me.Desc_Label.Size = New System.Drawing.Size(99, 16)
        Me.Desc_Label.TabIndex = 8
        Me.Desc_Label.Text = "Description: {1}"
        '
        'Button_Next
        '
        Me.Button_Next.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_Next.Location = New System.Drawing.Point(230, 374)
        Me.Button_Next.Name = "Button_Next"
        Me.Button_Next.Size = New System.Drawing.Size(69, 24)
        Me.Button_Next.TabIndex = 9
        Me.Button_Next.Text = "Next"
        Me.Button_Next.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = Global.PowerPointAddIn1.My.Resources.Resources.CreateNewPng
        Me.PictureBox1.InitialImage = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(154, 98)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(250, 180)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'FirstInstallForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(529, 411)
        Me.Controls.Add(Me.Button_Next)
        Me.Controls.Add(Me.Desc_Label)
        Me.Controls.Add(Me.Featuring_Label)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ButtonOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "FirstInstallForm"
        Me.ShowIcon = false
        Me.ShowInTaskbar = false
        Me.Text = "Custom Shows Wiz"
        CType(Me.PictureBox1,System.ComponentModel.ISupportInitialize).EndInit
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Featuring_Label As System.Windows.Forms.Label
    Friend WithEvents Desc_Label As System.Windows.Forms.Label
    Friend WithEvents Button_Next As System.Windows.Forms.Button
End Class
