<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WhatsNewForm
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
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.LabelVersion = New System.Windows.Forms.Label()
        Me.LabelMsg = New System.Windows.Forms.Label()
        Me.ButtonOK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox1.HideSelection = False
        Me.TextBox1.Location = New System.Drawing.Point(10, 85)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(357, 111)
        Me.TextBox1.TabIndex = 0
        '
        'LabelVersion
        '
        Me.LabelVersion.AutoSize = True
        Me.LabelVersion.Location = New System.Drawing.Point(8, 65)
        Me.LabelVersion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LabelVersion.Name = "LabelVersion"
        Me.LabelVersion.Size = New System.Drawing.Size(82, 13)
        Me.LabelVersion.TabIndex = 1
        Me.LabelVersion.Text = "Version X.X.X.X"
        '
        'LabelMsg
        '
        Me.LabelMsg.AutoSize = True
        Me.LabelMsg.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.LabelMsg.Location = New System.Drawing.Point(8, 7)
        Me.LabelMsg.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LabelMsg.MaximumSize = New System.Drawing.Size(356, 0)
        Me.LabelMsg.Name = "LabelMsg"
        Me.LabelMsg.Size = New System.Drawing.Size(354, 48)
        Me.LabelMsg.TabIndex = 2
        Me.LabelMsg.Text = "Congratulations! your Custom Shows Wiz addin has been updated."
        '
        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(159, 203)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(69, 24)
        Me.ButtonOK.TabIndex = 3
        Me.ButtonOK.Text = "OK"
        Me.ButtonOK.UseVisualStyleBackColor = True
        '
        'WhatsNewForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(382, 239)
        Me.ControlBox = False
        Me.Controls.Add(Me.ButtonOK)
        Me.Controls.Add(Me.LabelMsg)
        Me.Controls.Add(Me.LabelVersion)
        Me.Controls.Add(Me.TextBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WhatsNewForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Whats New"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents LabelVersion As System.Windows.Forms.Label
    Friend WithEvents LabelMsg As System.Windows.Forms.Label
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
End Class
