<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EmptyShowsForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EmptyShowsForm))
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckedListBox1.CheckOnClick = True
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(73, 27)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CheckedListBox1.Size = New System.Drawing.Size(391, 259)
        Me.CheckedListBox1.TabIndex = 0
        Me.CheckedListBox1.ThreeDCheckBoxes = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(73, 305)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(85, 27)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Select All"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Button2.Location = New System.Drawing.Point(226, 305)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(85, 27)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Delete"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button3.Location = New System.Drawing.Point(379, 305)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(85, 27)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Cancel"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(70, 345)
        Me.Label1.MaximumSize = New System.Drawing.Size(391, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(373, 26)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "* Empty shows are Custom slide shows without any slide in it. Deleting it won't a" & _
    "ffect your presentation."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(70, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(391, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Found XXXX empty custom shows out of total YYYY"
        '
        'EmptyShowsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button3
        Me.ClientSize = New System.Drawing.Size(535, 380)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CheckedListBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EmptyShowsForm"
        Me.ShowInTaskbar = False
        Me.Text = "Empty Custom Shows"
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
