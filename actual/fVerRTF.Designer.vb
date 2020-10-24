<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fVerRTF
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
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.btnText = New System.Windows.Forms.Button()
        Me.btnRTF = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.DetectUrls = False
        Me.RichTextBox1.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RichTextBox1.Location = New System.Drawing.Point(0, 0)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(684, 335)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'btnText
        '
        Me.btnText.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnText.AutoSize = True
        Me.btnText.Location = New System.Drawing.Point(0, 338)
        Me.btnText.Margin = New System.Windows.Forms.Padding(0)
        Me.btnText.Name = "btnText"
        Me.btnText.Size = New System.Drawing.Size(220, 23)
        Me.btnText.TabIndex = 1
        Me.btnText.Text = "Mostrar el texto del código del formato RTF"
        Me.btnText.UseVisualStyleBackColor = True
        '
        'btnRTF
        '
        Me.btnRTF.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRTF.AutoSize = True
        Me.btnRTF.Location = New System.Drawing.Point(223, 338)
        Me.btnRTF.Margin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.btnRTF.Name = "btnRTF"
        Me.btnRTF.Size = New System.Drawing.Size(166, 23)
        Me.btnRTF.TabIndex = 2
        Me.btnRTF.Text = "Mostrar el texto en formato RTF"
        Me.btnRTF.UseVisualStyleBackColor = True
        '
        'fVerRTF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(684, 361)
        Me.Controls.Add(Me.btnRTF)
        Me.Controls.Add(Me.btnText)
        Me.Controls.Add(Me.RichTextBox1)
        Me.MinimumSize = New System.Drawing.Size(450, 300)
        Me.Name = "fVerRTF"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ver código RTF"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents btnText As System.Windows.Forms.Button
    Private WithEvents btnRTF As System.Windows.Forms.Button
    Private WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
End Class
