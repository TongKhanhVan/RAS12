<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TimedMsgBox
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
        Me.components = New System.ComponentModel.Container()
        Me.TxtMsg = New System.Windows.Forms.RichTextBox()
        Me.LblGotIt = New System.Windows.Forms.LinkLabel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'TxtMsg
        '
        Me.TxtMsg.Enabled = False
        Me.TxtMsg.Location = New System.Drawing.Point(1, 0)
        Me.TxtMsg.Name = "TxtMsg"
        Me.TxtMsg.ReadOnly = True
        Me.TxtMsg.Size = New System.Drawing.Size(382, 93)
        Me.TxtMsg.TabIndex = 0
        Me.TxtMsg.Text = ""
        '
        'LblGotIt
        '
        Me.LblGotIt.AutoSize = True
        Me.LblGotIt.Location = New System.Drawing.Point(350, 96)
        Me.LblGotIt.Name = "LblGotIt"
        Me.LblGotIt.Size = New System.Drawing.Size(33, 13)
        Me.LblGotIt.TabIndex = 1
        Me.LblGotIt.TabStop = True
        Me.LblGotIt.Text = "Got It"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 6400
        '
        'TimedMsgBox
        '
        Me.AcceptButton = Me.LblGotIt
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 114)
        Me.Controls.Add(Me.LblGotIt)
        Me.Controls.Add(Me.TxtMsg)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TimedMsgBox"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "TransViet Travel :: RAS12 :. System Message"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtMsg As System.Windows.Forms.RichTextBox
    Friend WithEvents LblGotIt As System.Windows.Forms.LinkLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
