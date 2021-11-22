<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucCdr
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtValue = New System.Windows.Forms.TextBox
        Me.lblName = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtValue
        '
        Me.txtValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtValue.Location = New System.Drawing.Point(339, 3)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(258, 20)
        Me.txtValue.TabIndex = 0
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(3, 6)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(39, 13)
        Me.lblName.TabIndex = 1
        Me.lblName.Text = "Label1"
        '
        'ucCdr
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.txtValue)
        Me.Name = "ucCdr"
        Me.Size = New System.Drawing.Size(600, 25)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents lblName As System.Windows.Forms.Label

End Class
