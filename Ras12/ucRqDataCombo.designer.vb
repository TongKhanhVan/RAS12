<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucRqDataCombo
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
        Me.lblName = New System.Windows.Forms.Label()
        Me.cboValue = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(3, 5)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(39, 13)
        Me.lblName.TabIndex = 5
        Me.lblName.Text = "Label1"
        '
        'cboValue
        '
        Me.cboValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboValue.FormattingEnabled = True
        Me.cboValue.Location = New System.Drawing.Point(336, 2)
        Me.cboValue.Name = "cboValue"
        Me.cboValue.Size = New System.Drawing.Size(258, 21)
        Me.cboValue.TabIndex = 6
        '
        'ucRqDataCombo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.cboValue)
        Me.Controls.Add(Me.lblName)
        Me.Name = "ucRqDataCombo"
        Me.Size = New System.Drawing.Size(700, 21)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents cboValue As System.Windows.Forms.ComboBox

End Class
