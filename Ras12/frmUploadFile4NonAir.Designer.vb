<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUploadFile4NonAir
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
        Me.lbkOK = New System.Windows.Forms.LinkLabel()
        Me.lbkCancel = New System.Windows.Forms.LinkLabel()
        Me.cboFileType = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'lbkOK
        '
        Me.lbkOK.AutoSize = True
        Me.lbkOK.Location = New System.Drawing.Point(55, 53)
        Me.lbkOK.Name = "lbkOK"
        Me.lbkOK.Size = New System.Drawing.Size(22, 13)
        Me.lbkOK.TabIndex = 0
        Me.lbkOK.TabStop = True
        Me.lbkOK.Text = "OK"
        '
        'lbkCancel
        '
        Me.lbkCancel.AutoSize = True
        Me.lbkCancel.Location = New System.Drawing.Point(106, 53)
        Me.lbkCancel.Name = "lbkCancel"
        Me.lbkCancel.Size = New System.Drawing.Size(40, 13)
        Me.lbkCancel.TabIndex = 1
        Me.lbkCancel.TabStop = True
        Me.lbkCancel.Text = "Cancel"
        '
        'cboFileType
        '
        Me.cboFileType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFileType.FormattingEnabled = True
        Me.cboFileType.Items.AddRange(New Object() {"Registration", "Quotation"})
        Me.cboFileType.Location = New System.Drawing.Point(12, 12)
        Me.cboFileType.Name = "cboFileType"
        Me.cboFileType.Size = New System.Drawing.Size(184, 21)
        Me.cboFileType.TabIndex = 2
        '
        'frmUploadFile4NonAir
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 173)
        Me.Controls.Add(Me.cboFileType)
        Me.Controls.Add(Me.lbkCancel)
        Me.Controls.Add(Me.lbkOK)
        Me.Name = "frmUploadFile4NonAir"
        Me.Text = "UploadFile4NonAir"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbkOK As System.Windows.Forms.LinkLabel
    Friend WithEvents lbkCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents cboFileType As System.Windows.Forms.ComboBox
End Class
