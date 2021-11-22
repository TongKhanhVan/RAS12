<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ACM_ADM
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.GridACMADM = New System.Windows.Forms.DataGridView()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CmbCustomer = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtAmt = New System.Windows.Forms.TextBox()
        Me.OptADM = New System.Windows.Forms.RadioButton()
        Me.OptACM = New System.Windows.Forms.RadioButton()
        Me.TxtDate = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtNote = New System.Windows.Forms.TextBox()
        Me.LblApprove = New System.Windows.Forms.LinkLabel()
        Me.LblUpdate = New System.Windows.Forms.LinkLabel()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.LblTRF2SMS = New System.Windows.Forms.LinkLabel()
        CType(Me.GridACMADM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridACMADM
        '
        Me.GridACMADM.AllowUserToAddRows = False
        Me.GridACMADM.AllowUserToDeleteRows = False
        Me.GridACMADM.BackgroundColor = System.Drawing.Color.FloralWhite
        Me.GridACMADM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridACMADM.Location = New System.Drawing.Point(1, 54)
        Me.GridACMADM.Name = "GridACMADM"
        Me.GridACMADM.RowHeadersVisible = False
        Me.GridACMADM.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridACMADM.Size = New System.Drawing.Size(768, 403)
        Me.GridACMADM.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(0, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(28, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Cust"
        '
        'CmbCustomer
        '
        Me.CmbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCustomer.FormattingEnabled = True
        Me.CmbCustomer.Location = New System.Drawing.Point(34, 6)
        Me.CmbCustomer.Name = "CmbCustomer"
        Me.CmbCustomer.Size = New System.Drawing.Size(106, 21)
        Me.CmbCustomer.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(146, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "VND"
        '
        'TxtAmt
        '
        Me.TxtAmt.Location = New System.Drawing.Point(182, 6)
        Me.TxtAmt.Name = "TxtAmt"
        Me.TxtAmt.Size = New System.Drawing.Size(83, 20)
        Me.TxtAmt.TabIndex = 13
        Me.TxtAmt.Text = "0"
        Me.TxtAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'OptADM
        '
        Me.OptADM.AutoSize = True
        Me.OptADM.Location = New System.Drawing.Point(91, 30)
        Me.OptADM.Name = "OptADM"
        Me.OptADM.Size = New System.Drawing.Size(49, 17)
        Me.OptADM.TabIndex = 14
        Me.OptADM.TabStop = True
        Me.OptADM.Text = "ADM"
        Me.OptADM.UseVisualStyleBackColor = True
        '
        'OptACM
        '
        Me.OptACM.AutoSize = True
        Me.OptACM.Location = New System.Drawing.Point(34, 30)
        Me.OptACM.Name = "OptACM"
        Me.OptACM.Size = New System.Drawing.Size(48, 17)
        Me.OptACM.TabIndex = 14
        Me.OptACM.TabStop = True
        Me.OptACM.Text = "ACM"
        Me.OptACM.UseVisualStyleBackColor = True
        '
        'TxtDate
        '
        Me.TxtDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtDate.Location = New System.Drawing.Point(600, 6)
        Me.TxtDate.Name = "TxtDate"
        Me.TxtDate.Size = New System.Drawing.Size(83, 20)
        Me.TxtDate.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(544, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(50, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "DueDate"
        '
        'TxtNote
        '
        Me.TxtNote.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.TxtNote.Location = New System.Drawing.Point(271, 6)
        Me.TxtNote.Name = "TxtNote"
        Me.TxtNote.Size = New System.Drawing.Size(269, 20)
        Me.TxtNote.TabIndex = 16
        Me.TxtNote.Text = "Note"
        '
        'LblApprove
        '
        Me.LblApprove.AutoSize = True
        Me.LblApprove.Location = New System.Drawing.Point(722, 32)
        Me.LblApprove.Name = "LblApprove"
        Me.LblApprove.Size = New System.Drawing.Size(47, 13)
        Me.LblApprove.TabIndex = 17
        Me.LblApprove.TabStop = True
        Me.LblApprove.Text = "Approve"
        '
        'LblUpdate
        '
        Me.LblUpdate.AutoSize = True
        Me.LblUpdate.Location = New System.Drawing.Point(179, 32)
        Me.LblUpdate.Name = "LblUpdate"
        Me.LblUpdate.Size = New System.Drawing.Size(95, 13)
        Me.LblUpdate.TabIndex = 17
        Me.LblUpdate.TabStop = True
        Me.LblUpdate.Text = "Update ACM ADM"
        Me.LblUpdate.Visible = False
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(0, 460)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 17
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        '
        'LblTRF2SMS
        '
        Me.LblTRF2SMS.AutoSize = True
        Me.LblTRF2SMS.Enabled = False
        Me.LblTRF2SMS.Location = New System.Drawing.Point(426, 32)
        Me.LblTRF2SMS.Name = "LblTRF2SMS"
        Me.LblTRF2SMS.Size = New System.Drawing.Size(114, 13)
        Me.LblTRF2SMS.TabIndex = 18
        Me.LblTRF2SMS.TabStop = True
        Me.LblTRF2SMS.Text = "Transfer Credit to SMS"
        Me.LblTRF2SMS.Visible = False
        '
        'ACM_ADM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(769, 478)
        Me.Controls.Add(Me.LblTRF2SMS)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.LblUpdate)
        Me.Controls.Add(Me.LblApprove)
        Me.Controls.Add(Me.TxtNote)
        Me.Controls.Add(Me.TxtDate)
        Me.Controls.Add(Me.OptACM)
        Me.Controls.Add(Me.OptADM)
        Me.Controls.Add(Me.TxtAmt)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.CmbCustomer)
        Me.Controls.Add(Me.GridACMADM)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ACM_ADM"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. ACM_ADM Updater"
        CType(Me.GridACMADM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridACMADM As System.Windows.Forms.DataGridView
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CmbCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtAmt As System.Windows.Forms.TextBox
    Friend WithEvents OptADM As System.Windows.Forms.RadioButton
    Friend WithEvents OptACM As System.Windows.Forms.RadioButton
    Friend WithEvents TxtDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtNote As System.Windows.Forms.TextBox
    Friend WithEvents LblApprove As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LblTRF2SMS As System.Windows.Forms.LinkLabel
End Class
