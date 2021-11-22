<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NhapSoDu
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
        Me.txtCustFullName = New System.Windows.Forms.TextBox
        Me.CmbCustomer = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.CmbCNCurr = New System.Windows.Forms.ComboBox
        Me.txtCNDescription = New System.Windows.Forms.TextBox
        Me.txtCNAmount = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.CmdUpdate = New System.Windows.Forms.Button
        Me.TxtAsOf = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.CmdRefresh = New System.Windows.Forms.Button
        Me.CmbDebType = New System.Windows.Forms.ComboBox
        Me.GridSoDu = New System.Windows.Forms.DataGridView
        Me.LblDelete = New System.Windows.Forms.LinkLabel
        CType(Me.GridSoDu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCustFullName
        '
        Me.txtCustFullName.BackColor = System.Drawing.Color.White
        Me.txtCustFullName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustFullName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustFullName.Location = New System.Drawing.Point(170, 13)
        Me.txtCustFullName.Name = "txtCustFullName"
        Me.txtCustFullName.ReadOnly = True
        Me.txtCustFullName.Size = New System.Drawing.Size(232, 20)
        Me.txtCustFullName.TabIndex = 10
        '
        'CmbCustomer
        '
        Me.CmbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCustomer.FormattingEnabled = True
        Me.CmbCustomer.Location = New System.Drawing.Point(67, 12)
        Me.CmbCustomer.Name = "CmbCustomer"
        Me.CmbCustomer.Size = New System.Drawing.Size(100, 21)
        Me.CmbCustomer.TabIndex = 8
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(4, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 13)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Customer"
        '
        'CmbCNCurr
        '
        Me.CmbCNCurr.Enabled = False
        Me.CmbCNCurr.FormattingEnabled = True
        Me.CmbCNCurr.Location = New System.Drawing.Point(67, 35)
        Me.CmbCNCurr.Name = "CmbCNCurr"
        Me.CmbCNCurr.Size = New System.Drawing.Size(56, 21)
        Me.CmbCNCurr.TabIndex = 14
        Me.CmbCNCurr.Text = "VND"
        '
        'txtCNDescription
        '
        Me.txtCNDescription.Location = New System.Drawing.Point(67, 57)
        Me.txtCNDescription.Name = "txtCNDescription"
        Me.txtCNDescription.Size = New System.Drawing.Size(335, 20)
        Me.txtCNDescription.TabIndex = 17
        '
        'txtCNAmount
        '
        Me.txtCNAmount.Location = New System.Drawing.Point(170, 36)
        Me.txtCNAmount.Name = "txtCNAmount"
        Me.txtCNAmount.Size = New System.Drawing.Size(97, 20)
        Me.txtCNAmount.TabIndex = 15
        Me.txtCNAmount.Text = "0"
        Me.txtCNAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(4, 60)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(60, 13)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "Description"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(129, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Amount"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(4, 39)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(49, 13)
        Me.Label13.TabIndex = 13
        Me.Label13.Text = "Currency"
        '
        'CmdUpdate
        '
        Me.CmdUpdate.Location = New System.Drawing.Point(408, 55)
        Me.CmdUpdate.Name = "CmdUpdate"
        Me.CmdUpdate.Size = New System.Drawing.Size(71, 22)
        Me.CmdUpdate.TabIndex = 18
        Me.CmdUpdate.Text = "Update"
        Me.CmdUpdate.UseVisualStyleBackColor = True
        '
        'TxtAsOf
        '
        Me.TxtAsOf.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtAsOf.Location = New System.Drawing.Point(311, 36)
        Me.TxtAsOf.Name = "TxtAsOf"
        Me.TxtAsOf.Size = New System.Drawing.Size(91, 20)
        Me.TxtAsOf.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(273, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(33, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "As Of"
        '
        'CmdRefresh
        '
        Me.CmdRefresh.Location = New System.Drawing.Point(549, 55)
        Me.CmdRefresh.Name = "CmdRefresh"
        Me.CmdRefresh.Size = New System.Drawing.Size(86, 22)
        Me.CmdRefresh.TabIndex = 19
        Me.CmdRefresh.Text = "View Balance"
        Me.CmdRefresh.UseVisualStyleBackColor = True
        '
        'CmbDebType
        '
        Me.CmbDebType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbDebType.FormattingEnabled = True
        Me.CmbDebType.Items.AddRange(New Object() {"PSP", "PPD"})
        Me.CmbDebType.Location = New System.Drawing.Point(409, 12)
        Me.CmbDebType.Name = "CmbDebType"
        Me.CmbDebType.Size = New System.Drawing.Size(70, 21)
        Me.CmbDebType.TabIndex = 20
        '
        'GridSoDu
        '
        Me.GridSoDu.AllowUserToAddRows = False
        Me.GridSoDu.AllowUserToDeleteRows = False
        Me.GridSoDu.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridSoDu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridSoDu.Location = New System.Drawing.Point(1, 83)
        Me.GridSoDu.Name = "GridSoDu"
        Me.GridSoDu.RowHeadersVisible = False
        Me.GridSoDu.Size = New System.Drawing.Size(634, 216)
        Me.GridSoDu.TabIndex = 21
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(597, 302)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 22
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        '
        'NhapSoDu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(635, 318)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.GridSoDu)
        Me.Controls.Add(Me.CmbDebType)
        Me.Controls.Add(Me.CmdRefresh)
        Me.Controls.Add(Me.TxtAsOf)
        Me.Controls.Add(Me.CmdUpdate)
        Me.Controls.Add(Me.CmbCNCurr)
        Me.Controls.Add(Me.txtCNDescription)
        Me.Controls.Add(Me.txtCNAmount)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.txtCustFullName)
        Me.Controls.Add(Me.CmbCustomer)
        Me.Controls.Add(Me.Label6)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "NhapSoDu"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Credit Control : Beginning Balance"
        CType(Me.GridSoDu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtCustFullName As System.Windows.Forms.TextBox
    Friend WithEvents CmbCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CmbCNCurr As System.Windows.Forms.ComboBox
    Friend WithEvents txtCNDescription As System.Windows.Forms.TextBox
    Friend WithEvents txtCNAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents CmdUpdate As System.Windows.Forms.Button
    Friend WithEvents TxtAsOf As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CmdRefresh As System.Windows.Forms.Button
    Friend WithEvents CmbDebType As System.Windows.Forms.ComboBox
    Friend WithEvents GridSoDu As System.Windows.Forms.DataGridView
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
End Class
