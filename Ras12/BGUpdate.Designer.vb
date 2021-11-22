<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BGUpdate
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
        Me.LblRenewInProc = New System.Windows.Forms.LinkLabel()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.LckLblUpdate = New System.Windows.Forms.LinkLabel()
        Me.GridBG = New System.Windows.Forms.DataGridView()
        Me.txtNote = New System.Windows.Forms.RichTextBox()
        Me.LblExportBG = New System.Windows.Forms.LinkLabel()
        Me.TxtAmount = New System.Windows.Forms.TextBox()
        Me.TxtContractNo = New System.Windows.Forms.TextBox()
        Me.OptCredit = New System.Windows.Forms.RadioButton()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TxtValidFrm = New System.Windows.Forms.DateTimePicker()
        Me.TxtValidThru = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.CmbBank = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.LblValidFrm = New System.Windows.Forms.Label()
        Me.OptDeposit = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.OptBG = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ChkSelectCustOnly = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbTVC = New System.Windows.Forms.ComboBox()
        Me.ChkActiveOnly = New System.Windows.Forms.CheckBox()
        Me.CmbContractOrgCopy = New System.Windows.Forms.ComboBox()
        Me.txtContractValidUntil = New System.Windows.Forms.DateTimePicker()
        Me.LckLblUnlock = New System.Windows.Forms.LinkLabel()
        Me.LblLock = New System.Windows.Forms.LinkLabel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CmbCustomer = New System.Windows.Forms.ComboBox()
        Me.LckLblApprove = New System.Windows.Forms.LinkLabel()
        CType(Me.GridBG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblRenewInProc
        '
        Me.LblRenewInProc.AutoSize = True
        Me.LblRenewInProc.Location = New System.Drawing.Point(230, 484)
        Me.LblRenewInProc.Name = "LblRenewInProc"
        Me.LblRenewInProc.Size = New System.Drawing.Size(101, 13)
        Me.LblRenewInProc.TabIndex = 54
        Me.LblRenewInProc.TabStop = True
        Me.LblRenewInProc.Text = "Renewal in Process"
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(746, 484)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 53
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        '
        'LckLblUpdate
        '
        Me.LckLblUpdate.AutoSize = True
        Me.LckLblUpdate.Location = New System.Drawing.Point(569, 76)
        Me.LckLblUpdate.Name = "LckLblUpdate"
        Me.LckLblUpdate.Size = New System.Drawing.Size(45, 13)
        Me.LckLblUpdate.TabIndex = 46
        Me.LckLblUpdate.TabStop = True
        Me.LckLblUpdate.Text = "Update "
        Me.LckLblUpdate.Visible = False
        '
        'GridBG
        '
        Me.GridBG.AllowUserToAddRows = False
        Me.GridBG.AllowUserToDeleteRows = False
        Me.GridBG.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridBG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridBG.Location = New System.Drawing.Point(-3, 101)
        Me.GridBG.Name = "GridBG"
        Me.GridBG.ReadOnly = True
        Me.GridBG.RowHeadersVisible = False
        Me.GridBG.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridBG.Size = New System.Drawing.Size(788, 376)
        Me.GridBG.TabIndex = 47
        '
        'txtNote
        '
        Me.txtNote.Location = New System.Drawing.Point(337, 73)
        Me.txtNote.Name = "txtNote"
        Me.txtNote.Size = New System.Drawing.Size(226, 22)
        Me.txtNote.TabIndex = 45
        Me.txtNote.Text = ""
        '
        'LblExportBG
        '
        Me.LblExportBG.AutoSize = True
        Me.LblExportBG.Location = New System.Drawing.Point(746, 76)
        Me.LblExportBG.Name = "LblExportBG"
        Me.LblExportBG.Size = New System.Drawing.Size(37, 13)
        Me.LblExportBG.TabIndex = 48
        Me.LblExportBG.TabStop = True
        Me.LblExportBG.Text = "Export"
        '
        'TxtAmount
        '
        Me.TxtAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAmount.Location = New System.Drawing.Point(85, 48)
        Me.TxtAmount.Name = "TxtAmount"
        Me.TxtAmount.Size = New System.Drawing.Size(104, 20)
        Me.TxtAmount.TabIndex = 41
        Me.TxtAmount.Text = "0"
        Me.TxtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtContractNo
        '
        Me.TxtContractNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtContractNo.Location = New System.Drawing.Point(494, 3)
        Me.TxtContractNo.Name = "TxtContractNo"
        Me.TxtContractNo.Size = New System.Drawing.Size(129, 20)
        Me.TxtContractNo.TabIndex = 29
        '
        'OptCredit
        '
        Me.OptCredit.AutoSize = True
        Me.OptCredit.Location = New System.Drawing.Point(312, 4)
        Me.OptCredit.Name = "OptCredit"
        Me.OptCredit.Size = New System.Drawing.Size(52, 17)
        Me.OptCredit.TabIndex = 28
        Me.OptCredit.Text = "Credit"
        Me.OptCredit.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(626, 27)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(68, 13)
        Me.Label10.TabIndex = 38
        Me.Label10.Text = "TV Company"
        '
        'TxtValidFrm
        '
        Me.TxtValidFrm.CustomFormat = "dd-MMM-yy"
        Me.TxtValidFrm.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TxtValidFrm.Location = New System.Drawing.Point(258, 48)
        Me.TxtValidFrm.Name = "TxtValidFrm"
        Me.TxtValidFrm.Size = New System.Drawing.Size(77, 20)
        Me.TxtValidFrm.TabIndex = 42
        '
        'TxtValidThru
        '
        Me.TxtValidThru.CustomFormat = "dd-MMM-yy"
        Me.TxtValidThru.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TxtValidThru.Location = New System.Drawing.Point(85, 70)
        Me.TxtValidThru.Name = "TxtValidThru"
        Me.TxtValidThru.Size = New System.Drawing.Size(105, 20)
        Me.TxtValidThru.TabIndex = 43
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Red
        Me.Label11.Location = New System.Drawing.Point(0, 52)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(87, 13)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Amount (VND)"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(0, 76)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(55, 13)
        Me.Label12.TabIndex = 39
        Me.Label12.Text = "Valid Thru"
        '
        'CmbBank
        '
        Me.CmbBank.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbBank.FormattingEnabled = True
        Me.CmbBank.Location = New System.Drawing.Point(258, 73)
        Me.CmbBank.Name = "CmbBank"
        Me.CmbBank.Size = New System.Drawing.Size(77, 21)
        Me.CmbBank.TabIndex = 44
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(196, 76)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 13)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Bank"
        '
        'LblValidFrm
        '
        Me.LblValidFrm.AutoSize = True
        Me.LblValidFrm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblValidFrm.Location = New System.Drawing.Point(196, 51)
        Me.LblValidFrm.Name = "LblValidFrm"
        Me.LblValidFrm.Size = New System.Drawing.Size(56, 13)
        Me.LblValidFrm.TabIndex = 32
        Me.LblValidFrm.Text = "Valid From"
        '
        'OptDeposit
        '
        Me.OptDeposit.AutoSize = True
        Me.OptDeposit.Location = New System.Drawing.Point(242, 4)
        Me.OptDeposit.Name = "OptDeposit"
        Me.OptDeposit.Size = New System.Drawing.Size(61, 17)
        Me.OptDeposit.TabIndex = 27
        Me.OptDeposit.Text = "Deposit"
        Me.OptDeposit.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(406, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(97, 13)
        Me.Label3.TabIndex = 36
        Me.Label3.Text = "Contract Org. Copy"
        '
        'OptBG
        '
        Me.OptBG.AutoSize = True
        Me.OptBG.Checked = True
        Me.OptBG.Location = New System.Drawing.Point(196, 4)
        Me.OptBG.Name = "OptBG"
        Me.OptBG.Size = New System.Drawing.Size(40, 17)
        Me.OptBG.TabIndex = 26
        Me.OptBG.TabStop = True
        Me.OptBG.Text = "BG"
        Me.OptBG.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(625, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 34
        Me.Label2.Text = "Valid Thru"
        '
        'ChkSelectCustOnly
        '
        Me.ChkSelectCustOnly.AutoSize = True
        Me.ChkSelectCustOnly.Checked = True
        Me.ChkSelectCustOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkSelectCustOnly.Location = New System.Drawing.Point(85, 483)
        Me.ChkSelectCustOnly.Name = "ChkSelectCustOnly"
        Me.ChkSelectCustOnly.Size = New System.Drawing.Size(139, 17)
        Me.ChkSelectCustOnly.TabIndex = 52
        Me.ChkSelectCustOnly.Text = "Selected Customer Only"
        Me.ChkSelectCustOnly.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(406, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 13)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "Contract Number"
        '
        'cmbTVC
        '
        Me.cmbTVC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTVC.FormattingEnabled = True
        Me.cmbTVC.Items.AddRange(New Object() {"Cargo", "Acct", "Admin"})
        Me.cmbTVC.Location = New System.Drawing.Point(706, 24)
        Me.cmbTVC.Name = "cmbTVC"
        Me.cmbTVC.Size = New System.Drawing.Size(77, 21)
        Me.cmbTVC.TabIndex = 40
        '
        'ChkActiveOnly
        '
        Me.ChkActiveOnly.AutoSize = True
        Me.ChkActiveOnly.Checked = True
        Me.ChkActiveOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkActiveOnly.Location = New System.Drawing.Point(1, 483)
        Me.ChkActiveOnly.Name = "ChkActiveOnly"
        Me.ChkActiveOnly.Size = New System.Drawing.Size(80, 17)
        Me.ChkActiveOnly.TabIndex = 51
        Me.ChkActiveOnly.Text = "Active Only"
        Me.ChkActiveOnly.UseVisualStyleBackColor = True
        '
        'CmbContractOrgCopy
        '
        Me.CmbContractOrgCopy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbContractOrgCopy.FormattingEnabled = True
        Me.CmbContractOrgCopy.Items.AddRange(New Object() {"Sales", "Acct", "Admin"})
        Me.CmbContractOrgCopy.Location = New System.Drawing.Point(561, 24)
        Me.CmbContractOrgCopy.Name = "CmbContractOrgCopy"
        Me.CmbContractOrgCopy.Size = New System.Drawing.Size(62, 21)
        Me.CmbContractOrgCopy.TabIndex = 37
        '
        'txtContractValidUntil
        '
        Me.txtContractValidUntil.Checked = False
        Me.txtContractValidUntil.CustomFormat = "dd-MMM-yy"
        Me.txtContractValidUntil.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtContractValidUntil.Location = New System.Drawing.Point(687, 3)
        Me.txtContractValidUntil.Name = "txtContractValidUntil"
        Me.txtContractValidUntil.ShowCheckBox = True
        Me.txtContractValidUntil.Size = New System.Drawing.Size(96, 20)
        Me.txtContractValidUntil.TabIndex = 30
        '
        'LckLblUnlock
        '
        Me.LckLblUnlock.AutoSize = True
        Me.LckLblUnlock.LinkColor = System.Drawing.Color.Blue
        Me.LckLblUnlock.Location = New System.Drawing.Point(606, 484)
        Me.LckLblUnlock.Name = "LckLblUnlock"
        Me.LckLblUnlock.Size = New System.Drawing.Size(41, 13)
        Me.LckLblUnlock.TabIndex = 50
        Me.LckLblUnlock.TabStop = True
        Me.LckLblUnlock.Text = "Unlock"
        Me.LckLblUnlock.Visible = False
        '
        'LblLock
        '
        Me.LblLock.AutoSize = True
        Me.LblLock.LinkColor = System.Drawing.Color.Blue
        Me.LblLock.Location = New System.Drawing.Point(532, 484)
        Me.LblLock.Name = "LblLock"
        Me.LblLock.Size = New System.Drawing.Size(31, 13)
        Me.LblLock.TabIndex = 49
        Me.LblLock.TabStop = True
        Me.LblLock.Text = "Lock"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(0, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(59, 13)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Customer"
        '
        'CmbCustomer
        '
        Me.CmbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCustomer.FormattingEnabled = True
        Me.CmbCustomer.Location = New System.Drawing.Point(63, 3)
        Me.CmbCustomer.Name = "CmbCustomer"
        Me.CmbCustomer.Size = New System.Drawing.Size(127, 21)
        Me.CmbCustomer.TabIndex = 25
        '
        'LckLblApprove
        '
        Me.LckLblApprove.AutoSize = True
        Me.LckLblApprove.Location = New System.Drawing.Point(664, 484)
        Me.LckLblApprove.Name = "LckLblApprove"
        Me.LckLblApprove.Size = New System.Drawing.Size(47, 13)
        Me.LckLblApprove.TabIndex = 55
        Me.LckLblApprove.TabStop = True
        Me.LckLblApprove.Text = "Approve"
        Me.LckLblApprove.Visible = False
        '
        'BGUpdate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.LckLblApprove)
        Me.Controls.Add(Me.LblRenewInProc)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.LckLblUpdate)
        Me.Controls.Add(Me.GridBG)
        Me.Controls.Add(Me.txtNote)
        Me.Controls.Add(Me.LblExportBG)
        Me.Controls.Add(Me.TxtAmount)
        Me.Controls.Add(Me.TxtContractNo)
        Me.Controls.Add(Me.OptCredit)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.TxtValidFrm)
        Me.Controls.Add(Me.TxtValidThru)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.CmbBank)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.LblValidFrm)
        Me.Controls.Add(Me.OptDeposit)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.OptBG)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ChkSelectCustOnly)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbTVC)
        Me.Controls.Add(Me.ChkActiveOnly)
        Me.Controls.Add(Me.CmbContractOrgCopy)
        Me.Controls.Add(Me.txtContractValidUntil)
        Me.Controls.Add(Me.LckLblUnlock)
        Me.Controls.Add(Me.LblLock)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.CmbCustomer)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BGUpdate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines  :: RAS :.  BG Updater"
        CType(Me.GridBG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblRenewInProc As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LckLblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents GridBG As System.Windows.Forms.DataGridView
    Friend WithEvents txtNote As System.Windows.Forms.RichTextBox
    Friend WithEvents LblExportBG As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtAmount As System.Windows.Forms.TextBox
    Friend WithEvents TxtContractNo As System.Windows.Forms.TextBox
    Friend WithEvents OptCredit As System.Windows.Forms.RadioButton
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtValidFrm As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtValidThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents CmbBank As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LblValidFrm As System.Windows.Forms.Label
    Friend WithEvents OptDeposit As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents OptBG As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ChkSelectCustOnly As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbTVC As System.Windows.Forms.ComboBox
    Friend WithEvents ChkActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents CmbContractOrgCopy As System.Windows.Forms.ComboBox
    Friend WithEvents txtContractValidUntil As System.Windows.Forms.DateTimePicker
    Friend WithEvents LckLblUnlock As System.Windows.Forms.LinkLabel
    Friend WithEvents LblLock As System.Windows.Forms.LinkLabel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CmbCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents LckLblApprove As System.Windows.Forms.LinkLabel
End Class
