<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddCustomer
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
        Me.txtCustCode = New System.Windows.Forms.TextBox()
        Me.txtShortName = New System.Windows.Forms.TextBox()
        Me.txtFullName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LckCmdAdd = New System.Windows.Forms.Button()
        Me.LckCmdSave = New System.Windows.Forms.Button()
        Me.GridCustList = New System.Windows.Forms.DataGridView()
        Me.txtTaxCode = New System.Windows.Forms.TextBox()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GridCommDetail = New System.Windows.Forms.DataGridView()
        Me.LckLblDeactivate = New System.Windows.Forms.LinkLabel()
        Me.TxtEmail = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.LckLblReinstate = New System.Windows.Forms.LinkLabel()
        Me.ChkXXOnly = New System.Windows.Forms.CheckBox()
        Me.LckLblMoreDetail = New System.Windows.Forms.LinkLabel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CmbChannel = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TxtPhone = New System.Windows.Forms.TextBox()
        Me.LckChkNotCustomer = New System.Windows.Forms.CheckBox()
        Me.CmbLocation = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        CType(Me.GridCustList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridCommDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCustCode
        '
        Me.txtCustCode.Enabled = False
        Me.txtCustCode.Location = New System.Drawing.Point(65, 4)
        Me.txtCustCode.Name = "txtCustCode"
        Me.txtCustCode.Size = New System.Drawing.Size(85, 20)
        Me.txtCustCode.TabIndex = 0
        '
        'txtShortName
        '
        Me.txtShortName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtShortName.Location = New System.Drawing.Point(222, 4)
        Me.txtShortName.MaxLength = 16
        Me.txtShortName.Name = "txtShortName"
        Me.txtShortName.Size = New System.Drawing.Size(110, 20)
        Me.txtShortName.TabIndex = 0
        '
        'txtFullName
        '
        Me.txtFullName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFullName.Location = New System.Drawing.Point(65, 25)
        Me.txtFullName.Name = "txtFullName"
        Me.txtFullName.Size = New System.Drawing.Size(267, 20)
        Me.txtFullName.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "CustomerID"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(156, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Short Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(2, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Full Name"
        '
        'LckCmdAdd
        '
        Me.LckCmdAdd.Location = New System.Drawing.Point(742, 25)
        Me.LckCmdAdd.Name = "LckCmdAdd"
        Me.LckCmdAdd.Size = New System.Drawing.Size(43, 20)
        Me.LckCmdAdd.TabIndex = 7
        Me.LckCmdAdd.Text = "Add"
        Me.LckCmdAdd.UseVisualStyleBackColor = True
        Me.LckCmdAdd.Visible = False
        '
        'LckCmdSave
        '
        Me.LckCmdSave.Location = New System.Drawing.Point(684, 25)
        Me.LckCmdSave.Name = "LckCmdSave"
        Me.LckCmdSave.Size = New System.Drawing.Size(43, 20)
        Me.LckCmdSave.TabIndex = 7
        Me.LckCmdSave.Text = "Save"
        Me.LckCmdSave.UseVisualStyleBackColor = True
        Me.LckCmdSave.Visible = False
        '
        'GridCustList
        '
        Me.GridCustList.AllowUserToAddRows = False
        Me.GridCustList.AllowUserToDeleteRows = False
        Me.GridCustList.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
        Me.GridCustList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCustList.Location = New System.Drawing.Point(3, 100)
        Me.GridCustList.Name = "GridCustList"
        Me.GridCustList.ReadOnly = True
        Me.GridCustList.RowHeadersVisible = False
        Me.GridCustList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridCustList.Size = New System.Drawing.Size(487, 399)
        Me.GridCustList.TabIndex = 4
        '
        'txtTaxCode
        '
        Me.txtTaxCode.Location = New System.Drawing.Point(65, 47)
        Me.txtTaxCode.Name = "txtTaxCode"
        Me.txtTaxCode.Size = New System.Drawing.Size(85, 20)
        Me.txtTaxCode.TabIndex = 4
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(385, 47)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(297, 20)
        Me.txtAddress.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(2, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(50, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "TaxCode"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(334, 50)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 13)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Address"
        '
        'GridCommDetail
        '
        Me.GridCommDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCommDetail.Location = New System.Drawing.Point(491, 100)
        Me.GridCommDetail.Name = "GridCommDetail"
        Me.GridCommDetail.ReadOnly = True
        Me.GridCommDetail.RowHeadersVisible = False
        Me.GridCommDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridCommDetail.Size = New System.Drawing.Size(296, 399)
        Me.GridCommDetail.TabIndex = 8
        '
        'LckLblDeactivate
        '
        Me.LckLblDeactivate.AutoSize = True
        Me.LckLblDeactivate.Location = New System.Drawing.Point(531, 7)
        Me.LckLblDeactivate.Name = "LckLblDeactivate"
        Me.LckLblDeactivate.Size = New System.Drawing.Size(60, 13)
        Me.LckLblDeactivate.TabIndex = 9
        Me.LckLblDeactivate.TabStop = True
        Me.LckLblDeactivate.Text = "DeActivate"
        Me.LckLblDeactivate.Visible = False
        '
        'TxtEmail
        '
        Me.TxtEmail.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.TxtEmail.Location = New System.Drawing.Point(385, 25)
        Me.TxtEmail.Name = "TxtEmail"
        Me.TxtEmail.Size = New System.Drawing.Size(297, 20)
        Me.TxtEmail.TabIndex = 3
        Me.TxtEmail.Text = "Plz Enter upto TWO Emaisl 4 Generic Purpose"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(334, 27)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 13)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "Email"
        '
        'LckLblReinstate
        '
        Me.LckLblReinstate.AutoSize = True
        Me.LckLblReinstate.Location = New System.Drawing.Point(606, 7)
        Me.LckLblReinstate.Name = "LckLblReinstate"
        Me.LckLblReinstate.Size = New System.Drawing.Size(52, 13)
        Me.LckLblReinstate.TabIndex = 11
        Me.LckLblReinstate.TabStop = True
        Me.LckLblReinstate.Text = "Reinstate"
        Me.LckLblReinstate.Visible = False
        '
        'ChkXXOnly
        '
        Me.ChkXXOnly.AutoSize = True
        Me.ChkXXOnly.Location = New System.Drawing.Point(669, 6)
        Me.ChkXXOnly.Name = "ChkXXOnly"
        Me.ChkXXOnly.Size = New System.Drawing.Size(86, 17)
        Me.ChkXXOnly.TabIndex = 12
        Me.ChkXXOnly.Text = "InActiveOnly"
        Me.ChkXXOnly.UseVisualStyleBackColor = True
        '
        'LckLblMoreDetail
        '
        Me.LckLblMoreDetail.AutoSize = True
        Me.LckLblMoreDetail.Location = New System.Drawing.Point(453, 7)
        Me.LckLblMoreDetail.Name = "LckLblMoreDetail"
        Me.LckLblMoreDetail.Size = New System.Drawing.Size(61, 13)
        Me.LckLblMoreDetail.TabIndex = 13
        Me.LckLblMoreDetail.TabStop = True
        Me.LckLblMoreDetail.Text = "More Detail"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(334, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Channel"
        '
        'CmbChannel
        '
        Me.CmbChannel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbChannel.FormattingEnabled = True
        Me.CmbChannel.Location = New System.Drawing.Point(385, 3)
        Me.CmbChannel.Name = "CmbChannel"
        Me.CmbChannel.Size = New System.Drawing.Size(62, 21)
        Me.CmbChannel.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(156, 50)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(61, 13)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Phone Nbr."
        '
        'TxtPhone
        '
        Me.TxtPhone.Location = New System.Drawing.Point(222, 47)
        Me.TxtPhone.MaxLength = 11
        Me.TxtPhone.Name = "TxtPhone"
        Me.TxtPhone.Size = New System.Drawing.Size(110, 20)
        Me.TxtPhone.TabIndex = 5
        '
        'LckChkNotCustomer
        '
        Me.LckChkNotCustomer.AutoSize = True
        Me.LckChkNotCustomer.Checked = True
        Me.LckChkNotCustomer.CheckState = System.Windows.Forms.CheckState.Checked
        Me.LckChkNotCustomer.Location = New System.Drawing.Point(684, 49)
        Me.LckChkNotCustomer.Name = "LckChkNotCustomer"
        Me.LckChkNotCustomer.Size = New System.Drawing.Size(103, 17)
        Me.LckChkNotCustomer.TabIndex = 19
        Me.LckChkNotCustomer.Text = "NotYetCustomer"
        Me.LckChkNotCustomer.UseVisualStyleBackColor = True
        '
        'CmbLocation
        '
        Me.CmbLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbLocation.FormattingEnabled = True
        Me.CmbLocation.Items.AddRange(New Object() {"SGN", "HAN", "DAD", "HPH"})
        Me.CmbLocation.Location = New System.Drawing.Point(385, 69)
        Me.CmbLocation.Name = "CmbLocation"
        Me.CmbLocation.Size = New System.Drawing.Size(62, 21)
        Me.CmbLocation.TabIndex = 20
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(335, 72)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(48, 13)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "Location"
        '
        'AddCustomer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.CmbLocation)
        Me.Controls.Add(Me.LckChkNotCustomer)
        Me.Controls.Add(Me.TxtPhone)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.CmbChannel)
        Me.Controls.Add(Me.LckLblMoreDetail)
        Me.Controls.Add(Me.ChkXXOnly)
        Me.Controls.Add(Me.LckLblReinstate)
        Me.Controls.Add(Me.TxtEmail)
        Me.Controls.Add(Me.LckLblDeactivate)
        Me.Controls.Add(Me.GridCommDetail)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.txtTaxCode)
        Me.Controls.Add(Me.GridCustList)
        Me.Controls.Add(Me.LckCmdSave)
        Me.Controls.Add(Me.LckCmdAdd)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFullName)
        Me.Controls.Add(Me.txtShortName)
        Me.Controls.Add(Me.txtCustCode)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AddCustomer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Customer List"
        CType(Me.GridCustList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridCommDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtCustCode As System.Windows.Forms.TextBox
    Friend WithEvents txtShortName As System.Windows.Forms.TextBox
    Friend WithEvents txtFullName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LckCmdAdd As System.Windows.Forms.Button
    Friend WithEvents LckCmdSave As System.Windows.Forms.Button
    Friend WithEvents GridCustList As System.Windows.Forms.DataGridView
    Friend WithEvents txtTaxCode As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GridCommDetail As System.Windows.Forms.DataGridView
    Friend WithEvents LckLblDeactivate As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents LckLblReinstate As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkXXOnly As System.Windows.Forms.CheckBox
    Friend WithEvents LckLblMoreDetail As System.Windows.Forms.LinkLabel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CmbChannel As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtPhone As System.Windows.Forms.TextBox
    Friend WithEvents LckChkNotCustomer As System.Windows.Forms.CheckBox
    Friend WithEvents CmbLocation As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
End Class
