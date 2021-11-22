<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FoxSetting
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
        Me.GridAgentSI = New System.Windows.Forms.DataGridView()
        Me.TxtMinBLC = New System.Windows.Forms.TextBox()
        Me.TxtFQ = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LblSet = New System.Windows.Forms.LinkLabel()
        Me.LblUpdateIniPSW = New System.Windows.Forms.LinkLabel()
        Me.TxtPSW = New System.Windows.Forms.TextBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabBLC = New System.Windows.Forms.TabPage()
        Me.LblCustomer = New System.Windows.Forms.LinkLabel()
        Me.LblRQDelete = New System.Windows.Forms.LinkLabel()
        Me.LblAddSI = New System.Windows.Forms.LinkLabel()
        Me.CmbCustomer = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TxtSI = New System.Windows.Forms.TextBox()
        Me.TabSMS = New System.Windows.Forms.TabPage()
        Me.LblUpdateQNbr = New System.Windows.Forms.LinkLabel()
        Me.TxtQNbr = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.LblViewSentMail = New System.Windows.Forms.LinkLabel()
        Me.LblCFM_PSW_Changed_QQ = New System.Windows.Forms.LinkLabel()
        Me.LblCFM_PSW_Changed_RemoveSI = New System.Windows.Forms.LinkLabel()
        Me.TxtSkype = New System.Windows.Forms.TextBox()
        Me.TxtMobileNo = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TxtStaff = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.LblUpdateMobileNo = New System.Windows.Forms.LinkLabel()
        Me.GridSMS = New System.Windows.Forms.DataGridView()
        Me.TxtSIList = New System.Windows.Forms.TextBox()
        CType(Me.GridAgentSI, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabBLC.SuspendLayout()
        Me.TabSMS.SuspendLayout()
        CType(Me.GridSMS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridAgentSI
        '
        Me.GridAgentSI.AllowUserToAddRows = False
        Me.GridAgentSI.AllowUserToDeleteRows = False
        Me.GridAgentSI.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridAgentSI.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridAgentSI.Location = New System.Drawing.Point(1, 1)
        Me.GridAgentSI.Name = "GridAgentSI"
        Me.GridAgentSI.RowHeadersVisible = False
        Me.GridAgentSI.Size = New System.Drawing.Size(470, 350)
        Me.GridAgentSI.TabIndex = 5
        '
        'TxtMinBLC
        '
        Me.TxtMinBLC.Location = New System.Drawing.Point(79, 59)
        Me.TxtMinBLC.Name = "TxtMinBLC"
        Me.TxtMinBLC.Size = New System.Drawing.Size(108, 20)
        Me.TxtMinBLC.TabIndex = 1
        Me.TxtMinBLC.Text = "0"
        Me.TxtMinBLC.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtFQ
        '
        Me.TxtFQ.Location = New System.Drawing.Point(79, 82)
        Me.TxtFQ.MaxLength = 8
        Me.TxtFQ.Name = "TxtFQ"
        Me.TxtFQ.Size = New System.Drawing.Size(49, 20)
        Me.TxtFQ.TabIndex = 3
        Me.TxtFQ.Text = "0"
        Me.TxtFQ.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 85)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Interval Coef"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Min Balance"
        '
        'LblSet
        '
        Me.LblSet.AutoSize = True
        Me.LblSet.Enabled = False
        Me.LblSet.Location = New System.Drawing.Point(164, 85)
        Me.LblSet.Name = "LblSet"
        Me.LblSet.Size = New System.Drawing.Size(23, 13)
        Me.LblSet.TabIndex = 3
        Me.LblSet.TabStop = True
        Me.LblSet.Text = "Set"
        '
        'LblUpdateIniPSW
        '
        Me.LblUpdateIniPSW.AutoSize = True
        Me.LblUpdateIniPSW.Enabled = False
        Me.LblUpdateIniPSW.Location = New System.Drawing.Point(157, 234)
        Me.LblUpdateIniPSW.Name = "LblUpdateIniPSW"
        Me.LblUpdateIniPSW.Size = New System.Drawing.Size(87, 13)
        Me.LblUpdateIniPSW.TabIndex = 4
        Me.LblUpdateIniPSW.TabStop = True
        Me.LblUpdateIniPSW.Text = "Update Ini. PSW"
        '
        'TxtPSW
        '
        Me.TxtPSW.Location = New System.Drawing.Point(62, 231)
        Me.TxtPSW.MaxLength = 8
        Me.TxtPSW.Name = "TxtPSW"
        Me.TxtPSW.Size = New System.Drawing.Size(93, 20)
        Me.TxtPSW.TabIndex = 2
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabBLC)
        Me.TabControl1.Controls.Add(Me.TabSMS)
        Me.TabControl1.Location = New System.Drawing.Point(471, 1)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(253, 352)
        Me.TabControl1.TabIndex = 6
        '
        'TabBLC
        '
        Me.TabBLC.Controls.Add(Me.LblCustomer)
        Me.TabBLC.Controls.Add(Me.LblRQDelete)
        Me.TabBLC.Controls.Add(Me.LblAddSI)
        Me.TabBLC.Controls.Add(Me.CmbCustomer)
        Me.TabBLC.Controls.Add(Me.Label3)
        Me.TabBLC.Controls.Add(Me.Label2)
        Me.TabBLC.Controls.Add(Me.TxtSI)
        Me.TabBLC.Controls.Add(Me.TxtMinBLC)
        Me.TabBLC.Controls.Add(Me.TxtFQ)
        Me.TabBLC.Controls.Add(Me.LblSet)
        Me.TabBLC.Controls.Add(Me.Label1)
        Me.TabBLC.Location = New System.Drawing.Point(4, 22)
        Me.TabBLC.Name = "TabBLC"
        Me.TabBLC.Padding = New System.Windows.Forms.Padding(3)
        Me.TabBLC.Size = New System.Drawing.Size(245, 326)
        Me.TabBLC.TabIndex = 0
        Me.TabBLC.Text = "Balance And Interval"
        Me.TabBLC.UseVisualStyleBackColor = True
        '
        'LblCustomer
        '
        Me.LblCustomer.AutoSize = True
        Me.LblCustomer.Enabled = False
        Me.LblCustomer.Location = New System.Drawing.Point(7, 14)
        Me.LblCustomer.Name = "LblCustomer"
        Me.LblCustomer.Size = New System.Drawing.Size(51, 13)
        Me.LblCustomer.TabIndex = 12
        Me.LblCustomer.TabStop = True
        Me.LblCustomer.Text = "Customer"
        '
        'LblRQDelete
        '
        Me.LblRQDelete.AutoSize = True
        Me.LblRQDelete.Location = New System.Drawing.Point(2, 310)
        Me.LblRQDelete.Name = "LblRQDelete"
        Me.LblRQDelete.Size = New System.Drawing.Size(169, 13)
        Me.LblRQDelete.TabIndex = 11
        Me.LblRQDelete.TabStop = True
        Me.LblRQDelete.Text = "Q 4PSW Change B4 Taking Back"
        Me.LblRQDelete.Visible = False
        '
        'LblAddSI
        '
        Me.LblAddSI.AutoSize = True
        Me.LblAddSI.Enabled = False
        Me.LblAddSI.Location = New System.Drawing.Point(143, 36)
        Me.LblAddSI.Name = "LblAddSI"
        Me.LblAddSI.Size = New System.Drawing.Size(44, 13)
        Me.LblAddSI.TabIndex = 10
        Me.LblAddSI.TabStop = True
        Me.LblAddSI.Text = "Add S/I"
        '
        'CmbCustomer
        '
        Me.CmbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCustomer.FormattingEnabled = True
        Me.CmbCustomer.Location = New System.Drawing.Point(79, 9)
        Me.CmbCustomer.Name = "CmbCustomer"
        Me.CmbCustomer.Size = New System.Drawing.Size(108, 23)
        Me.CmbCustomer.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "SignIn"
        '
        'TxtSI
        '
        Me.TxtSI.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSI.Location = New System.Drawing.Point(79, 33)
        Me.TxtSI.MaxLength = 3
        Me.TxtSI.Name = "TxtSI"
        Me.TxtSI.Size = New System.Drawing.Size(49, 20)
        Me.TxtSI.TabIndex = 0
        '
        'TabSMS
        '
        Me.TabSMS.Controls.Add(Me.LblUpdateQNbr)
        Me.TabSMS.Controls.Add(Me.TxtQNbr)
        Me.TabSMS.Controls.Add(Me.Label6)
        Me.TabSMS.Controls.Add(Me.LblViewSentMail)
        Me.TabSMS.Controls.Add(Me.LblCFM_PSW_Changed_RemoveSI)
        Me.TabSMS.Controls.Add(Me.TxtSkype)
        Me.TabSMS.Controls.Add(Me.TxtMobileNo)
        Me.TabSMS.Controls.Add(Me.Label8)
        Me.TabSMS.Controls.Add(Me.Label7)
        Me.TabSMS.Controls.Add(Me.Label5)
        Me.TabSMS.Controls.Add(Me.TxtStaff)
        Me.TabSMS.Controls.Add(Me.Label4)
        Me.TabSMS.Controls.Add(Me.LblUpdateMobileNo)
        Me.TabSMS.Controls.Add(Me.GridSMS)
        Me.TabSMS.Controls.Add(Me.TxtPSW)
        Me.TabSMS.Controls.Add(Me.LblUpdateIniPSW)
        Me.TabSMS.Location = New System.Drawing.Point(4, 22)
        Me.TabSMS.Name = "TabSMS"
        Me.TabSMS.Padding = New System.Windows.Forms.Padding(3)
        Me.TabSMS.Size = New System.Drawing.Size(245, 326)
        Me.TabSMS.TabIndex = 1
        Me.TabSMS.Text = "PSW and SMS"
        Me.TabSMS.UseVisualStyleBackColor = True
        '
        'LblUpdateQNbr
        '
        Me.LblUpdateQNbr.AutoSize = True
        Me.LblUpdateQNbr.Enabled = False
        Me.LblUpdateQNbr.Location = New System.Drawing.Point(157, 259)
        Me.LblUpdateQNbr.Name = "LblUpdateQNbr"
        Me.LblUpdateQNbr.Size = New System.Drawing.Size(76, 13)
        Me.LblUpdateQNbr.TabIndex = 15
        Me.LblUpdateQNbr.TabStop = True
        Me.LblUpdateQNbr.Text = "Update Q Nbr."
        '
        'TxtQNbr
        '
        Me.TxtQNbr.Location = New System.Drawing.Point(62, 252)
        Me.TxtQNbr.MaxLength = 3
        Me.TxtQNbr.Name = "TxtQNbr"
        Me.TxtQNbr.Size = New System.Drawing.Size(93, 20)
        Me.TxtQNbr.TabIndex = 14
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 255)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(38, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Q Nbr."
        '
        'LblViewSentMail
        '
        Me.LblViewSentMail.AutoSize = True
        Me.LblViewSentMail.Location = New System.Drawing.Point(164, 300)
        Me.LblViewSentMail.Name = "LblViewSentMail"
        Me.LblViewSentMail.Size = New System.Drawing.Size(82, 13)
        Me.LblViewSentMail.TabIndex = 12
        Me.LblViewSentMail.TabStop = True
        Me.LblViewSentMail.Text = "View Sent Mails"
        Me.LblViewSentMail.Visible = False
        '
        'LblCFM_PSW_Changed_QQ
        '
        Me.LblCFM_PSW_Changed_QQ.AutoSize = True
        Me.LblCFM_PSW_Changed_QQ.Location = New System.Drawing.Point(564, 356)
        Me.LblCFM_PSW_Changed_QQ.Name = "LblCFM_PSW_Changed_QQ"
        Me.LblCFM_PSW_Changed_QQ.Size = New System.Drawing.Size(153, 13)
        Me.LblCFM_PSW_Changed_QQ.TabIndex = 11
        Me.LblCFM_PSW_Changed_QQ.TabStop = True
        Me.LblCFM_PSW_Changed_QQ.Text = "PSW Changed and Stop Sales"
        '
        'LblCFM_PSW_Changed_RemoveSI
        '
        Me.LblCFM_PSW_Changed_RemoveSI.AutoSize = True
        Me.LblCFM_PSW_Changed_RemoveSI.Location = New System.Drawing.Point(0, 300)
        Me.LblCFM_PSW_Changed_RemoveSI.Name = "LblCFM_PSW_Changed_RemoveSI"
        Me.LblCFM_PSW_Changed_RemoveSI.Size = New System.Drawing.Size(155, 13)
        Me.LblCFM_PSW_Changed_RemoveSI.TabIndex = 10
        Me.LblCFM_PSW_Changed_RemoveSI.TabStop = True
        Me.LblCFM_PSW_Changed_RemoveSI.Text = "PSW Changed and Remove SI"
        Me.LblCFM_PSW_Changed_RemoveSI.Visible = False
        '
        'TxtSkype
        '
        Me.TxtSkype.Location = New System.Drawing.Point(62, 188)
        Me.TxtSkype.MaxLength = 11
        Me.TxtSkype.Name = "TxtSkype"
        Me.TxtSkype.Size = New System.Drawing.Size(183, 20)
        Me.TxtSkype.TabIndex = 7
        '
        'TxtMobileNo
        '
        Me.TxtMobileNo.Location = New System.Drawing.Point(62, 209)
        Me.TxtMobileNo.MaxLength = 11
        Me.TxtMobileNo.Name = "TxtMobileNo"
        Me.TxtMobileNo.Size = New System.Drawing.Size(93, 20)
        Me.TxtMobileNo.TabIndex = 7
        Me.TxtMobileNo.Text = "0"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 191)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(37, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Skype"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 234)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 13)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "PSW"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 212)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "MobileNo."
        '
        'TxtStaff
        '
        Me.TxtStaff.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStaff.Location = New System.Drawing.Point(62, 167)
        Me.TxtStaff.Name = "TxtStaff"
        Me.TxtStaff.Size = New System.Drawing.Size(183, 20)
        Me.TxtStaff.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 170)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Staff"
        '
        'LblUpdateMobileNo
        '
        Me.LblUpdateMobileNo.AutoSize = True
        Me.LblUpdateMobileNo.Location = New System.Drawing.Point(157, 212)
        Me.LblUpdateMobileNo.Name = "LblUpdateMobileNo"
        Me.LblUpdateMobileNo.Size = New System.Drawing.Size(82, 13)
        Me.LblUpdateMobileNo.TabIndex = 8
        Me.LblUpdateMobileNo.TabStop = True
        Me.LblUpdateMobileNo.Text = "Update Contact"
        '
        'GridSMS
        '
        Me.GridSMS.AllowUserToAddRows = False
        Me.GridSMS.AllowUserToDeleteRows = False
        Me.GridSMS.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridSMS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridSMS.Location = New System.Drawing.Point(0, 0)
        Me.GridSMS.Name = "GridSMS"
        Me.GridSMS.RowHeadersVisible = False
        Me.GridSMS.Size = New System.Drawing.Size(245, 161)
        Me.GridSMS.TabIndex = 5
        '
        'TxtSIList
        '
        Me.TxtSIList.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSIList.Location = New System.Drawing.Point(1, 353)
        Me.TxtSIList.Name = "TxtSIList"
        Me.TxtSIList.Size = New System.Drawing.Size(470, 20)
        Me.TxtSIList.TabIndex = 9
        '
        'FoxSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(722, 376)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.TxtSIList)
        Me.Controls.Add(Me.GridAgentSI)
        Me.Controls.Add(Me.LblCFM_PSW_Changed_QQ)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FoxSetting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Fox Setting"
        CType(Me.GridAgentSI, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabBLC.ResumeLayout(False)
        Me.TabBLC.PerformLayout()
        Me.TabSMS.ResumeLayout(False)
        Me.TabSMS.PerformLayout()
        CType(Me.GridSMS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridAgentSI As System.Windows.Forms.DataGridView
    Friend WithEvents TxtMinBLC As System.Windows.Forms.TextBox
    Friend WithEvents TxtFQ As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblSet As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUpdateIniPSW As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtPSW As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabBLC As System.Windows.Forms.TabPage
    Friend WithEvents TabSMS As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtSI As System.Windows.Forms.TextBox
    Friend WithEvents LblAddSI As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents GridSMS As System.Windows.Forms.DataGridView
    Friend WithEvents LblUpdateMobileNo As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtMobileNo As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtStaff As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents LblRQDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCustomer As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCFM_PSW_Changed_RemoveSI As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtSIList As System.Windows.Forms.TextBox
    Friend WithEvents LblCFM_PSW_Changed_QQ As System.Windows.Forms.LinkLabel
    Friend WithEvents LblViewSentMail As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUpdateQNbr As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtQNbr As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtSkype As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
End Class
