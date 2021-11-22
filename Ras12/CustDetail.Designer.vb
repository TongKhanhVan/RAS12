<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CustDetail
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
        Me.GridCustList = New System.Windows.Forms.DataGridView()
        Me.CmbChannel = New System.Windows.Forms.ComboBox()
        Me.LblUpdateOID = New System.Windows.Forms.LinkLabel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GridOffice = New System.Windows.Forms.DataGridView()
        Me.TxtOfficeID = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CmbChannelFilter = New System.Windows.Forms.ComboBox()
        Me.LblRemoveOID = New System.Windows.Forms.LinkLabel()
        Me.ChkQOnly = New System.Windows.Forms.CheckBox()
        Me.LblApproveOID = New System.Windows.Forms.LinkLabel()
        Me.TxtSearch = New System.Windows.Forms.TextBox()
        Me.LblSearch = New System.Windows.Forms.LinkLabel()
        Me.ChkAll = New System.Windows.Forms.CheckBox()
        Me.ChkAL = New System.Windows.Forms.CheckedListBox()
        Me.LblUpdateChannel = New System.Windows.Forms.LinkLabel()
        Me.LblUpdateAL = New System.Windows.Forms.LinkLabel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabAL = New System.Windows.Forms.TabPage()
        Me.TabChannel = New System.Windows.Forms.TabPage()
        Me.GrpCA = New System.Windows.Forms.GroupBox()
        Me.CmbAL = New System.Windows.Forms.ComboBox()
        Me.CmbSBU = New System.Windows.Forms.ComboBox()
        Me.TxtCA = New System.Windows.Forms.TextBox()
        Me.LblDeleteCA = New System.Windows.Forms.LinkLabel()
        Me.LblUpdateCA = New System.Windows.Forms.LinkLabel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.GridCA = New System.Windows.Forms.DataGridView()
        Me.TabOID = New System.Windows.Forms.TabPage()
        Me.LstOIDUsage = New System.Windows.Forms.CheckedListBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.LblAddUsage = New System.Windows.Forms.LinkLabel()
        Me.TabEmail = New System.Windows.Forms.TabPage()
        Me.LblUpdateEmail = New System.Windows.Forms.LinkLabel()
        Me.GridEmail = New System.Windows.Forms.DataGridView()
        Me.eMailType = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.EmailAcct = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.GridCustList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridOffice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabAL.SuspendLayout()
        Me.TabChannel.SuspendLayout()
        Me.GrpCA.SuspendLayout()
        CType(Me.GridCA, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabOID.SuspendLayout()
        Me.TabEmail.SuspendLayout()
        CType(Me.GridEmail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridCustList
        '
        Me.GridCustList.AllowUserToAddRows = False
        Me.GridCustList.AllowUserToDeleteRows = False
        Me.GridCustList.BackgroundColor = System.Drawing.Color.Ivory
        Me.GridCustList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCustList.Location = New System.Drawing.Point(0, 1)
        Me.GridCustList.Name = "GridCustList"
        Me.GridCustList.RowHeadersVisible = False
        Me.GridCustList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridCustList.Size = New System.Drawing.Size(476, 472)
        Me.GridCustList.TabIndex = 0
        '
        'CmbChannel
        '
        Me.CmbChannel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbChannel.FormattingEnabled = True
        Me.CmbChannel.Location = New System.Drawing.Point(55, 6)
        Me.CmbChannel.Name = "CmbChannel"
        Me.CmbChannel.Size = New System.Drawing.Size(80, 21)
        Me.CmbChannel.TabIndex = 2
        '
        'LblUpdateOID
        '
        Me.LblUpdateOID.AutoSize = True
        Me.LblUpdateOID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblUpdateOID.Location = New System.Drawing.Point(145, 7)
        Me.LblUpdateOID.Name = "LblUpdateOID"
        Me.LblUpdateOID.Size = New System.Drawing.Size(48, 13)
        Me.LblUpdateOID.TabIndex = 3
        Me.LblUpdateOID.TabStop = True
        Me.LblUpdateOID.Text = "Add OID"
        Me.LblUpdateOID.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Channel"
        '
        'GridOffice
        '
        Me.GridOffice.AllowUserToAddRows = False
        Me.GridOffice.AllowUserToDeleteRows = False
        Me.GridOffice.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridOffice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridOffice.Location = New System.Drawing.Point(3, 54)
        Me.GridOffice.Name = "GridOffice"
        Me.GridOffice.RowHeadersVisible = False
        Me.GridOffice.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridOffice.Size = New System.Drawing.Size(298, 323)
        Me.GridOffice.TabIndex = 5
        '
        'TxtOfficeID
        '
        Me.TxtOfficeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtOfficeID.Location = New System.Drawing.Point(65, 4)
        Me.TxtOfficeID.MaxLength = 9
        Me.TxtOfficeID.Name = "TxtOfficeID"
        Me.TxtOfficeID.Size = New System.Drawing.Size(80, 20)
        Me.TxtOfficeID.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "OfficeID"
        '
        'CmbChannelFilter
        '
        Me.CmbChannelFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbChannelFilter.FormattingEnabled = True
        Me.CmbChannelFilter.Location = New System.Drawing.Point(3, 476)
        Me.CmbChannelFilter.Name = "CmbChannelFilter"
        Me.CmbChannelFilter.Size = New System.Drawing.Size(55, 21)
        Me.CmbChannelFilter.TabIndex = 7
        '
        'LblRemoveOID
        '
        Me.LblRemoveOID.AutoSize = True
        Me.LblRemoveOID.Location = New System.Drawing.Point(232, 428)
        Me.LblRemoveOID.Name = "LblRemoveOID"
        Me.LblRemoveOID.Size = New System.Drawing.Size(69, 13)
        Me.LblRemoveOID.TabIndex = 8
        Me.LblRemoveOID.TabStop = True
        Me.LblRemoveOID.Text = "Remove OID"
        '
        'ChkQOnly
        '
        Me.ChkQOnly.AutoSize = True
        Me.ChkQOnly.Location = New System.Drawing.Point(3, 427)
        Me.ChkQOnly.Name = "ChkQOnly"
        Me.ChkQOnly.Size = New System.Drawing.Size(58, 17)
        Me.ChkQOnly.TabIndex = 9
        Me.ChkQOnly.Text = "Q Only"
        Me.ChkQOnly.UseVisualStyleBackColor = True
        '
        'LblApproveOID
        '
        Me.LblApproveOID.AutoSize = True
        Me.LblApproveOID.Location = New System.Drawing.Point(61, 428)
        Me.LblApproveOID.Name = "LblApproveOID"
        Me.LblApproveOID.Size = New System.Drawing.Size(47, 13)
        Me.LblApproveOID.TabIndex = 10
        Me.LblApproveOID.TabStop = True
        Me.LblApproveOID.Text = "Approve"
        '
        'TxtSearch
        '
        Me.TxtSearch.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSearch.Location = New System.Drawing.Point(64, 477)
        Me.TxtSearch.Name = "TxtSearch"
        Me.TxtSearch.Size = New System.Drawing.Size(63, 20)
        Me.TxtSearch.TabIndex = 11
        '
        'LblSearch
        '
        Me.LblSearch.AutoSize = True
        Me.LblSearch.Location = New System.Drawing.Point(133, 479)
        Me.LblSearch.Name = "LblSearch"
        Me.LblSearch.Size = New System.Drawing.Size(41, 13)
        Me.LblSearch.TabIndex = 12
        Me.LblSearch.TabStop = True
        Me.LblSearch.Text = "Search"
        '
        'ChkAll
        '
        Me.ChkAll.AutoSize = True
        Me.ChkAll.Location = New System.Drawing.Point(6, 6)
        Me.ChkAll.Name = "ChkAll"
        Me.ChkAll.Size = New System.Drawing.Size(45, 17)
        Me.ChkAll.TabIndex = 23
        Me.ChkAll.Text = "ALL"
        Me.ChkAll.UseVisualStyleBackColor = True
        '
        'ChkAL
        '
        Me.ChkAL.CheckOnClick = True
        Me.ChkAL.ColumnWidth = 40
        Me.ChkAL.FormattingEnabled = True
        Me.ChkAL.Location = New System.Drawing.Point(3, 23)
        Me.ChkAL.MultiColumn = True
        Me.ChkAL.Name = "ChkAL"
        Me.ChkAL.Size = New System.Drawing.Size(300, 79)
        Me.ChkAL.TabIndex = 21
        '
        'LblUpdateChannel
        '
        Me.LblUpdateChannel.AutoSize = True
        Me.LblUpdateChannel.Location = New System.Drawing.Point(141, 9)
        Me.LblUpdateChannel.Name = "LblUpdateChannel"
        Me.LblUpdateChannel.Size = New System.Drawing.Size(84, 13)
        Me.LblUpdateChannel.TabIndex = 24
        Me.LblUpdateChannel.TabStop = True
        Me.LblUpdateChannel.Text = "Update Channel"
        Me.LblUpdateChannel.Visible = False
        '
        'LblUpdateAL
        '
        Me.LblUpdateAL.AutoSize = True
        Me.LblUpdateAL.Location = New System.Drawing.Point(244, 105)
        Me.LblUpdateAL.Name = "LblUpdateAL"
        Me.LblUpdateAL.Size = New System.Drawing.Size(58, 13)
        Me.LblUpdateAL.TabIndex = 25
        Me.LblUpdateAL.TabStop = True
        Me.LblUpdateAL.Text = "Update AL"
        Me.LblUpdateAL.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabAL)
        Me.TabControl1.Controls.Add(Me.TabChannel)
        Me.TabControl1.Controls.Add(Me.TabOID)
        Me.TabControl1.Controls.Add(Me.TabEmail)
        Me.TabControl1.Location = New System.Drawing.Point(477, 1)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(312, 472)
        Me.TabControl1.TabIndex = 26
        '
        'TabAL
        '
        Me.TabAL.Controls.Add(Me.ChkAL)
        Me.TabAL.Controls.Add(Me.LblUpdateAL)
        Me.TabAL.Controls.Add(Me.ChkAll)
        Me.TabAL.Location = New System.Drawing.Point(4, 22)
        Me.TabAL.Name = "TabAL"
        Me.TabAL.Padding = New System.Windows.Forms.Padding(3)
        Me.TabAL.Size = New System.Drawing.Size(304, 446)
        Me.TabAL.TabIndex = 0
        Me.TabAL.Text = "Airlines"
        Me.TabAL.UseVisualStyleBackColor = True
        '
        'TabChannel
        '
        Me.TabChannel.Controls.Add(Me.GrpCA)
        Me.TabChannel.Controls.Add(Me.CmbChannel)
        Me.TabChannel.Controls.Add(Me.LblUpdateChannel)
        Me.TabChannel.Controls.Add(Me.Label2)
        Me.TabChannel.Location = New System.Drawing.Point(4, 22)
        Me.TabChannel.Name = "TabChannel"
        Me.TabChannel.Padding = New System.Windows.Forms.Padding(3)
        Me.TabChannel.Size = New System.Drawing.Size(304, 446)
        Me.TabChannel.TabIndex = 1
        Me.TabChannel.Text = "Channel"
        Me.TabChannel.UseVisualStyleBackColor = True
        '
        'GrpCA
        '
        Me.GrpCA.Controls.Add(Me.CmbAL)
        Me.GrpCA.Controls.Add(Me.CmbSBU)
        Me.GrpCA.Controls.Add(Me.TxtCA)
        Me.GrpCA.Controls.Add(Me.LblDeleteCA)
        Me.GrpCA.Controls.Add(Me.LblUpdateCA)
        Me.GrpCA.Controls.Add(Me.Label10)
        Me.GrpCA.Controls.Add(Me.Label9)
        Me.GrpCA.Controls.Add(Me.Label8)
        Me.GrpCA.Controls.Add(Me.GridCA)
        Me.GrpCA.Location = New System.Drawing.Point(3, 69)
        Me.GrpCA.Name = "GrpCA"
        Me.GrpCA.Size = New System.Drawing.Size(298, 152)
        Me.GrpCA.TabIndex = 28
        Me.GrpCA.TabStop = False
        Me.GrpCA.Text = "CA Serverd by This TA"
        '
        'CmbAL
        '
        Me.CmbAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(104, 110)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(39, 21)
        Me.CmbAL.TabIndex = 30
        '
        'CmbSBU
        '
        Me.CmbSBU.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbSBU.FormattingEnabled = True
        Me.CmbSBU.Items.AddRange(New Object() {"EDU", "GSA", "TVS"})
        Me.CmbSBU.Location = New System.Drawing.Point(33, 110)
        Me.CmbSBU.Name = "CmbSBU"
        Me.CmbSBU.Size = New System.Drawing.Size(46, 21)
        Me.CmbSBU.TabIndex = 30
        '
        'TxtCA
        '
        Me.TxtCA.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCA.Location = New System.Drawing.Point(169, 110)
        Me.TxtCA.MaxLength = 8
        Me.TxtCA.Name = "TxtCA"
        Me.TxtCA.Size = New System.Drawing.Size(129, 20)
        Me.TxtCA.TabIndex = 29
        '
        'LblDeleteCA
        '
        Me.LblDeleteCA.AutoSize = True
        Me.LblDeleteCA.Location = New System.Drawing.Point(3, 136)
        Me.LblDeleteCA.Name = "LblDeleteCA"
        Me.LblDeleteCA.Size = New System.Drawing.Size(38, 13)
        Me.LblDeleteCA.TabIndex = 28
        Me.LblDeleteCA.TabStop = True
        Me.LblDeleteCA.Text = "Delete"
        Me.LblDeleteCA.Visible = False
        '
        'LblUpdateCA
        '
        Me.LblUpdateCA.AutoSize = True
        Me.LblUpdateCA.Location = New System.Drawing.Point(253, 136)
        Me.LblUpdateCA.Name = "LblUpdateCA"
        Me.LblUpdateCA.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdateCA.TabIndex = 28
        Me.LblUpdateCA.TabStop = True
        Me.LblUpdateCA.Text = "Update"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(148, 113)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(21, 13)
        Me.Label10.TabIndex = 4
        Me.Label10.Text = "CA"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(80, 113)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(20, 13)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "AL"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 113)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(29, 13)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "SBU"
        '
        'GridCA
        '
        Me.GridCA.AllowUserToAddRows = False
        Me.GridCA.AllowUserToDeleteRows = False
        Me.GridCA.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridCA.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCA.Location = New System.Drawing.Point(3, 19)
        Me.GridCA.Name = "GridCA"
        Me.GridCA.RowHeadersVisible = False
        Me.GridCA.Size = New System.Drawing.Size(295, 88)
        Me.GridCA.TabIndex = 27
        '
        'TabOID
        '
        Me.TabOID.Controls.Add(Me.LstOIDUsage)
        Me.TabOID.Controls.Add(Me.Label5)
        Me.TabOID.Controls.Add(Me.Label4)
        Me.TabOID.Controls.Add(Me.Label3)
        Me.TabOID.Controls.Add(Me.GridOffice)
        Me.TabOID.Controls.Add(Me.Label6)
        Me.TabOID.Controls.Add(Me.Label1)
        Me.TabOID.Controls.Add(Me.TxtOfficeID)
        Me.TabOID.Controls.Add(Me.LblApproveOID)
        Me.TabOID.Controls.Add(Me.ChkQOnly)
        Me.TabOID.Controls.Add(Me.LblUpdateOID)
        Me.TabOID.Controls.Add(Me.LblAddUsage)
        Me.TabOID.Controls.Add(Me.LblRemoveOID)
        Me.TabOID.Location = New System.Drawing.Point(4, 22)
        Me.TabOID.Name = "TabOID"
        Me.TabOID.Size = New System.Drawing.Size(304, 446)
        Me.TabOID.TabIndex = 2
        Me.TabOID.Text = "OfficeID"
        Me.TabOID.UseVisualStyleBackColor = True
        '
        'LstOIDUsage
        '
        Me.LstOIDUsage.CheckOnClick = True
        Me.LstOIDUsage.ColumnWidth = 48
        Me.LstOIDUsage.FormattingEnabled = True
        Me.LstOIDUsage.Location = New System.Drawing.Point(65, 29)
        Me.LstOIDUsage.MultiColumn = True
        Me.LstOIDUsage.Name = "LstOIDUsage"
        Me.LstOIDUsage.Size = New System.Drawing.Size(236, 19)
        Me.LstOIDUsage.TabIndex = 12
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(0, 407)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(154, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Use XXX for Unknown SI code"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(0, 393)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(284, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Sabre PCC: 3 Charactors followed by 3 Charactors SI code"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(0, 380)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(108, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "1A OID: 9 Charactors"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(2, 31)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "OID Usage"
        '
        'LblAddUsage
        '
        Me.LblAddUsage.AutoSize = True
        Me.LblAddUsage.Location = New System.Drawing.Point(145, 428)
        Me.LblAddUsage.Name = "LblAddUsage"
        Me.LblAddUsage.Size = New System.Drawing.Size(60, 13)
        Me.LblAddUsage.TabIndex = 8
        Me.LblAddUsage.TabStop = True
        Me.LblAddUsage.Text = "Add Usage"
        '
        'TabEmail
        '
        Me.TabEmail.Controls.Add(Me.LblUpdateEmail)
        Me.TabEmail.Controls.Add(Me.GridEmail)
        Me.TabEmail.Location = New System.Drawing.Point(4, 22)
        Me.TabEmail.Name = "TabEmail"
        Me.TabEmail.Size = New System.Drawing.Size(304, 446)
        Me.TabEmail.TabIndex = 3
        Me.TabEmail.Text = "Emails"
        Me.TabEmail.UseVisualStyleBackColor = True
        '
        'LblUpdateEmail
        '
        Me.LblUpdateEmail.AutoSize = True
        Me.LblUpdateEmail.Location = New System.Drawing.Point(234, 192)
        Me.LblUpdateEmail.Name = "LblUpdateEmail"
        Me.LblUpdateEmail.Size = New System.Drawing.Size(70, 13)
        Me.LblUpdateEmail.TabIndex = 1
        Me.LblUpdateEmail.TabStop = True
        Me.LblUpdateEmail.Text = "Update Email"
        '
        'GridEmail
        '
        Me.GridEmail.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridEmail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridEmail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.eMailType, Me.EmailAcct})
        Me.GridEmail.Location = New System.Drawing.Point(0, 0)
        Me.GridEmail.Name = "GridEmail"
        Me.GridEmail.RowHeadersWidth = 24
        Me.GridEmail.Size = New System.Drawing.Size(303, 189)
        Me.GridEmail.TabIndex = 0
        '
        'eMailType
        '
        Me.eMailType.HeaderText = "Dept"
        Me.eMailType.Name = "eMailType"
        Me.eMailType.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.eMailType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.eMailType.Width = 56
        '
        'EmailAcct
        '
        Me.EmailAcct.HeaderText = "Emails"
        Me.EmailAcct.Name = "EmailAcct"
        Me.EmailAcct.Width = 200
        '
        'CustDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.LblSearch)
        Me.Controls.Add(Me.TxtSearch)
        Me.Controls.Add(Me.CmbChannelFilter)
        Me.Controls.Add(Me.GridCustList)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CustDetail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Customer Details"
        CType(Me.GridCustList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridOffice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabAL.ResumeLayout(False)
        Me.TabAL.PerformLayout()
        Me.TabChannel.ResumeLayout(False)
        Me.TabChannel.PerformLayout()
        Me.GrpCA.ResumeLayout(False)
        Me.GrpCA.PerformLayout()
        CType(Me.GridCA, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabOID.ResumeLayout(False)
        Me.TabOID.PerformLayout()
        Me.TabEmail.ResumeLayout(False)
        Me.TabEmail.PerformLayout()
        CType(Me.GridEmail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridCustList As System.Windows.Forms.DataGridView
    Friend WithEvents CmbChannel As System.Windows.Forms.ComboBox
    Friend WithEvents LblUpdateOID As System.Windows.Forms.LinkLabel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GridOffice As System.Windows.Forms.DataGridView
    Friend WithEvents TxtOfficeID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CmbChannelFilter As System.Windows.Forms.ComboBox
    Friend WithEvents LblRemoveOID As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkQOnly As System.Windows.Forms.CheckBox
    Friend WithEvents LblApproveOID As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtSearch As System.Windows.Forms.TextBox
    Friend WithEvents LblSearch As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkAll As System.Windows.Forms.CheckBox
    Friend WithEvents ChkAL As System.Windows.Forms.CheckedListBox
    Friend WithEvents LblUpdateChannel As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUpdateAL As System.Windows.Forms.LinkLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabAL As System.Windows.Forms.TabPage
    Friend WithEvents TabChannel As System.Windows.Forms.TabPage
    Friend WithEvents TabOID As System.Windows.Forms.TabPage
    Friend WithEvents TabEmail As System.Windows.Forms.TabPage
    Friend WithEvents GridEmail As System.Windows.Forms.DataGridView
    Friend WithEvents LblUpdateEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents eMailType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents EmailAcct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LstOIDUsage As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LblAddUsage As System.Windows.Forms.LinkLabel
    Friend WithEvents GrpCA As System.Windows.Forms.GroupBox
    Friend WithEvents GridCA As System.Windows.Forms.DataGridView
    Friend WithEvents LblUpdateCA As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDeleteCA As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtCA As System.Windows.Forms.TextBox
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents CmbSBU As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
End Class
