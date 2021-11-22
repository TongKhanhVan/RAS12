<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChargeAndFee
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GridROE = New System.Windows.Forms.DataGridView()
        Me.CmbCurr = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CmdAdd = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtEffectFrom = New System.Windows.Forms.DateTimePicker()
        Me.txtFilterValue = New System.Windows.Forms.TextBox()
        Me.CmdShow = New System.Windows.Forms.Button()
        Me.CmdFilter = New System.Windows.Forms.Button()
        Me.CmdMark = New System.Windows.Forms.Button()
        Me.CmbAL = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtChargeName = New System.Windows.Forms.TextBox()
        Me.CmbChargeOn = New System.Windows.Forms.ComboBox()
        Me.txtEffectThru = New System.Windows.Forms.DateTimePicker()
        Me.ChkNonRefundable = New System.Windows.Forms.CheckBox()
        Me.ChkActive = New System.Windows.Forms.CheckBox()
        Me.OptTV = New System.Windows.Forms.RadioButton()
        Me.OptAL = New System.Windows.Forms.RadioButton()
        Me.CmbTRX = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.CmbSBU = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.CmbChannel = New System.Windows.Forms.ComboBox()
        Me.CmbAmtType = New System.Windows.Forms.ComboBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.LblExtend = New System.Windows.Forms.LinkLabel()
        Me.txtNewValidTo = New System.Windows.Forms.DateTimePicker()
        Me.txtCurrentValidTo = New System.Windows.Forms.DateTimePicker()
        Me.LstAL_ext = New System.Windows.Forms.CheckedListBox()
        Me.CmbLevel_Ext = New System.Windows.Forms.ComboBox()
        Me.CmbTA_TO_Ext = New System.Windows.Forms.ComboBox()
        Me.CmbSBU_ext = New System.Windows.Forms.ComboBox()
        CType(Me.GridROE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridROE
        '
        Me.GridROE.AllowUserToAddRows = False
        Me.GridROE.AllowUserToDeleteRows = False
        Me.GridROE.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridROE.BackgroundColor = System.Drawing.Color.LemonChiffon
        Me.GridROE.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridROE.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.GridROE.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.GridROE.DefaultCellStyle = DataGridViewCellStyle2
        Me.GridROE.Location = New System.Drawing.Point(3, 75)
        Me.GridROE.Name = "GridROE"
        Me.GridROE.RowHeadersWidth = 25
        Me.GridROE.Size = New System.Drawing.Size(771, 425)
        Me.GridROE.TabIndex = 4
        '
        'CmbCurr
        '
        Me.CmbCurr.FormattingEnabled = True
        Me.CmbCurr.Location = New System.Drawing.Point(121, 6)
        Me.CmbCurr.Name = "CmbCurr"
        Me.CmbCurr.Size = New System.Drawing.Size(56, 21)
        Me.CmbCurr.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(75, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Currency"
        '
        'CmdAdd
        '
        Me.CmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdAdd.Location = New System.Drawing.Point(678, 47)
        Me.CmdAdd.Name = "CmdAdd"
        Me.CmdAdd.Size = New System.Drawing.Size(51, 22)
        Me.CmdAdd.TabIndex = 7
        Me.CmdAdd.Text = "Add"
        Me.CmdAdd.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(183, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "ChargeName"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(336, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "EffectThru"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(477, 31)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "ChargeOn"
        '
        'txtAmount
        '
        Me.txtAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAmount.Location = New System.Drawing.Point(393, 8)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(78, 20)
        Me.txtAmount.TabIndex = 3
        Me.txtAmount.Text = "0"
        Me.txtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(336, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(43, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Amount"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(183, 31)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(58, 13)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "EffectFrom"
        '
        'txtEffectFrom
        '
        Me.txtEffectFrom.CustomFormat = "dd-MMM-yy"
        Me.txtEffectFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtEffectFrom.Location = New System.Drawing.Point(256, 27)
        Me.txtEffectFrom.Name = "txtEffectFrom"
        Me.txtEffectFrom.Size = New System.Drawing.Size(79, 20)
        Me.txtEffectFrom.TabIndex = 5
        '
        'txtFilterValue
        '
        Me.txtFilterValue.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtFilterValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFilterValue.Location = New System.Drawing.Point(2, 508)
        Me.txtFilterValue.Name = "txtFilterValue"
        Me.txtFilterValue.Size = New System.Drawing.Size(54, 20)
        Me.txtFilterValue.TabIndex = 10
        '
        'CmdShow
        '
        Me.CmdShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdShow.Location = New System.Drawing.Point(126, 507)
        Me.CmdShow.Name = "CmdShow"
        Me.CmdShow.Size = New System.Drawing.Size(83, 22)
        Me.CmdShow.TabIndex = 8
        Me.CmdShow.Text = "ShowAll"
        Me.CmdShow.UseVisualStyleBackColor = True
        '
        'CmdFilter
        '
        Me.CmdFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdFilter.Location = New System.Drawing.Point(62, 507)
        Me.CmdFilter.Name = "CmdFilter"
        Me.CmdFilter.Size = New System.Drawing.Size(58, 22)
        Me.CmdFilter.TabIndex = 8
        Me.CmdFilter.Text = "Filter"
        Me.CmdFilter.UseVisualStyleBackColor = True
        '
        'CmdMark
        '
        Me.CmdMark.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmdMark.Location = New System.Drawing.Point(674, 506)
        Me.CmdMark.Name = "CmdMark"
        Me.CmdMark.Size = New System.Drawing.Size(100, 22)
        Me.CmdMark.TabIndex = 8
        Me.CmdMark.Text = "Mark As InActive"
        Me.CmdMark.UseVisualStyleBackColor = True
        '
        'CmbAL
        '
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(28, 6)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(42, 21)
        Me.CmbAL.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(20, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "AL"
        '
        'txtChargeName
        '
        Me.txtChargeName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtChargeName.Location = New System.Drawing.Point(256, 6)
        Me.txtChargeName.MaxLength = 6
        Me.txtChargeName.Name = "txtChargeName"
        Me.txtChargeName.Size = New System.Drawing.Size(79, 20)
        Me.txtChargeName.TabIndex = 2
        '
        'CmbChargeOn
        '
        Me.CmbChargeOn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbChargeOn.Enabled = False
        Me.CmbChargeOn.FormattingEnabled = True
        Me.CmbChargeOn.Items.AddRange(New Object() {"TKT", "SEG"})
        Me.CmbChargeOn.Location = New System.Drawing.Point(538, 26)
        Me.CmbChargeOn.Name = "CmbChargeOn"
        Me.CmbChargeOn.Size = New System.Drawing.Size(51, 21)
        Me.CmbChargeOn.TabIndex = 4
        '
        'txtEffectThru
        '
        Me.txtEffectThru.CustomFormat = "dd-MMM-yy"
        Me.txtEffectThru.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtEffectThru.Location = New System.Drawing.Point(393, 28)
        Me.txtEffectThru.Name = "txtEffectThru"
        Me.txtEffectThru.Size = New System.Drawing.Size(78, 20)
        Me.txtEffectThru.TabIndex = 6
        '
        'ChkNonRefundable
        '
        Me.ChkNonRefundable.AutoSize = True
        Me.ChkNonRefundable.Checked = True
        Me.ChkNonRefundable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkNonRefundable.Location = New System.Drawing.Point(678, 9)
        Me.ChkNonRefundable.Name = "ChkNonRefundable"
        Me.ChkNonRefundable.Size = New System.Drawing.Size(81, 17)
        Me.ChkNonRefundable.TabIndex = 11
        Me.ChkNonRefundable.Text = "NonRefund"
        Me.ChkNonRefundable.UseVisualStyleBackColor = True
        '
        'ChkActive
        '
        Me.ChkActive.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkActive.AutoSize = True
        Me.ChkActive.Checked = True
        Me.ChkActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkActive.Location = New System.Drawing.Point(215, 510)
        Me.ChkActive.Name = "ChkActive"
        Me.ChkActive.Size = New System.Drawing.Size(77, 17)
        Me.ChkActive.TabIndex = 12
        Me.ChkActive.Text = "ActiveOnly"
        Me.ChkActive.UseVisualStyleBackColor = True
        '
        'OptTV
        '
        Me.OptTV.AutoSize = True
        Me.OptTV.Enabled = False
        Me.OptTV.Location = New System.Drawing.Point(596, 9)
        Me.OptTV.Name = "OptTV"
        Me.OptTV.Size = New System.Drawing.Size(76, 17)
        Me.OptTV.TabIndex = 13
        Me.OptTV.Text = "TV Charge"
        Me.OptTV.UseVisualStyleBackColor = True
        '
        'OptAL
        '
        Me.OptAL.AutoSize = True
        Me.OptAL.Checked = True
        Me.OptAL.Location = New System.Drawing.Point(596, 27)
        Me.OptAL.Name = "OptAL"
        Me.OptAL.Size = New System.Drawing.Size(75, 17)
        Me.OptAL.TabIndex = 13
        Me.OptAL.TabStop = True
        Me.OptAL.Text = "AL Charge"
        Me.OptAL.UseVisualStyleBackColor = True
        '
        'CmbTRX
        '
        Me.CmbTRX.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbTRX.FormattingEnabled = True
        Me.CmbTRX.Items.AddRange(New Object() {"S - Sales, Void, ", "R - Refund, Void RECRD", "A - ALL"})
        Me.CmbTRX.Location = New System.Drawing.Point(69, 28)
        Me.CmbTRX.Name = "CmbTRX"
        Me.CmbTRX.Size = New System.Drawing.Size(108, 21)
        Me.CmbTRX.TabIndex = 14
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Enabled = False
        Me.Label8.Location = New System.Drawing.Point(6, 31)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 13)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "TRX Type"
        '
        'CmbSBU
        '
        Me.CmbSBU.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbSBU.FormattingEnabled = True
        Me.CmbSBU.Items.AddRange(New Object() {"EDU", "GSA", "TVS"})
        Me.CmbSBU.Location = New System.Drawing.Point(69, 49)
        Me.CmbSBU.Name = "CmbSBU"
        Me.CmbSBU.Size = New System.Drawing.Size(74, 21)
        Me.CmbSBU.TabIndex = 16
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 52)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(29, 13)
        Me.Label9.TabIndex = 17
        Me.Label9.Text = "SBU"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(183, 52)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(46, 13)
        Me.Label10.TabIndex = 19
        Me.Label10.Text = "Channel"
        '
        'CmbChannel
        '
        Me.CmbChannel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbChannel.FormattingEnabled = True
        Me.CmbChannel.Items.AddRange(New Object() {"CA", "CS", "TA", "TO", "??"})
        Me.CmbChannel.Location = New System.Drawing.Point(256, 49)
        Me.CmbChannel.Name = "CmbChannel"
        Me.CmbChannel.Size = New System.Drawing.Size(79, 21)
        Me.CmbChannel.TabIndex = 18
        '
        'CmbAmtType
        '
        Me.CmbAmtType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAmtType.FormattingEnabled = True
        Me.CmbAmtType.Items.AddRange(New Object() {"PCT", "USD"})
        Me.CmbAmtType.Location = New System.Drawing.Point(475, 7)
        Me.CmbAmtType.Name = "CmbAmtType"
        Me.CmbAmtType.Size = New System.Drawing.Size(51, 21)
        Me.CmbAmtType.TabIndex = 20
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(2, 3)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(785, 558)
        Me.TabControl1.TabIndex = 21
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GridROE)
        Me.TabPage1.Controls.Add(Me.Label9)
        Me.TabPage1.Controls.Add(Me.CmbAmtType)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.CmdShow)
        Me.TabPage1.Controls.Add(Me.Label10)
        Me.TabPage1.Controls.Add(Me.CmdFilter)
        Me.TabPage1.Controls.Add(Me.CmbChannel)
        Me.TabPage1.Controls.Add(Me.CmdMark)
        Me.TabPage1.Controls.Add(Me.txtFilterValue)
        Me.TabPage1.Controls.Add(Me.CmbSBU)
        Me.TabPage1.Controls.Add(Me.ChkActive)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.txtChargeName)
        Me.TabPage1.Controls.Add(Me.CmbTRX)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.OptAL)
        Me.TabPage1.Controls.Add(Me.CmbCurr)
        Me.TabPage1.Controls.Add(Me.OptTV)
        Me.TabPage1.Controls.Add(Me.CmbChargeOn)
        Me.TabPage1.Controls.Add(Me.ChkNonRefundable)
        Me.TabPage1.Controls.Add(Me.CmbAL)
        Me.TabPage1.Controls.Add(Me.txtEffectThru)
        Me.TabPage1.Controls.Add(Me.txtAmount)
        Me.TabPage1.Controls.Add(Me.txtEffectFrom)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.CmdAdd)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.Label7)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(777, 532)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Update"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label11)
        Me.TabPage2.Controls.Add(Me.LblExtend)
        Me.TabPage2.Controls.Add(Me.txtNewValidTo)
        Me.TabPage2.Controls.Add(Me.txtCurrentValidTo)
        Me.TabPage2.Controls.Add(Me.LstAL_ext)
        Me.TabPage2.Controls.Add(Me.CmbLevel_Ext)
        Me.TabPage2.Controls.Add(Me.CmbTA_TO_Ext)
        Me.TabPage2.Controls.Add(Me.CmbSBU_ext)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(777, 532)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Bulk Extension"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(229, 9)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(99, 13)
        Me.Label11.TabIndex = 4
        Me.Label11.Text = "Currently Valid Thru"
        '
        'LblExtend
        '
        Me.LblExtend.AutoSize = True
        Me.LblExtend.Location = New System.Drawing.Point(624, 9)
        Me.LblExtend.Name = "LblExtend"
        Me.LblExtend.Size = New System.Drawing.Size(56, 13)
        Me.LblExtend.TabIndex = 3
        Me.LblExtend.TabStop = True
        Me.LblExtend.Text = "Extend To"
        '
        'txtNewValidTo
        '
        Me.txtNewValidTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtNewValidTo.Location = New System.Drawing.Point(688, 7)
        Me.txtNewValidTo.Name = "txtNewValidTo"
        Me.txtNewValidTo.Size = New System.Drawing.Size(82, 20)
        Me.txtNewValidTo.TabIndex = 2
        '
        'txtCurrentValidTo
        '
        Me.txtCurrentValidTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtCurrentValidTo.Location = New System.Drawing.Point(334, 6)
        Me.txtCurrentValidTo.Name = "txtCurrentValidTo"
        Me.txtCurrentValidTo.Size = New System.Drawing.Size(82, 20)
        Me.txtCurrentValidTo.TabIndex = 2
        '
        'LstAL_ext
        '
        Me.LstAL_ext.CheckOnClick = True
        Me.LstAL_ext.ColumnWidth = 36
        Me.LstAL_ext.FormattingEnabled = True
        Me.LstAL_ext.Location = New System.Drawing.Point(6, 33)
        Me.LstAL_ext.MultiColumn = True
        Me.LstAL_ext.Name = "LstAL_ext"
        Me.LstAL_ext.Size = New System.Drawing.Size(410, 94)
        Me.LstAL_ext.TabIndex = 1
        '
        'CmbLevel_Ext
        '
        Me.CmbLevel_Ext.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbLevel_Ext.FormattingEnabled = True
        Me.CmbLevel_Ext.Location = New System.Drawing.Point(129, 6)
        Me.CmbLevel_Ext.Name = "CmbLevel_Ext"
        Me.CmbLevel_Ext.Size = New System.Drawing.Size(96, 21)
        Me.CmbLevel_Ext.TabIndex = 0
        '
        'CmbTA_TO_Ext
        '
        Me.CmbTA_TO_Ext.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbTA_TO_Ext.FormattingEnabled = True
        Me.CmbTA_TO_Ext.Items.AddRange(New Object() {"TA", "TO", "CA"})
        Me.CmbTA_TO_Ext.Location = New System.Drawing.Point(65, 6)
        Me.CmbTA_TO_Ext.Name = "CmbTA_TO_Ext"
        Me.CmbTA_TO_Ext.Size = New System.Drawing.Size(60, 21)
        Me.CmbTA_TO_Ext.TabIndex = 0
        '
        'CmbSBU_ext
        '
        Me.CmbSBU_ext.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbSBU_ext.FormattingEnabled = True
        Me.CmbSBU_ext.Items.AddRange(New Object() {"GSA", "TVS"})
        Me.CmbSBU_ext.Location = New System.Drawing.Point(6, 6)
        Me.CmbSBU_ext.Name = "CmbSBU_ext"
        Me.CmbSBU_ext.Size = New System.Drawing.Size(56, 21)
        Me.CmbSBU_ext.TabIndex = 0
        '
        'ChargeAndFee
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ChargeAndFee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Airlines Charges "
        CType(Me.GridROE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GridROE As System.Windows.Forms.DataGridView
    Friend WithEvents CmbCurr As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CmdAdd As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtEffectFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtFilterValue As System.Windows.Forms.TextBox
    Friend WithEvents CmdShow As System.Windows.Forms.Button
    Friend WithEvents CmdFilter As System.Windows.Forms.Button
    Friend WithEvents CmdMark As System.Windows.Forms.Button
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtChargeName As System.Windows.Forms.TextBox
    Friend WithEvents CmbChargeOn As System.Windows.Forms.ComboBox
    Friend WithEvents txtEffectThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents ChkNonRefundable As System.Windows.Forms.CheckBox
    Friend WithEvents ChkActive As System.Windows.Forms.CheckBox
    Friend WithEvents OptTV As System.Windows.Forms.RadioButton
    Friend WithEvents OptAL As System.Windows.Forms.RadioButton
    Friend WithEvents CmbTRX As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CmbSBU As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CmbChannel As System.Windows.Forms.ComboBox
    Friend WithEvents CmbAmtType As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents LblExtend As System.Windows.Forms.LinkLabel
    Friend WithEvents txtNewValidTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtCurrentValidTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents LstAL_ext As System.Windows.Forms.CheckedListBox
    Friend WithEvents CmbLevel_Ext As System.Windows.Forms.ComboBox
    Friend WithEvents CmbTA_TO_Ext As System.Windows.Forms.ComboBox
    Friend WithEvents CmbSBU_ext As System.Windows.Forms.ComboBox
End Class
