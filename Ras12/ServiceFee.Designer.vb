<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ServiceFee
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
        Me.CmbSBU = New System.Windows.Forms.ComboBox()
        Me.CmbAL = New System.Windows.Forms.ComboBox()
        Me.CmbTRXtype = New System.Windows.Forms.ComboBox()
        Me.CmbChannel = New System.Windows.Forms.ComboBox()
        Me.CmbBase = New System.Windows.Forms.ComboBox()
        Me.TxtValidFrom = New System.Windows.Forms.DateTimePicker()
        Me.TxtValidThru = New System.Windows.Forms.DateTimePicker()
        Me.TxtVAL = New System.Windows.Forms.TextBox()
        Me.TxtFareThru = New System.Windows.Forms.TextBox()
        Me.TxtfareFrm = New System.Windows.Forms.TextBox()
        Me.TxtMin = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TxtMax = New System.Windows.Forms.TextBox()
        Me.LblUpdateSF = New System.Windows.Forms.LinkLabel()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.LckLblApproveSF = New System.Windows.Forms.LinkLabel()
        Me.GridSF = New System.Windows.Forms.DataGridView()
        Me.CmbSFType = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.LblPCT = New System.Windows.Forms.Label()
        Me.ChkNonRefundable = New System.Windows.Forms.CheckBox()
        Me.CmbLevel = New System.Windows.Forms.ComboBox()
        Me.ChkQOnly = New System.Windows.Forms.CheckBox()
        Me.ChkActiveOnly = New System.Windows.Forms.CheckBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabSF = New System.Windows.Forms.TabPage()
        Me.TxtDest = New System.Windows.Forms.TextBox()
        Me.CmbISI = New System.Windows.Forms.ComboBox()
        Me.CmbFType = New System.Windows.Forms.ComboBox()
        Me.CmbArea = New System.Windows.Forms.ComboBox()
        Me.ChkINF = New System.Windows.Forms.CheckBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.TabDiscount = New System.Windows.Forms.TabPage()
        Me.CmbDiscBase = New System.Windows.Forms.ComboBox()
        Me.LblUpdateDisc = New System.Windows.Forms.LinkLabel()
        Me.GridDisc = New System.Windows.Forms.DataGridView()
        Me.TxtCOC = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.CmbFSource = New System.Windows.Forms.ComboBox()
        Me.LblFSource = New System.Windows.Forms.Label()
        Me.LblShowAll = New System.Windows.Forms.LinkLabel()
        Me.LblFilter = New System.Windows.Forms.LinkLabel()
        Me.TxtWhat = New System.Windows.Forms.TextBox()
        Me.LblChangeValidity = New System.Windows.Forms.LinkLabel()
        Me.LblViewXpireDate = New System.Windows.Forms.LinkLabel()
        Me.GridExpireDate = New System.Windows.Forms.DataGridView()
        CType(Me.GridSF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabSF.SuspendLayout()
        Me.TabDiscount.SuspendLayout()
        CType(Me.GridDisc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridExpireDate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CmbSBU
        '
        Me.CmbSBU.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbSBU.FormattingEnabled = True
        Me.CmbSBU.Items.AddRange(New Object() {"TVS", "GSA", "EDU"})
        Me.CmbSBU.Location = New System.Drawing.Point(37, 5)
        Me.CmbSBU.Name = "CmbSBU"
        Me.CmbSBU.Size = New System.Drawing.Size(50, 21)
        Me.CmbSBU.TabIndex = 0
        '
        'CmbAL
        '
        Me.CmbAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(110, 6)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(43, 21)
        Me.CmbAL.TabIndex = 1
        '
        'CmbTRXtype
        '
        Me.CmbTRXtype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbTRXtype.FormattingEnabled = True
        Me.CmbTRXtype.Items.AddRange(New Object() {"S", "R", "V", "I"})
        Me.CmbTRXtype.Location = New System.Drawing.Point(53, 7)
        Me.CmbTRXtype.Name = "CmbTRXtype"
        Me.CmbTRXtype.Size = New System.Drawing.Size(49, 21)
        Me.CmbTRXtype.TabIndex = 8
        '
        'CmbChannel
        '
        Me.CmbChannel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbChannel.FormattingEnabled = True
        Me.CmbChannel.Location = New System.Drawing.Point(234, 6)
        Me.CmbChannel.Name = "CmbChannel"
        Me.CmbChannel.Size = New System.Drawing.Size(44, 21)
        Me.CmbChannel.TabIndex = 2
        '
        'CmbBase
        '
        Me.CmbBase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbBase.Enabled = False
        Me.CmbBase.FormattingEnabled = True
        Me.CmbBase.Items.AddRange(New Object() {"TKT", "SEG"})
        Me.CmbBase.Location = New System.Drawing.Point(140, 7)
        Me.CmbBase.Name = "CmbBase"
        Me.CmbBase.Size = New System.Drawing.Size(49, 21)
        Me.CmbBase.TabIndex = 9
        '
        'TxtValidFrom
        '
        Me.TxtValidFrom.CustomFormat = "dd-MMM-yy"
        Me.TxtValidFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TxtValidFrom.Location = New System.Drawing.Point(597, 5)
        Me.TxtValidFrom.Name = "TxtValidFrom"
        Me.TxtValidFrom.Size = New System.Drawing.Size(78, 20)
        Me.TxtValidFrom.TabIndex = 6
        '
        'TxtValidThru
        '
        Me.TxtValidThru.CustomFormat = "dd-MMM-yy"
        Me.TxtValidThru.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TxtValidThru.Location = New System.Drawing.Point(679, 5)
        Me.TxtValidThru.Name = "TxtValidThru"
        Me.TxtValidThru.Size = New System.Drawing.Size(78, 20)
        Me.TxtValidThru.TabIndex = 7
        '
        'TxtVAL
        '
        Me.TxtVAL.Location = New System.Drawing.Point(453, 6)
        Me.TxtVAL.Name = "TxtVAL"
        Me.TxtVAL.Size = New System.Drawing.Size(44, 20)
        Me.TxtVAL.TabIndex = 5
        Me.TxtVAL.Text = "0"
        Me.TxtVAL.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtFareThru
        '
        Me.TxtFareThru.Location = New System.Drawing.Point(299, 6)
        Me.TxtFareThru.Name = "TxtFareThru"
        Me.TxtFareThru.Size = New System.Drawing.Size(61, 20)
        Me.TxtFareThru.TabIndex = 11
        Me.TxtFareThru.Text = "0"
        Me.TxtFareThru.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtfareFrm
        '
        Me.TxtfareFrm.Location = New System.Drawing.Point(253, 6)
        Me.TxtfareFrm.Name = "TxtfareFrm"
        Me.TxtfareFrm.Size = New System.Drawing.Size(45, 20)
        Me.TxtfareFrm.TabIndex = 10
        Me.TxtfareFrm.Text = "0"
        Me.TxtfareFrm.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMin
        '
        Me.TxtMin.Enabled = False
        Me.TxtMin.Location = New System.Drawing.Point(398, 6)
        Me.TxtMin.Name = "TxtMin"
        Me.TxtMin.Size = New System.Drawing.Size(49, 20)
        Me.TxtMin.TabIndex = 12
        Me.TxtMin.Text = "0"
        Me.TxtMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "SBU"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(-2, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "TRX Type"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(155, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Channel/Level"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(90, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "AL"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(103, 10)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(31, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Base"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(191, 10)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "FareRange"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(513, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(80, 13)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Valid from/Thru"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(410, 9)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 13)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Amount"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(371, 9)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(24, 13)
        Me.Label10.TabIndex = 10
        Me.Label10.Text = "Min"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(450, 9)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(27, 13)
        Me.Label11.TabIndex = 10
        Me.Label11.Text = "Max"
        '
        'TxtMax
        '
        Me.TxtMax.Enabled = False
        Me.TxtMax.Location = New System.Drawing.Point(477, 6)
        Me.TxtMax.Name = "TxtMax"
        Me.TxtMax.Size = New System.Drawing.Size(44, 20)
        Me.TxtMax.TabIndex = 13
        Me.TxtMax.Text = "0"
        Me.TxtMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblUpdateSF
        '
        Me.LblUpdateSF.AutoSize = True
        Me.LblUpdateSF.Location = New System.Drawing.Point(719, 33)
        Me.LblUpdateSF.Name = "LblUpdateSF"
        Me.LblUpdateSF.Size = New System.Drawing.Size(58, 13)
        Me.LblUpdateSF.TabIndex = 20
        Me.LblUpdateSF.TabStop = True
        Me.LblUpdateSF.Text = "Update SF"
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(79, 482)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 18
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        Me.LblDelete.Visible = False
        '
        'LckLblApproveSF
        '
        Me.LckLblApproveSF.AutoSize = True
        Me.LckLblApproveSF.Location = New System.Drawing.Point(737, 482)
        Me.LckLblApproveSF.Name = "LckLblApproveSF"
        Me.LckLblApproveSF.Size = New System.Drawing.Size(47, 13)
        Me.LckLblApproveSF.TabIndex = 19
        Me.LckLblApproveSF.TabStop = True
        Me.LckLblApproveSF.Text = "Approve"
        Me.LckLblApproveSF.Visible = False
        '
        'GridSF
        '
        Me.GridSF.AllowUserToAddRows = False
        Me.GridSF.AllowUserToDeleteRows = False
        Me.GridSF.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridSF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridSF.Location = New System.Drawing.Point(0, 56)
        Me.GridSF.Name = "GridSF"
        Me.GridSF.RowHeadersVisible = False
        Me.GridSF.Size = New System.Drawing.Size(783, 343)
        Me.GridSF.TabIndex = 16
        '
        'CmbSFType
        '
        Me.CmbSFType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbSFType.FormattingEnabled = True
        Me.CmbSFType.Items.AddRange(New Object() {"PCT", "VAL"})
        Me.CmbSFType.Location = New System.Drawing.Point(365, 6)
        Me.CmbSFType.Name = "CmbSFType"
        Me.CmbSFType.Size = New System.Drawing.Size(44, 21)
        Me.CmbSFType.TabIndex = 4
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(334, 9)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(31, 13)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Type"
        '
        'LblPCT
        '
        Me.LblPCT.AutoSize = True
        Me.LblPCT.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.LblPCT.Location = New System.Drawing.Point(494, 9)
        Me.LblPCT.Name = "LblPCT"
        Me.LblPCT.Size = New System.Drawing.Size(15, 13)
        Me.LblPCT.TabIndex = 18
        Me.LblPCT.Text = "%"
        '
        'ChkNonRefundable
        '
        Me.ChkNonRefundable.AutoSize = True
        Me.ChkNonRefundable.Checked = True
        Me.ChkNonRefundable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkNonRefundable.Location = New System.Drawing.Point(528, 9)
        Me.ChkNonRefundable.Name = "ChkNonRefundable"
        Me.ChkNonRefundable.Size = New System.Drawing.Size(101, 17)
        Me.ChkNonRefundable.TabIndex = 14
        Me.ChkNonRefundable.Text = "NonRefundable"
        Me.ChkNonRefundable.UseVisualStyleBackColor = True
        '
        'CmbLevel
        '
        Me.CmbLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbLevel.FormattingEnabled = True
        Me.CmbLevel.Location = New System.Drawing.Point(280, 6)
        Me.CmbLevel.Name = "CmbLevel"
        Me.CmbLevel.Size = New System.Drawing.Size(55, 21)
        Me.CmbLevel.TabIndex = 3
        '
        'ChkQOnly
        '
        Me.ChkQOnly.AutoSize = True
        Me.ChkQOnly.Location = New System.Drawing.Point(679, 481)
        Me.ChkQOnly.Name = "ChkQOnly"
        Me.ChkQOnly.Size = New System.Drawing.Size(55, 17)
        Me.ChkQOnly.TabIndex = 20
        Me.ChkQOnly.Text = "QOnly"
        Me.ChkQOnly.UseVisualStyleBackColor = True
        '
        'ChkActiveOnly
        '
        Me.ChkActiveOnly.AutoSize = True
        Me.ChkActiveOnly.Location = New System.Drawing.Point(4, 481)
        Me.ChkActiveOnly.Name = "ChkActiveOnly"
        Me.ChkActiveOnly.Size = New System.Drawing.Size(77, 17)
        Me.ChkActiveOnly.TabIndex = 17
        Me.ChkActiveOnly.Text = "ActiveOnly"
        Me.ChkActiveOnly.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabSF)
        Me.TabControl1.Controls.Add(Me.TabDiscount)
        Me.TabControl1.Location = New System.Drawing.Point(0, 52)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(791, 427)
        Me.TabControl1.TabIndex = 21
        '
        'TabSF
        '
        Me.TabSF.Controls.Add(Me.TxtDest)
        Me.TabSF.Controls.Add(Me.CmbISI)
        Me.TabSF.Controls.Add(Me.CmbFType)
        Me.TabSF.Controls.Add(Me.CmbTRXtype)
        Me.TabSF.Controls.Add(Me.Label2)
        Me.TabSF.Controls.Add(Me.CmbArea)
        Me.TabSF.Controls.Add(Me.CmbBase)
        Me.TabSF.Controls.Add(Me.ChkINF)
        Me.TabSF.Controls.Add(Me.ChkNonRefundable)
        Me.TabSF.Controls.Add(Me.Label18)
        Me.TabSF.Controls.Add(Me.Label14)
        Me.TabSF.Controls.Add(Me.Label12)
        Me.TabSF.Controls.Add(Me.Label5)
        Me.TabSF.Controls.Add(Me.GridSF)
        Me.TabSF.Controls.Add(Me.TxtMin)
        Me.TabSF.Controls.Add(Me.TxtMax)
        Me.TabSF.Controls.Add(Me.Label10)
        Me.TabSF.Controls.Add(Me.LblUpdateSF)
        Me.TabSF.Controls.Add(Me.Label11)
        Me.TabSF.Controls.Add(Me.TxtFareThru)
        Me.TabSF.Controls.Add(Me.TxtfareFrm)
        Me.TabSF.Controls.Add(Me.Label17)
        Me.TabSF.Controls.Add(Me.Label6)
        Me.TabSF.Location = New System.Drawing.Point(4, 22)
        Me.TabSF.Name = "TabSF"
        Me.TabSF.Padding = New System.Windows.Forms.Padding(3)
        Me.TabSF.Size = New System.Drawing.Size(783, 401)
        Me.TabSF.TabIndex = 0
        Me.TabSF.Text = "ServiceFee"
        Me.TabSF.UseVisualStyleBackColor = True
        '
        'TxtDest
        '
        Me.TxtDest.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDest.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.TxtDest.Location = New System.Drawing.Point(477, 30)
        Me.TxtDest.Name = "TxtDest"
        Me.TxtDest.Size = New System.Drawing.Size(218, 20)
        Me.TxtDest.TabIndex = 19
        Me.TxtDest.Text = "--COMMA SEPARATED COUNTRY CODE--"
        '
        'CmbISI
        '
        Me.CmbISI.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbISI.FormattingEnabled = True
        Me.CmbISI.Items.AddRange(New Object() {"ALL", "SIT*", "SOT*"})
        Me.CmbISI.Location = New System.Drawing.Point(140, 30)
        Me.CmbISI.Name = "CmbISI"
        Me.CmbISI.Size = New System.Drawing.Size(49, 21)
        Me.CmbISI.TabIndex = 17
        '
        'CmbFType
        '
        Me.CmbFType.FormattingEnabled = True
        Me.CmbFType.Location = New System.Drawing.Point(253, 30)
        Me.CmbFType.Name = "CmbFType"
        Me.CmbFType.Size = New System.Drawing.Size(45, 21)
        Me.CmbFType.TabIndex = 18
        '
        'CmbArea
        '
        Me.CmbArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbArea.FormattingEnabled = True
        Me.CmbArea.Items.AddRange(New Object() {"ALL", "DOM", "INT"})
        Me.CmbArea.Location = New System.Drawing.Point(53, 30)
        Me.CmbArea.Name = "CmbArea"
        Me.CmbArea.Size = New System.Drawing.Size(49, 21)
        Me.CmbArea.TabIndex = 9
        '
        'ChkINF
        '
        Me.ChkINF.AutoSize = True
        Me.ChkINF.Location = New System.Drawing.Point(635, 9)
        Me.ChkINF.Name = "ChkINF"
        Me.ChkINF.Size = New System.Drawing.Size(95, 17)
        Me.ChkINF.TabIndex = 14
        Me.ChkINF.Text = "INF Applicable"
        Me.ChkINF.UseVisualStyleBackColor = True
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(103, 33)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(20, 13)
        Me.Label18.TabIndex = 8
        Me.Label18.Text = "ISI"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(3, 33)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(29, 13)
        Me.Label14.TabIndex = 8
        Me.Label14.Text = "Area"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(191, 33)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(51, 13)
        Me.Label12.TabIndex = 8
        Me.Label12.Text = "GRP/FIT"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(392, 33)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(79, 13)
        Me.Label17.TabIndex = 9
        Me.Label17.Text = "Dest. Countries"
        '
        'TabDiscount
        '
        Me.TabDiscount.Controls.Add(Me.CmbDiscBase)
        Me.TabDiscount.Controls.Add(Me.LblUpdateDisc)
        Me.TabDiscount.Controls.Add(Me.GridDisc)
        Me.TabDiscount.Controls.Add(Me.TxtCOC)
        Me.TabDiscount.Controls.Add(Me.Label13)
        Me.TabDiscount.Controls.Add(Me.Label16)
        Me.TabDiscount.Location = New System.Drawing.Point(4, 22)
        Me.TabDiscount.Name = "TabDiscount"
        Me.TabDiscount.Padding = New System.Windows.Forms.Padding(3)
        Me.TabDiscount.Size = New System.Drawing.Size(783, 401)
        Me.TabDiscount.TabIndex = 1
        Me.TabDiscount.Text = "Discount"
        Me.TabDiscount.UseVisualStyleBackColor = True
        '
        'CmbDiscBase
        '
        Me.CmbDiscBase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbDiscBase.FormattingEnabled = True
        Me.CmbDiscBase.Location = New System.Drawing.Point(150, 6)
        Me.CmbDiscBase.Name = "CmbDiscBase"
        Me.CmbDiscBase.Size = New System.Drawing.Size(52, 21)
        Me.CmbDiscBase.TabIndex = 11
        '
        'LblUpdateDisc
        '
        Me.LblUpdateDisc.AutoSize = True
        Me.LblUpdateDisc.Location = New System.Drawing.Point(694, 9)
        Me.LblUpdateDisc.Name = "LblUpdateDisc"
        Me.LblUpdateDisc.Size = New System.Drawing.Size(87, 13)
        Me.LblUpdateDisc.TabIndex = 12
        Me.LblUpdateDisc.TabStop = True
        Me.LblUpdateDisc.Text = "Update Discount"
        '
        'GridDisc
        '
        Me.GridDisc.AllowUserToAddRows = False
        Me.GridDisc.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GridDisc.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.GridDisc.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridDisc.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridDisc.Location = New System.Drawing.Point(0, 32)
        Me.GridDisc.Name = "GridDisc"
        Me.GridDisc.RowHeadersVisible = False
        Me.GridDisc.Size = New System.Drawing.Size(783, 368)
        Me.GridDisc.TabIndex = 10
        '
        'TxtCOC
        '
        Me.TxtCOC.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCOC.Location = New System.Drawing.Point(33, 6)
        Me.TxtCOC.Name = "TxtCOC"
        Me.TxtCOC.Size = New System.Drawing.Size(73, 20)
        Me.TxtCOC.TabIndex = 8
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(0, 9)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(29, 13)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "COC"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(113, 9)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(31, 13)
        Me.Label16.TabIndex = 10
        Me.Label16.Text = "Base"
        '
        'CmbFSource
        '
        Me.CmbFSource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbFSource.FormattingEnabled = True
        Me.CmbFSource.Location = New System.Drawing.Point(419, 28)
        Me.CmbFSource.Name = "CmbFSource"
        Me.CmbFSource.Size = New System.Drawing.Size(78, 21)
        Me.CmbFSource.TabIndex = 26
        '
        'LblFSource
        '
        Me.LblFSource.AutoSize = True
        Me.LblFSource.Location = New System.Drawing.Point(338, 31)
        Me.LblFSource.Name = "LblFSource"
        Me.LblFSource.Size = New System.Drawing.Size(71, 13)
        Me.LblFSource.TabIndex = 10
        Me.LblFSource.Text = "Fare Identifier"
        '
        'LblShowAll
        '
        Me.LblShowAll.AutoSize = True
        Me.LblShowAll.Location = New System.Drawing.Point(255, 482)
        Me.LblShowAll.Name = "LblShowAll"
        Me.LblShowAll.Size = New System.Drawing.Size(45, 13)
        Me.LblShowAll.TabIndex = 28
        Me.LblShowAll.TabStop = True
        Me.LblShowAll.Text = "ShowAll"
        '
        'LblFilter
        '
        Me.LblFilter.AutoSize = True
        Me.LblFilter.Location = New System.Drawing.Point(220, 482)
        Me.LblFilter.Name = "LblFilter"
        Me.LblFilter.Size = New System.Drawing.Size(29, 13)
        Me.LblFilter.TabIndex = 27
        Me.LblFilter.TabStop = True
        Me.LblFilter.Text = "Filter"
        '
        'TxtWhat
        '
        Me.TxtWhat.Location = New System.Drawing.Point(167, 479)
        Me.TxtWhat.Name = "TxtWhat"
        Me.TxtWhat.Size = New System.Drawing.Size(51, 20)
        Me.TxtWhat.TabIndex = 26
        '
        'LblChangeValidity
        '
        Me.LblChangeValidity.AutoSize = True
        Me.LblChangeValidity.Location = New System.Drawing.Point(582, 482)
        Me.LblChangeValidity.Name = "LblChangeValidity"
        Me.LblChangeValidity.Size = New System.Drawing.Size(77, 13)
        Me.LblChangeValidity.TabIndex = 29
        Me.LblChangeValidity.TabStop = True
        Me.LblChangeValidity.Text = "ChangeValidity"
        Me.LblChangeValidity.Visible = False
        '
        'LblViewXpireDate
        '
        Me.LblViewXpireDate.AutoSize = True
        Me.LblViewXpireDate.Location = New System.Drawing.Point(594, 31)
        Me.LblViewXpireDate.Name = "LblViewXpireDate"
        Me.LblViewXpireDate.Size = New System.Drawing.Size(119, 13)
        Me.LblViewXpireDate.TabIndex = 30
        Me.LblViewXpireDate.TabStop = True
        Me.LblViewXpireDate.Text = "MassUpdateExpireDate"
        '
        'GridExpireDate
        '
        Me.GridExpireDate.AllowUserToAddRows = False
        Me.GridExpireDate.AllowUserToDeleteRows = False
        Me.GridExpireDate.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridExpireDate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridExpireDate.Location = New System.Drawing.Point(145, 28)
        Me.GridExpireDate.Name = "GridExpireDate"
        Me.GridExpireDate.RowHeadersVisible = False
        Me.GridExpireDate.Size = New System.Drawing.Size(190, 21)
        Me.GridExpireDate.TabIndex = 21
        Me.GridExpireDate.Visible = False
        '
        'ServiceFee
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.GridExpireDate)
        Me.Controls.Add(Me.LblViewXpireDate)
        Me.Controls.Add(Me.LblChangeValidity)
        Me.Controls.Add(Me.CmbFSource)
        Me.Controls.Add(Me.LblShowAll)
        Me.Controls.Add(Me.LblFilter)
        Me.Controls.Add(Me.TxtWhat)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.ChkQOnly)
        Me.Controls.Add(Me.LblFSource)
        Me.Controls.Add(Me.CmbLevel)
        Me.Controls.Add(Me.ChkActiveOnly)
        Me.Controls.Add(Me.LckLblApproveSF)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtVAL)
        Me.Controls.Add(Me.TxtValidThru)
        Me.Controls.Add(Me.TxtValidFrom)
        Me.Controls.Add(Me.CmbChannel)
        Me.Controls.Add(Me.CmbSFType)
        Me.Controls.Add(Me.CmbAL)
        Me.Controls.Add(Me.CmbSBU)
        Me.Controls.Add(Me.LblPCT)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ServiceFee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Service Fee"
        CType(Me.GridSF, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabSF.ResumeLayout(False)
        Me.TabSF.PerformLayout()
        Me.TabDiscount.ResumeLayout(False)
        Me.TabDiscount.PerformLayout()
        CType(Me.GridDisc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridExpireDate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CmbSBU As System.Windows.Forms.ComboBox
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents CmbTRXtype As System.Windows.Forms.ComboBox
    Friend WithEvents CmbChannel As System.Windows.Forms.ComboBox
    Friend WithEvents CmbBase As System.Windows.Forms.ComboBox
    Friend WithEvents TxtValidFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtValidThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtVAL As System.Windows.Forms.TextBox
    Friend WithEvents TxtFareThru As System.Windows.Forms.TextBox
    Friend WithEvents TxtfareFrm As System.Windows.Forms.TextBox
    Friend WithEvents TxtMin As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtMax As System.Windows.Forms.TextBox
    Friend WithEvents LblUpdateSF As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LckLblApproveSF As System.Windows.Forms.LinkLabel
    Friend WithEvents GridSF As System.Windows.Forms.DataGridView
    Friend WithEvents CmbSFType As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents LblPCT As System.Windows.Forms.Label
    Friend WithEvents ChkNonRefundable As System.Windows.Forms.CheckBox
    Friend WithEvents CmbLevel As System.Windows.Forms.ComboBox
    Friend WithEvents ChkQOnly As System.Windows.Forms.CheckBox
    Friend WithEvents ChkActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabSF As System.Windows.Forms.TabPage
    Friend WithEvents TabDiscount As System.Windows.Forms.TabPage
    Friend WithEvents GridDisc As System.Windows.Forms.DataGridView
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtCOC As System.Windows.Forms.TextBox
    Friend WithEvents LblShowAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LblFilter As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtWhat As System.Windows.Forms.TextBox
    Friend WithEvents LblUpdateDisc As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkINF As System.Windows.Forms.CheckBox
    Friend WithEvents CmbArea As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents LblChangeValidity As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbFSource As System.Windows.Forms.ComboBox
    Friend WithEvents CmbDiscBase As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents LblFSource As System.Windows.Forms.Label
    Friend WithEvents CmbFType As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtDest As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents CmbISI As System.Windows.Forms.ComboBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents LblViewXpireDate As System.Windows.Forms.LinkLabel
    Friend WithEvents GridExpireDate As System.Windows.Forms.DataGridView
End Class
