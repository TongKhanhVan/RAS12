<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvoicePrinting
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
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtDocNo = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TxtSRV = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LblPreview = New System.Windows.Forms.LinkLabel()
        Me.LblDocSearch = New System.Windows.Forms.LinkLabel()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.LblAutoFillTVTR = New System.Windows.Forms.LinkLabel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.txtTaxCode = New System.Windows.Forms.TextBox()
        Me.txtFullName = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.LblGo2InvHandler = New System.Windows.Forms.LinkLabel()
        Me.LblPrintInv = New System.Windows.Forms.LinkLabel()
        Me.LblCalc = New System.Windows.Forms.LinkLabel()
        Me.GridTKT2BInv = New System.Windows.Forms.DataGridView()
        Me.TxtTTLofInv = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.LblSelectAll = New System.Windows.Forms.LinkLabel()
        Me.LblAddToInvoice = New System.Windows.Forms.LinkLabel()
        Me.LblSplitInvoice = New System.Windows.Forms.LinkLabel()
        Me.GridTKT = New System.Windows.Forms.DataGridView()
        Me.S = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtTTLofSelectedTKT = New System.Windows.Forms.TextBox()
        Me.txtIssDate = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.LblClose = New System.Windows.Forms.LinkLabel()
        Me.CmbCust = New System.Windows.Forms.ComboBox()
        Me.P = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.TKNO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FTKT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Curr = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fare = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Tax = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Charge = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ChargeTV = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Discount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InvNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ROE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RTG = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RCPNO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RCPID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RecID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SRV = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Qty = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NetToAL = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.GridTKT2BInv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        CType(Me.GridTKT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(441, 173)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "VND Total"
        '
        'TxtDocNo
        '
        Me.TxtDocNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDocNo.Location = New System.Drawing.Point(35, 18)
        Me.TxtDocNo.Name = "TxtDocNo"
        Me.TxtDocNo.Size = New System.Drawing.Size(89, 20)
        Me.TxtDocNo.TabIndex = 3
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TxtSRV)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.LblPreview)
        Me.GroupBox2.Controls.Add(Me.LblDocSearch)
        Me.GroupBox2.Controls.Add(Me.TxtDocNo)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(151, 68)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Doc. Search"
        '
        'TxtSRV
        '
        Me.TxtSRV.Enabled = False
        Me.TxtSRV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtSRV.Location = New System.Drawing.Point(127, 18)
        Me.TxtSRV.Name = "TxtSRV"
        Me.TxtSRV.Size = New System.Drawing.Size(21, 20)
        Me.TxtSRV.TabIndex = 17
        Me.TxtSRV.Text = "S"
        Me.TxtSRV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 13)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "TRX"
        '
        'LblPreview
        '
        Me.LblPreview.AutoSize = True
        Me.LblPreview.Location = New System.Drawing.Point(84, 45)
        Me.LblPreview.Name = "LblPreview"
        Me.LblPreview.Size = New System.Drawing.Size(45, 13)
        Me.LblPreview.TabIndex = 11
        Me.LblPreview.TabStop = True
        Me.LblPreview.Text = "Preview"
        Me.LblPreview.Visible = False
        '
        'LblDocSearch
        '
        Me.LblDocSearch.AutoSize = True
        Me.LblDocSearch.Location = New System.Drawing.Point(30, 45)
        Me.LblDocSearch.Name = "LblDocSearch"
        Me.LblDocSearch.Size = New System.Drawing.Size(41, 13)
        Me.LblDocSearch.TabIndex = 14
        Me.LblDocSearch.TabStop = True
        Me.LblDocSearch.Text = "Search"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.LblAutoFillTVTR)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.txtAddress)
        Me.GroupBox3.Controls.Add(Me.txtTaxCode)
        Me.GroupBox3.Controls.Add(Me.txtFullName)
        Me.GroupBox3.Location = New System.Drawing.Point(159, 5)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(482, 69)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Customer Details"
        '
        'LblAutoFillTVTR
        '
        Me.LblAutoFillTVTR.AutoSize = True
        Me.LblAutoFillTVTR.Location = New System.Drawing.Point(417, 21)
        Me.LblAutoFillTVTR.Name = "LblAutoFillTVTR"
        Me.LblAutoFillTVTR.Size = New System.Drawing.Size(19, 13)
        Me.LblAutoFillTVTR.TabIndex = 13
        Me.LblAutoFillTVTR.TabStop = True
        Me.LblAutoFillTVTR.Text = "...."
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 45)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Address"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "FullName"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(367, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(50, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "TaxCode"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(61, 42)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(306, 20)
        Me.txtAddress.TabIndex = 10
        '
        'txtTaxCode
        '
        Me.txtTaxCode.Location = New System.Drawing.Point(370, 42)
        Me.txtTaxCode.Name = "txtTaxCode"
        Me.txtTaxCode.Size = New System.Drawing.Size(92, 20)
        Me.txtTaxCode.TabIndex = 10
        '
        'txtFullName
        '
        Me.txtFullName.Location = New System.Drawing.Point(61, 19)
        Me.txtFullName.Name = "txtFullName"
        Me.txtFullName.Size = New System.Drawing.Size(306, 20)
        Me.txtFullName.TabIndex = 10
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.LblGo2InvHandler)
        Me.GroupBox4.Controls.Add(Me.LblPrintInv)
        Me.GroupBox4.Controls.Add(Me.LblCalc)
        Me.GroupBox4.Controls.Add(Me.GridTKT2BInv)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.TxtTTLofInv)
        Me.GroupBox4.Location = New System.Drawing.Point(5, 307)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(636, 191)
        Me.GroupBox4.TabIndex = 5
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Tickets to be Invoiced"
        '
        'LblGo2InvHandler
        '
        Me.LblGo2InvHandler.AutoSize = True
        Me.LblGo2InvHandler.Enabled = False
        Me.LblGo2InvHandler.Location = New System.Drawing.Point(327, 173)
        Me.LblGo2InvHandler.Name = "LblGo2InvHandler"
        Me.LblGo2InvHandler.Size = New System.Drawing.Size(79, 13)
        Me.LblGo2InvHandler.TabIndex = 12
        Me.LblGo2InvHandler.TabStop = True
        Me.LblGo2InvHandler.Text = "Go2InvHandler"
        '
        'LblPrintInv
        '
        Me.LblPrintInv.AutoSize = True
        Me.LblPrintInv.Enabled = False
        Me.LblPrintInv.Location = New System.Drawing.Point(74, 173)
        Me.LblPrintInv.Name = "LblPrintInv"
        Me.LblPrintInv.Size = New System.Drawing.Size(32, 13)
        Me.LblPrintInv.TabIndex = 10
        Me.LblPrintInv.TabStop = True
        Me.LblPrintInv.Text = "Save"
        '
        'LblCalc
        '
        Me.LblCalc.AutoSize = True
        Me.LblCalc.Location = New System.Drawing.Point(6, 173)
        Me.LblCalc.Name = "LblCalc"
        Me.LblCalc.Size = New System.Drawing.Size(51, 13)
        Me.LblCalc.TabIndex = 10
        Me.LblCalc.TabStop = True
        Me.LblCalc.Text = "Calculate"
        '
        'GridTKT2BInv
        '
        Me.GridTKT2BInv.AllowUserToAddRows = False
        Me.GridTKT2BInv.AllowUserToDeleteRows = False
        Me.GridTKT2BInv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridTKT2BInv.BackgroundColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridTKT2BInv.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.GridTKT2BInv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTKT2BInv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.P, Me.TKNO, Me.FTKT, Me.Curr, Me.Fare, Me.Tax, Me.Charge, Me.ChargeTV, Me.Discount, Me.InvNo, Me.ROE, Me.RTG, Me.RCPNO, Me.RCPID, Me.RecID, Me.SRV, Me.Qty, Me.NetToAL})
        Me.GridTKT2BInv.Location = New System.Drawing.Point(6, 13)
        Me.GridTKT2BInv.Name = "GridTKT2BInv"
        Me.GridTKT2BInv.RowHeadersVisible = False
        Me.GridTKT2BInv.Size = New System.Drawing.Size(624, 151)
        Me.GridTKT2BInv.TabIndex = 0
        '
        'TxtTTLofInv
        '
        Me.TxtTTLofInv.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTTLofInv.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTTLofInv.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtTTLofInv.Location = New System.Drawing.Point(504, 170)
        Me.TxtTTLofInv.Name = "TxtTTLofInv"
        Me.TxtTTLofInv.Size = New System.Drawing.Size(126, 20)
        Me.TxtTTLofInv.TabIndex = 3
        Me.TxtTTLofInv.Text = "0"
        Me.TxtTTLofInv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.LblSelectAll)
        Me.GroupBox5.Controls.Add(Me.LblAddToInvoice)
        Me.GroupBox5.Controls.Add(Me.LblSplitInvoice)
        Me.GroupBox5.Controls.Add(Me.GridTKT)
        Me.GroupBox5.Controls.Add(Me.Label2)
        Me.GroupBox5.Controls.Add(Me.txtTTLofSelectedTKT)
        Me.GroupBox5.Location = New System.Drawing.Point(5, 79)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(636, 222)
        Me.GroupBox5.TabIndex = 5
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Ticket List"
        '
        'LblSelectAll
        '
        Me.LblSelectAll.AutoSize = True
        Me.LblSelectAll.Location = New System.Drawing.Point(94, 201)
        Me.LblSelectAll.Name = "LblSelectAll"
        Me.LblSelectAll.Size = New System.Drawing.Size(51, 13)
        Me.LblSelectAll.TabIndex = 10
        Me.LblSelectAll.TabStop = True
        Me.LblSelectAll.Text = "Select All"
        '
        'LblAddToInvoice
        '
        Me.LblAddToInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblAddToInvoice.AutoSize = True
        Me.LblAddToInvoice.Location = New System.Drawing.Point(168, 201)
        Me.LblAddToInvoice.Name = "LblAddToInvoice"
        Me.LblAddToInvoice.Size = New System.Drawing.Size(80, 13)
        Me.LblAddToInvoice.TabIndex = 9
        Me.LblAddToInvoice.TabStop = True
        Me.LblAddToInvoice.Text = "Add To Invoice"
        '
        'LblSplitInvoice
        '
        Me.LblSplitInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblSplitInvoice.AutoSize = True
        Me.LblSplitInvoice.Location = New System.Drawing.Point(6, 201)
        Me.LblSplitInvoice.Name = "LblSplitInvoice"
        Me.LblSplitInvoice.Size = New System.Drawing.Size(65, 13)
        Me.LblSplitInvoice.TabIndex = 8
        Me.LblSplitInvoice.TabStop = True
        Me.LblSplitInvoice.Text = "Split Invoice"
        Me.LblSplitInvoice.Visible = False
        '
        'GridTKT
        '
        Me.GridTKT.AllowUserToAddRows = False
        Me.GridTKT.AllowUserToDeleteRows = False
        Me.GridTKT.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridTKT.BackgroundColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridTKT.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.GridTKT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTKT.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.S})
        Me.GridTKT.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.GridTKT.Location = New System.Drawing.Point(6, 13)
        Me.GridTKT.Name = "GridTKT"
        Me.GridTKT.RowHeadersVisible = False
        Me.GridTKT.Size = New System.Drawing.Size(624, 179)
        Me.GridTKT.TabIndex = 0
        '
        'S
        '
        Me.S.HeaderText = "S"
        Me.S.Name = "S"
        Me.S.Width = 25
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(441, 204)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "VND Total"
        '
        'txtTTLofSelectedTKT
        '
        Me.txtTTLofSelectedTKT.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTTLofSelectedTKT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTTLofSelectedTKT.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTTLofSelectedTKT.Location = New System.Drawing.Point(504, 201)
        Me.txtTTLofSelectedTKT.Name = "txtTTLofSelectedTKT"
        Me.txtTTLofSelectedTKT.Size = New System.Drawing.Size(126, 20)
        Me.txtTTLofSelectedTKT.TabIndex = 3
        Me.txtTTLofSelectedTKT.Text = "0"
        Me.txtTTLofSelectedTKT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtIssDate
        '
        Me.txtIssDate.CustomFormat = "dd-MMM-yy"
        Me.txtIssDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtIssDate.Location = New System.Drawing.Point(705, 23)
        Me.txtIssDate.Name = "txtIssDate"
        Me.txtIssDate.Size = New System.Drawing.Size(83, 20)
        Me.txtIssDate.TabIndex = 8
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(644, 27)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(55, 13)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "IssueDate"
        '
        'LblClose
        '
        Me.LblClose.AutoSize = True
        Me.LblClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblClose.Location = New System.Drawing.Point(750, 0)
        Me.LblClose.Name = "LblClose"
        Me.LblClose.Size = New System.Drawing.Size(38, 13)
        Me.LblClose.TabIndex = 14
        Me.LblClose.TabStop = True
        Me.LblClose.Text = "Close"
        '
        'CmbCust
        '
        Me.CmbCust.Enabled = False
        Me.CmbCust.FormattingEnabled = True
        Me.CmbCust.Location = New System.Drawing.Point(647, 47)
        Me.CmbCust.Name = "CmbCust"
        Me.CmbCust.Size = New System.Drawing.Size(141, 21)
        Me.CmbCust.TabIndex = 15
        '
        'P
        '
        Me.P.HeaderText = "P"
        Me.P.Name = "P"
        Me.P.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.P.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.P.Visible = False
        Me.P.Width = 25
        '
        'TKNO
        '
        Me.TKNO.HeaderText = "TKNO"
        Me.TKNO.Name = "TKNO"
        Me.TKNO.ReadOnly = True
        Me.TKNO.Width = 95
        '
        'FTKT
        '
        Me.FTKT.HeaderText = "FTKT"
        Me.FTKT.Name = "FTKT"
        Me.FTKT.ReadOnly = True
        Me.FTKT.Width = 50
        '
        'Curr
        '
        Me.Curr.HeaderText = "Curr"
        Me.Curr.Name = "Curr"
        Me.Curr.ReadOnly = True
        Me.Curr.Width = 35
        '
        'Fare
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle2.Format = "#,##0.00"
        Me.Fare.DefaultCellStyle = DataGridViewCellStyle2
        Me.Fare.HeaderText = "Fare"
        Me.Fare.Name = "Fare"
        Me.Fare.Width = 75
        '
        'Tax
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.Format = "#,##0.00"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.Tax.DefaultCellStyle = DataGridViewCellStyle3
        Me.Tax.HeaderText = "Tax"
        Me.Tax.Name = "Tax"
        Me.Tax.Width = 75
        '
        'Charge
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.Format = "#,##0.00"
        Me.Charge.DefaultCellStyle = DataGridViewCellStyle4
        Me.Charge.HeaderText = "Charge"
        Me.Charge.Name = "Charge"
        Me.Charge.Width = 70
        '
        'ChargeTV
        '
        Me.ChargeTV.HeaderText = "TVCharge"
        Me.ChargeTV.Name = "ChargeTV"
        '
        'Discount
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.Format = "#,##0.00"
        Me.Discount.DefaultCellStyle = DataGridViewCellStyle5
        Me.Discount.HeaderText = "Discount"
        Me.Discount.Name = "Discount"
        Me.Discount.Width = 70
        '
        'InvNo
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.InvNo.DefaultCellStyle = DataGridViewCellStyle6
        Me.InvNo.HeaderText = "InvNo"
        Me.InvNo.Name = "InvNo"
        Me.InvNo.ReadOnly = True
        Me.InvNo.Width = 45
        '
        'ROE
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.Format = "#,##0.00"
        Me.ROE.DefaultCellStyle = DataGridViewCellStyle7
        Me.ROE.HeaderText = "ROE"
        Me.ROE.Name = "ROE"
        Me.ROE.ReadOnly = True
        Me.ROE.Width = 70
        '
        'RTG
        '
        Me.RTG.HeaderText = "RTG"
        Me.RTG.Name = "RTG"
        Me.RTG.Visible = False
        '
        'RCPNO
        '
        Me.RCPNO.HeaderText = "RCPNO"
        Me.RCPNO.Name = "RCPNO"
        Me.RCPNO.Visible = False
        '
        'RCPID
        '
        Me.RCPID.HeaderText = "RCPID"
        Me.RCPID.Name = "RCPID"
        Me.RCPID.Visible = False
        '
        'RecID
        '
        Me.RecID.HeaderText = "RecID"
        Me.RecID.Name = "RecID"
        Me.RecID.Visible = False
        '
        'SRV
        '
        Me.SRV.HeaderText = "SRV"
        Me.SRV.Name = "SRV"
        Me.SRV.Visible = False
        '
        'Qty
        '
        Me.Qty.HeaderText = "Qty"
        Me.Qty.Name = "Qty"
        Me.Qty.Visible = False
        '
        'NetToAL
        '
        Me.NetToAL.HeaderText = "NetToAL"
        Me.NetToAL.Name = "NetToAL"
        Me.NetToAL.Visible = False
        '
        'InvoicePrinting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.MistyRose
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.ControlBox = False
        Me.Controls.Add(Me.CmbCust)
        Me.Controls.Add(Me.LblClose)
        Me.Controls.Add(Me.txtIssDate)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.Name = "InvoicePrinting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Invoice Printer"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.GridTKT2BInv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        CType(Me.GridTKT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtDocNo As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GridTKT2BInv As System.Windows.Forms.DataGridView
    Friend WithEvents TxtTTLofInv As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GridTKT As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtTTLofSelectedTKT As System.Windows.Forms.TextBox
    Friend WithEvents txtTaxCode As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtFullName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents S As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents LblSplitInvoice As System.Windows.Forms.LinkLabel
    Friend WithEvents LblAddToInvoice As System.Windows.Forms.LinkLabel
    Friend WithEvents LblPrintInv As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCalc As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDocSearch As System.Windows.Forms.LinkLabel
    Friend WithEvents LblSelectAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LblGo2InvHandler As System.Windows.Forms.LinkLabel
    Friend WithEvents txtIssDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents LblClose As System.Windows.Forms.LinkLabel
    Friend WithEvents LblPreview As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbCust As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtSRV As System.Windows.Forms.TextBox
    Friend WithEvents LblAutoFillTVTR As System.Windows.Forms.LinkLabel
    Friend WithEvents P As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents TKNO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FTKT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Curr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Fare As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Tax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Charge As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChargeTV As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Discount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InvNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ROE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RTG As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RCPNO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RCPID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RecID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SRV As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NetToAL As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
