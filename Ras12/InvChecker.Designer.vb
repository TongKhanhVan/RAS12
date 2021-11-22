<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvChecker
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
        Me.CmbAL = New System.Windows.Forms.ComboBox()
        Me.CmbMonth = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GridHD = New System.Windows.Forms.DataGridView()
        Me.CmbINVType = New System.Windows.Forms.ComboBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.LblIssueINV = New System.Windows.Forms.LinkLabel()
        Me.LblLoadGridTRXwoINV = New System.Windows.Forms.LinkLabel()
        Me.GridTRXwoINV = New System.Windows.Forms.DataGridView()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.CmbRoundingAmt = New System.Windows.Forms.ComboBox()
        Me.LblAdj = New System.Windows.Forms.LinkLabel()
        Me.LblView = New System.Windows.Forms.LinkLabel()
        Me.LblRounding = New System.Windows.Forms.LinkLabel()
        Me.LblUpdateInfor = New System.Windows.Forms.LinkLabel()
        Me.LblUpdatePickUp = New System.Windows.Forms.LinkLabel()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.txtQty = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtTTLINVByS1 = New System.Windows.Forms.TextBox()
        Me.TxtDiff_VND = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TxtTTLINV = New System.Windows.Forms.TextBox()
        Me.txtADM_VND = New System.Windows.Forms.TextBox()
        Me.TxtComm_VND = New System.Windows.Forms.TextBox()
        Me.TxtComm = New System.Windows.Forms.TextBox()
        Me.TxtInvTaxCharge = New System.Windows.Forms.TextBox()
        Me.TxtInvFare = New System.Windows.Forms.TextBox()
        Me.Txt2AL_VND = New System.Windows.Forms.TextBox()
        Me.CmbCurr = New System.Windows.Forms.ComboBox()
        Me.LblUpdateTKT = New System.Windows.Forms.LinkLabel()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TxtTax = New System.Windows.Forms.TextBox()
        Me.TxtVND = New System.Windows.Forms.TextBox()
        Me.txtFare = New System.Windows.Forms.TextBox()
        Me.TxtRTG = New System.Windows.Forms.TextBox()
        Me.txtTKNO = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GridTKT = New System.Windows.Forms.DataGridView()
        Me.GridFOP = New System.Windows.Forms.DataGridView()
        Me.GridRCP = New System.Windows.Forms.DataGridView()
        Me.CmbBSPAL = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LblCalc = New System.Windows.Forms.LinkLabel()
        Me.LblSaveAmtToAL = New System.Windows.Forms.LinkLabel()
        CType(Me.GridHD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        CType(Me.GridTRXwoINV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.GridTKT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridFOP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridRCP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CmbAL
        '
        Me.CmbAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(23, 6)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(39, 21)
        Me.CmbAL.TabIndex = 0
        '
        'CmbMonth
        '
        Me.CmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbMonth.FormattingEnabled = True
        Me.CmbMonth.Location = New System.Drawing.Point(100, 6)
        Me.CmbMonth.Name = "CmbMonth"
        Me.CmbMonth.Size = New System.Drawing.Size(60, 21)
        Me.CmbMonth.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(20, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "AL"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(62, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Month"
        '
        'GridHD
        '
        Me.GridHD.AllowUserToAddRows = False
        Me.GridHD.AllowUserToDeleteRows = False
        Me.GridHD.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridHD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridHD.GridColor = System.Drawing.Color.AliceBlue
        Me.GridHD.Location = New System.Drawing.Point(3, 3)
        Me.GridHD.Name = "GridHD"
        Me.GridHD.RowHeadersVisible = False
        Me.GridHD.Size = New System.Drawing.Size(760, 407)
        Me.GridHD.TabIndex = 3
        '
        'CmbINVType
        '
        Me.CmbINVType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbINVType.FormattingEnabled = True
        Me.CmbINVType.Items.AddRange(New Object() {"ALL", "ORG", "RSV"})
        Me.CmbINVType.Location = New System.Drawing.Point(163, 6)
        Me.CmbINVType.Name = "CmbINVType"
        Me.CmbINVType.Size = New System.Drawing.Size(49, 21)
        Me.CmbINVType.TabIndex = 4
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(3, 33)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(777, 466)
        Me.TabControl1.TabIndex = 5
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.LblIssueINV)
        Me.TabPage4.Controls.Add(Me.LblLoadGridTRXwoINV)
        Me.TabPage4.Controls.Add(Me.GridTRXwoINV)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(769, 440)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "S0-TRX wo INV"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'LblIssueINV
        '
        Me.LblIssueINV.AutoSize = True
        Me.LblIssueINV.Location = New System.Drawing.Point(713, 419)
        Me.LblIssueINV.Name = "LblIssueINV"
        Me.LblIssueINV.Size = New System.Drawing.Size(53, 13)
        Me.LblIssueINV.TabIndex = 2
        Me.LblIssueINV.TabStop = True
        Me.LblIssueINV.Text = "Issue INV"
        Me.LblIssueINV.Visible = False
        '
        'LblLoadGridTRXwoINV
        '
        Me.LblLoadGridTRXwoINV.AutoSize = True
        Me.LblLoadGridTRXwoINV.Location = New System.Drawing.Point(5, 419)
        Me.LblLoadGridTRXwoINV.Name = "LblLoadGridTRXwoINV"
        Me.LblLoadGridTRXwoINV.Size = New System.Drawing.Size(31, 13)
        Me.LblLoadGridTRXwoINV.TabIndex = 1
        Me.LblLoadGridTRXwoINV.TabStop = True
        Me.LblLoadGridTRXwoINV.Text = "Load"
        '
        'GridTRXwoINV
        '
        Me.GridTRXwoINV.AllowUserToAddRows = False
        Me.GridTRXwoINV.AllowUserToDeleteRows = False
        Me.GridTRXwoINV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTRXwoINV.Location = New System.Drawing.Point(2, 0)
        Me.GridTRXwoINV.Name = "GridTRXwoINV"
        Me.GridTRXwoINV.RowHeadersVisible = False
        Me.GridTRXwoINV.Size = New System.Drawing.Size(764, 416)
        Me.GridTRXwoINV.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.CmbRoundingAmt)
        Me.TabPage1.Controls.Add(Me.LblAdj)
        Me.TabPage1.Controls.Add(Me.LblView)
        Me.TabPage1.Controls.Add(Me.LblRounding)
        Me.TabPage1.Controls.Add(Me.LblUpdateInfor)
        Me.TabPage1.Controls.Add(Me.LblUpdatePickUp)
        Me.TabPage1.Controls.Add(Me.GridHD)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(769, 440)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "S1-Inv Infor"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'CmbRoundingAmt
        '
        Me.CmbRoundingAmt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbRoundingAmt.FormattingEnabled = True
        Me.CmbRoundingAmt.Items.AddRange(New Object() {"+8", "+7", "+6", "+5", "+4", "+3", "+2", "+1", "-1", "-2", "-3", "-4", "-5", "-6", "-7", "-8"})
        Me.CmbRoundingAmt.Location = New System.Drawing.Point(282, 416)
        Me.CmbRoundingAmt.Name = "CmbRoundingAmt"
        Me.CmbRoundingAmt.Size = New System.Drawing.Size(37, 21)
        Me.CmbRoundingAmt.TabIndex = 10
        '
        'LblAdj
        '
        Me.LblAdj.AutoSize = True
        Me.LblAdj.Location = New System.Drawing.Point(365, 419)
        Me.LblAdj.Name = "LblAdj"
        Me.LblAdj.Size = New System.Drawing.Size(57, 13)
        Me.LblAdj.TabIndex = 9
        Me.LblAdj.TabStop = True
        Me.LblAdj.Text = "Adjust Amt"
        Me.LblAdj.Visible = False
        '
        'LblView
        '
        Me.LblView.AutoSize = True
        Me.LblView.Location = New System.Drawing.Point(732, 419)
        Me.LblView.Name = "LblView"
        Me.LblView.Size = New System.Drawing.Size(30, 13)
        Me.LblView.TabIndex = 8
        Me.LblView.TabStop = True
        Me.LblView.Text = "View"
        '
        'LblRounding
        '
        Me.LblRounding.AutoSize = True
        Me.LblRounding.Location = New System.Drawing.Point(223, 419)
        Me.LblRounding.Name = "LblRounding"
        Me.LblRounding.Size = New System.Drawing.Size(53, 13)
        Me.LblRounding.TabIndex = 5
        Me.LblRounding.TabStop = True
        Me.LblRounding.Text = "Rounding"
        Me.LblRounding.Visible = False
        '
        'LblUpdateInfor
        '
        Me.LblUpdateInfor.AutoSize = True
        Me.LblUpdateInfor.Location = New System.Drawing.Point(122, 419)
        Me.LblUpdateInfor.Name = "LblUpdateInfor"
        Me.LblUpdateInfor.Size = New System.Drawing.Size(66, 13)
        Me.LblUpdateInfor.TabIndex = 5
        Me.LblUpdateInfor.TabStop = True
        Me.LblUpdateInfor.Text = "Update Infor"
        Me.LblUpdateInfor.Visible = False
        '
        'LblUpdatePickUp
        '
        Me.LblUpdatePickUp.AutoSize = True
        Me.LblUpdatePickUp.Location = New System.Drawing.Point(3, 419)
        Me.LblUpdatePickUp.Name = "LblUpdatePickUp"
        Me.LblUpdatePickUp.Size = New System.Drawing.Size(113, 13)
        Me.LblUpdatePickUp.TabIndex = 5
        Me.LblUpdatePickUp.TabStop = True
        Me.LblUpdatePickUp.Text = "Update PickUp Status"
        Me.LblUpdatePickUp.Visible = False
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.txtQty)
        Me.TabPage3.Controls.Add(Me.GroupBox1)
        Me.TabPage3.Controls.Add(Me.CmbCurr)
        Me.TabPage3.Controls.Add(Me.LblUpdateTKT)
        Me.TabPage3.Controls.Add(Me.Label12)
        Me.TabPage3.Controls.Add(Me.Label11)
        Me.TabPage3.Controls.Add(Me.Label9)
        Me.TabPage3.Controls.Add(Me.Label7)
        Me.TabPage3.Controls.Add(Me.Label6)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.TxtTax)
        Me.TabPage3.Controls.Add(Me.TxtVND)
        Me.TabPage3.Controls.Add(Me.txtFare)
        Me.TabPage3.Controls.Add(Me.TxtRTG)
        Me.TabPage3.Controls.Add(Me.txtTKNO)
        Me.TabPage3.Controls.Add(Me.Label10)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(769, 440)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "S2-Adj Amt"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'txtQty
        '
        Me.txtQty.Location = New System.Drawing.Point(184, 5)
        Me.txtQty.Name = "txtQty"
        Me.txtQty.Size = New System.Drawing.Size(55, 20)
        Me.txtQty.TabIndex = 9
        Me.txtQty.Text = "1"
        Me.txtQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LblSaveAmtToAL)
        Me.GroupBox1.Controls.Add(Me.TxtTTLINVByS1)
        Me.GroupBox1.Controls.Add(Me.TxtDiff_VND)
        Me.GroupBox1.Controls.Add(Me.Label22)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.Label20)
        Me.GroupBox1.Controls.Add(Me.Label21)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TxtTTLINV)
        Me.GroupBox1.Controls.Add(Me.txtADM_VND)
        Me.GroupBox1.Controls.Add(Me.TxtComm_VND)
        Me.GroupBox1.Controls.Add(Me.TxtComm)
        Me.GroupBox1.Controls.Add(Me.TxtInvTaxCharge)
        Me.GroupBox1.Controls.Add(Me.TxtInvFare)
        Me.GroupBox1.Controls.Add(Me.Txt2AL_VND)
        Me.GroupBox1.Location = New System.Drawing.Point(514, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(249, 283)
        Me.GroupBox1.TabIndex = 23
        Me.GroupBox1.TabStop = False
        '
        'TxtTTLINVByS1
        '
        Me.TxtTTLINVByS1.ForeColor = System.Drawing.SystemColors.ActiveBorder
        Me.TxtTTLINVByS1.Location = New System.Drawing.Point(139, 11)
        Me.TxtTTLINVByS1.Name = "TxtTTLINVByS1"
        Me.TxtTTLINVByS1.Size = New System.Drawing.Size(100, 20)
        Me.TxtTTLINVByS1.TabIndex = 15
        Me.TxtTTLINVByS1.Text = "0"
        Me.TxtTTLINVByS1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtDiff_VND
        '
        Me.TxtDiff_VND.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDiff_VND.Location = New System.Drawing.Point(140, 255)
        Me.TxtDiff_VND.Name = "TxtDiff_VND"
        Me.TxtDiff_VND.ReadOnly = True
        Me.TxtDiff_VND.Size = New System.Drawing.Size(100, 20)
        Me.TxtDiff_VND.TabIndex = 7
        Me.TxtDiff_VND.Text = "0"
        Me.TxtDiff_VND.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(6, 258)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(83, 13)
        Me.Label22.TabIndex = 14
        Me.Label22.Text = "e. Diff (a-b2-c-d)"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(7, 231)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(48, 13)
        Me.Label17.TabIndex = 14
        Me.Label17.Text = "d. To AL"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(7, 204)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(70, 13)
        Me.Label20.TabIndex = 14
        Me.Label20.Text = "c. Thu chi ho"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(6, 181)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(81, 13)
        Me.Label21.TabIndex = 14
        Me.Label21.Text = "  b2. Comm Amt"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(7, 151)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(68, 13)
        Me.Label14.TabIndex = 14
        Me.Label14.Text = "  b1. %Comm"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(6, 73)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(70, 13)
        Me.Label19.TabIndex = 14
        Me.Label19.Text = "  a1. Inv Fare"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(113, 14)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(20, 13)
        Me.Label13.TabIndex = 14
        Me.Label13.Text = "S1"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(84, 13)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "a, TTL INV Amt "
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(6, 122)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(48, 13)
        Me.Label18.TabIndex = 13
        Me.Label18.Text = "b. Comm"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "  a2. Inv Tax/Charge"
        '
        'TxtTTLINV
        '
        Me.TxtTTLINV.Enabled = False
        Me.TxtTTLINV.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtTTLINV.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.TxtTTLINV.Location = New System.Drawing.Point(139, 45)
        Me.TxtTTLINV.Name = "TxtTTLINV"
        Me.TxtTTLINV.Size = New System.Drawing.Size(100, 20)
        Me.TxtTTLINV.TabIndex = 9
        Me.TxtTTLINV.Text = "0"
        Me.TxtTTLINV.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtADM_VND
        '
        Me.txtADM_VND.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtADM_VND.Location = New System.Drawing.Point(140, 201)
        Me.txtADM_VND.Name = "txtADM_VND"
        Me.txtADM_VND.Size = New System.Drawing.Size(100, 20)
        Me.txtADM_VND.TabIndex = 12
        Me.txtADM_VND.Text = "0"
        Me.txtADM_VND.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtComm_VND
        '
        Me.TxtComm_VND.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtComm_VND.Location = New System.Drawing.Point(140, 178)
        Me.TxtComm_VND.Name = "TxtComm_VND"
        Me.TxtComm_VND.ReadOnly = True
        Me.TxtComm_VND.Size = New System.Drawing.Size(100, 20)
        Me.TxtComm_VND.TabIndex = 12
        Me.TxtComm_VND.Text = "0"
        Me.TxtComm_VND.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtComm
        '
        Me.TxtComm.Location = New System.Drawing.Point(153, 148)
        Me.TxtComm.Name = "TxtComm"
        Me.TxtComm.Size = New System.Drawing.Size(86, 20)
        Me.TxtComm.TabIndex = 12
        Me.TxtComm.Text = "3"
        Me.TxtComm.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtInvTaxCharge
        '
        Me.TxtInvTaxCharge.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtInvTaxCharge.Location = New System.Drawing.Point(139, 93)
        Me.TxtInvTaxCharge.Name = "TxtInvTaxCharge"
        Me.TxtInvTaxCharge.Size = New System.Drawing.Size(100, 20)
        Me.TxtInvTaxCharge.TabIndex = 9
        Me.TxtInvTaxCharge.Text = "0"
        Me.TxtInvTaxCharge.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtInvFare
        '
        Me.TxtInvFare.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtInvFare.Location = New System.Drawing.Point(139, 70)
        Me.TxtInvFare.Name = "TxtInvFare"
        Me.TxtInvFare.Size = New System.Drawing.Size(100, 20)
        Me.TxtInvFare.TabIndex = 9
        Me.TxtInvFare.Text = "0"
        Me.TxtInvFare.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Txt2AL_VND
        '
        Me.Txt2AL_VND.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt2AL_VND.Location = New System.Drawing.Point(139, 228)
        Me.Txt2AL_VND.Name = "Txt2AL_VND"
        Me.Txt2AL_VND.Size = New System.Drawing.Size(100, 20)
        Me.Txt2AL_VND.TabIndex = 9
        Me.Txt2AL_VND.Text = "0"
        Me.Txt2AL_VND.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'CmbCurr
        '
        Me.CmbCurr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCurr.Enabled = False
        Me.CmbCurr.FormattingEnabled = True
        Me.CmbCurr.Items.AddRange(New Object() {"USD", "VND"})
        Me.CmbCurr.Location = New System.Drawing.Point(278, 5)
        Me.CmbCurr.Name = "CmbCurr"
        Me.CmbCurr.Size = New System.Drawing.Size(87, 21)
        Me.CmbCurr.TabIndex = 17
        '
        'LblUpdateTKT
        '
        Me.LblUpdateTKT.AutoSize = True
        Me.LblUpdateTKT.Location = New System.Drawing.Point(323, 81)
        Me.LblUpdateTKT.Name = "LblUpdateTKT"
        Me.LblUpdateTKT.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdateTKT.TabIndex = 22
        Me.LblUpdateTKT.TabStop = True
        Me.LblUpdateTKT.Text = "Update"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(243, 8)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(29, 13)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "Curr."
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(152, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(23, 13)
        Me.Label11.TabIndex = 17
        Me.Label11.Text = "Qty"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(152, 58)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(25, 13)
        Me.Label9.TabIndex = 17
        Me.Label9.Text = "Tax"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 58)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(28, 13)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Fare"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 13)
        Me.Label6.TabIndex = 20
        Me.Label6.Text = "Itinerary"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(4, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(37, 13)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "TKNO"
        '
        'TxtTax
        '
        Me.TxtTax.Location = New System.Drawing.Point(184, 55)
        Me.TxtTax.Name = "TxtTax"
        Me.TxtTax.Size = New System.Drawing.Size(55, 20)
        Me.TxtTax.TabIndex = 20
        Me.TxtTax.Text = "0"
        Me.TxtTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtVND
        '
        Me.TxtVND.Location = New System.Drawing.Point(278, 55)
        Me.TxtVND.Name = "TxtVND"
        Me.TxtVND.ReadOnly = True
        Me.TxtVND.Size = New System.Drawing.Size(87, 20)
        Me.TxtVND.TabIndex = 22
        Me.TxtVND.Text = "0"
        Me.TxtVND.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtFare
        '
        Me.txtFare.Location = New System.Drawing.Point(48, 55)
        Me.txtFare.Name = "txtFare"
        Me.txtFare.Size = New System.Drawing.Size(98, 20)
        Me.txtFare.TabIndex = 19
        Me.txtFare.Text = "0"
        Me.txtFare.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtRTG
        '
        Me.TxtRTG.Location = New System.Drawing.Point(48, 29)
        Me.TxtRTG.Name = "TxtRTG"
        Me.TxtRTG.Size = New System.Drawing.Size(317, 20)
        Me.TxtRTG.TabIndex = 18
        '
        'txtTKNO
        '
        Me.txtTKNO.Location = New System.Drawing.Point(48, 5)
        Me.txtTKNO.Name = "txtTKNO"
        Me.txtTKNO.ReadOnly = True
        Me.txtTKNO.Size = New System.Drawing.Size(98, 20)
        Me.txtTKNO.TabIndex = 16
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(242, 58)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(37, 13)
        Me.Label10.TabIndex = 19
        Me.Label10.Text = "~VND"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GridTKT)
        Me.TabPage2.Controls.Add(Me.GridFOP)
        Me.TabPage2.Controls.Add(Me.GridRCP)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(769, 440)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "View"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GridTKT
        '
        Me.GridTKT.AllowUserToAddRows = False
        Me.GridTKT.AllowUserToDeleteRows = False
        Me.GridTKT.BackgroundColor = System.Drawing.Color.Azure
        Me.GridTKT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTKT.Location = New System.Drawing.Point(3, 82)
        Me.GridTKT.Name = "GridTKT"
        Me.GridTKT.RowHeadersVisible = False
        Me.GridTKT.Size = New System.Drawing.Size(760, 221)
        Me.GridTKT.TabIndex = 0
        '
        'GridFOP
        '
        Me.GridFOP.AllowUserToAddRows = False
        Me.GridFOP.AllowUserToDeleteRows = False
        Me.GridFOP.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridFOP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridFOP.Location = New System.Drawing.Point(3, 305)
        Me.GridFOP.Name = "GridFOP"
        Me.GridFOP.RowHeadersVisible = False
        Me.GridFOP.Size = New System.Drawing.Size(760, 106)
        Me.GridFOP.TabIndex = 0
        '
        'GridRCP
        '
        Me.GridRCP.AllowUserToAddRows = False
        Me.GridRCP.AllowUserToDeleteRows = False
        Me.GridRCP.BackgroundColor = System.Drawing.Color.AliceBlue
        Me.GridRCP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridRCP.Location = New System.Drawing.Point(3, 3)
        Me.GridRCP.Name = "GridRCP"
        Me.GridRCP.RowHeadersVisible = False
        Me.GridRCP.Size = New System.Drawing.Size(760, 78)
        Me.GridRCP.TabIndex = 0
        '
        'CmbBSPAL
        '
        Me.CmbBSPAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbBSPAL.FormattingEnabled = True
        Me.CmbBSPAL.Location = New System.Drawing.Point(261, 6)
        Me.CmbBSPAL.Name = "CmbBSPAL"
        Me.CmbBSPAL.Size = New System.Drawing.Size(43, 21)
        Me.CmbBSPAL.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(215, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "BSP AL"
        '
        'LblCalc
        '
        Me.LblCalc.AutoSize = True
        Me.LblCalc.Location = New System.Drawing.Point(310, 9)
        Me.LblCalc.Name = "LblCalc"
        Me.LblCalc.Size = New System.Drawing.Size(49, 13)
        Me.LblCalc.TabIndex = 8
        Me.LblCalc.TabStop = True
        Me.LblCalc.Text = "Calc Amt"
        '
        'LblSaveAmtToAL
        '
        Me.LblSaveAmtToAL.AutoSize = True
        Me.LblSaveAmtToAL.Location = New System.Drawing.Point(101, 231)
        Me.LblSaveAmtToAL.Name = "LblSaveAmtToAL"
        Me.LblSaveAmtToAL.Size = New System.Drawing.Size(32, 13)
        Me.LblSaveAmtToAL.TabIndex = 16
        Me.LblSaveAmtToAL.TabStop = True
        Me.LblSaveAmtToAL.Text = "Save"
        '
        'InvChecker
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.CmbBSPAL)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.CmbINVType)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmbMonth)
        Me.Controls.Add(Me.LblCalc)
        Me.Controls.Add(Me.CmbAL)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "InvChecker"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS 12 :. InvChecker"
        CType(Me.GridHD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        CType(Me.GridTRXwoINV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.GridTKT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridFOP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridRCP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents CmbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GridHD As System.Windows.Forms.DataGridView
    Friend WithEvents CmbINVType As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents LblUpdatePickUp As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbBSPAL As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LblCalc As System.Windows.Forms.LinkLabel
    Friend WithEvents Txt2AL_VND As System.Windows.Forms.TextBox
    Friend WithEvents TxtTTLINV As System.Windows.Forms.TextBox
    Friend WithEvents TxtDiff_VND As System.Windows.Forms.TextBox
    Friend WithEvents LblView As System.Windows.Forms.LinkLabel
    Friend WithEvents GridTKT As System.Windows.Forms.DataGridView
    Friend WithEvents GridFOP As System.Windows.Forms.DataGridView
    Friend WithEvents GridRCP As System.Windows.Forms.DataGridView
    Friend WithEvents LblAdj As System.Windows.Forms.LinkLabel
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents LblUpdateTKT As System.Windows.Forms.LinkLabel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtTax As System.Windows.Forms.TextBox
    Friend WithEvents txtFare As System.Windows.Forms.TextBox
    Friend WithEvents TxtRTG As System.Windows.Forms.TextBox
    Friend WithEvents txtTKNO As System.Windows.Forms.TextBox
    Friend WithEvents CmbCurr As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtVND As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtComm As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents LblLoadGridTRXwoINV As System.Windows.Forms.LinkLabel
    Friend WithEvents GridTRXwoINV As System.Windows.Forms.DataGridView
    Friend WithEvents LblIssueINV As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtComm_VND As System.Windows.Forms.TextBox
    Friend WithEvents TxtInvFare As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtADM_VND As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TxtInvTaxCharge As System.Windows.Forms.TextBox
    Friend WithEvents CmbRoundingAmt As System.Windows.Forms.ComboBox
    Friend WithEvents LblRounding As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUpdateInfor As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtTTLINVByS1 As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtQty As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents LblSaveAmtToAL As System.Windows.Forms.LinkLabel
End Class
