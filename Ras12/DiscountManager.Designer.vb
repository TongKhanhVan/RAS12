<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DiscountManager
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
        Me.CmbNoOfVCR = New System.Windows.Forms.ComboBox()
        Me.LblUpdateVCRAMT = New System.Windows.Forms.LinkLabel()
        Me.TxtCondi = New System.Windows.Forms.RichTextBox()
        Me.TxtValidThru = New System.Windows.Forms.DateTimePicker()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GridVCR = New System.Windows.Forms.DataGridView()
        Me.S = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.CmbCurr = New System.Windows.Forms.ComboBox()
        Me.TxtAmt = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CmbType = New System.Windows.Forms.ComboBox()
        Me.txtPrjCode = New System.Windows.Forms.TextBox()
        Me.TxtRMK = New System.Windows.Forms.TextBox()
        Me.LblBlock = New System.Windows.Forms.LinkLabel()
        Me.TxtBlockAmt = New System.Windows.Forms.TextBox()
        Me.CmbNoValue = New System.Windows.Forms.ComboBox()
        Me.LbLFilter = New System.Windows.Forms.LinkLabel()
        Me.txtTenkhach = New System.Windows.Forms.TextBox()
        Me.cboFilter = New System.Windows.Forms.ComboBox()
        Me.txtMAX_OB = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboTerritory = New System.Windows.Forms.ComboBox()
        Me.cboSVC = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtKickBack = New System.Windows.Forms.TextBox()
        CType(Me.GridVCR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CmbNoOfVCR
        '
        Me.CmbNoOfVCR.FormattingEnabled = True
        Me.CmbNoOfVCR.Items.AddRange(New Object() {"01", "10", "20", "30"})
        Me.CmbNoOfVCR.Location = New System.Drawing.Point(213, 30)
        Me.CmbNoOfVCR.Name = "CmbNoOfVCR"
        Me.CmbNoOfVCR.Size = New System.Drawing.Size(35, 21)
        Me.CmbNoOfVCR.TabIndex = 10
        '
        'LblUpdateVCRAMT
        '
        Me.LblUpdateVCRAMT.AutoSize = True
        Me.LblUpdateVCRAMT.Location = New System.Drawing.Point(402, 9)
        Me.LblUpdateVCRAMT.Name = "LblUpdateVCRAMT"
        Me.LblUpdateVCRAMT.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdateVCRAMT.TabIndex = 7
        Me.LblUpdateVCRAMT.TabStop = True
        Me.LblUpdateVCRAMT.Text = "Update"
        '
        'TxtCondi
        '
        Me.TxtCondi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtCondi.Location = New System.Drawing.Point(314, 31)
        Me.TxtCondi.Name = "TxtCondi"
        Me.TxtCondi.Size = New System.Drawing.Size(473, 25)
        Me.TxtCondi.TabIndex = 5
        Me.TxtCondi.Text = ""
        '
        'TxtValidThru
        '
        Me.TxtValidThru.CustomFormat = "dd-MMM-yy"
        Me.TxtValidThru.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtValidThru.Location = New System.Drawing.Point(314, 6)
        Me.TxtValidThru.Name = "TxtValidThru"
        Me.TxtValidThru.Size = New System.Drawing.Size(82, 20)
        Me.TxtValidThru.TabIndex = 4
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(142, 34)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(65, 13)
        Me.Label12.TabIndex = 1
        Me.Label12.Text = "No Of VCRs"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(254, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Valid Until"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(254, 34)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Condition"
        '
        'GridVCR
        '
        Me.GridVCR.AllowUserToAddRows = False
        Me.GridVCR.AllowUserToDeleteRows = False
        Me.GridVCR.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GridVCR.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridVCR.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridVCR.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.S})
        Me.GridVCR.Location = New System.Drawing.Point(1, 97)
        Me.GridVCR.Name = "GridVCR"
        Me.GridVCR.RowHeadersVisible = False
        Me.GridVCR.Size = New System.Drawing.Size(786, 382)
        Me.GridVCR.TabIndex = 3
        '
        'S
        '
        Me.S.HeaderText = "S"
        Me.S.Name = "S"
        Me.S.Width = 32
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(749, 484)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 4
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        Me.LblDelete.Visible = False
        '
        'CmbCurr
        '
        Me.CmbCurr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCurr.FormattingEnabled = True
        Me.CmbCurr.Items.AddRange(New Object() {"USD", "VND"})
        Me.CmbCurr.Location = New System.Drawing.Point(145, 6)
        Me.CmbCurr.Name = "CmbCurr"
        Me.CmbCurr.Size = New System.Drawing.Size(47, 21)
        Me.CmbCurr.TabIndex = 1
        '
        'TxtAmt
        '
        Me.TxtAmt.Location = New System.Drawing.Point(56, 6)
        Me.TxtAmt.Name = "TxtAmt"
        Me.TxtAmt.Size = New System.Drawing.Size(83, 20)
        Me.TxtAmt.TabIndex = 0
        Me.TxtAmt.Text = "0"
        Me.TxtAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(-1, 34)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(54, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "VCRCode"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(-1, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Amount"
        '
        'CmbType
        '
        Me.CmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbType.Enabled = False
        Me.CmbType.FormattingEnabled = True
        Me.CmbType.Items.AddRange(New Object() {"SELL", "FOC"})
        Me.CmbType.Location = New System.Drawing.Point(737, 6)
        Me.CmbType.Name = "CmbType"
        Me.CmbType.Size = New System.Drawing.Size(50, 21)
        Me.CmbType.TabIndex = 3
        Me.CmbType.Visible = False
        '
        'txtPrjCode
        '
        Me.txtPrjCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPrjCode.Location = New System.Drawing.Point(56, 31)
        Me.txtPrjCode.Name = "txtPrjCode"
        Me.txtPrjCode.Size = New System.Drawing.Size(83, 20)
        Me.txtPrjCode.TabIndex = 52
        '
        'TxtRMK
        '
        Me.TxtRMK.Location = New System.Drawing.Point(336, 481)
        Me.TxtRMK.Name = "TxtRMK"
        Me.TxtRMK.Size = New System.Drawing.Size(242, 20)
        Me.TxtRMK.TabIndex = 53
        '
        'LblBlock
        '
        Me.LblBlock.AutoSize = True
        Me.LblBlock.Location = New System.Drawing.Point(197, 484)
        Me.LblBlock.Name = "LblBlock"
        Me.LblBlock.Size = New System.Drawing.Size(34, 13)
        Me.LblBlock.TabIndex = 54
        Me.LblBlock.TabStop = True
        Me.LblBlock.Text = "Block"
        Me.LblBlock.Visible = False
        '
        'TxtBlockAmt
        '
        Me.TxtBlockAmt.Location = New System.Drawing.Point(237, 481)
        Me.TxtBlockAmt.Name = "TxtBlockAmt"
        Me.TxtBlockAmt.Size = New System.Drawing.Size(93, 20)
        Me.TxtBlockAmt.TabIndex = 0
        Me.TxtBlockAmt.Text = "0"
        Me.TxtBlockAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'CmbNoValue
        '
        Me.CmbNoValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbNoValue.FormattingEnabled = True
        Me.CmbNoValue.Items.AddRange(New Object() {"", "#N/A#"})
        Me.CmbNoValue.Location = New System.Drawing.Point(194, 6)
        Me.CmbNoValue.Name = "CmbNoValue"
        Me.CmbNoValue.Size = New System.Drawing.Size(54, 21)
        Me.CmbNoValue.TabIndex = 55
        '
        'LbLFilter
        '
        Me.LbLFilter.AutoSize = True
        Me.LbLFilter.Location = New System.Drawing.Point(162, 484)
        Me.LbLFilter.Name = "LbLFilter"
        Me.LbLFilter.Size = New System.Drawing.Size(29, 13)
        Me.LbLFilter.TabIndex = 56
        Me.LbLFilter.TabStop = True
        Me.LbLFilter.Text = "Filter"
        '
        'txtTenkhach
        '
        Me.txtTenkhach.Location = New System.Drawing.Point(581, 481)
        Me.txtTenkhach.Name = "txtTenkhach"
        Me.txtTenkhach.Size = New System.Drawing.Size(165, 20)
        Me.txtTenkhach.TabIndex = 53
        '
        'cboFilter
        '
        Me.cboFilter.FormattingEnabled = True
        Me.cboFilter.Items.AddRange(New Object() {"OK only", "XX only", "BK only", "Expired"})
        Me.cboFilter.Location = New System.Drawing.Point(1, 480)
        Me.cboFilter.Name = "cboFilter"
        Me.cboFilter.Size = New System.Drawing.Size(126, 21)
        Me.cboFilter.TabIndex = 57
        '
        'txtMAX_OB
        '
        Me.txtMAX_OB.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMAX_OB.Location = New System.Drawing.Point(313, 57)
        Me.txtMAX_OB.Name = "txtMAX_OB"
        Me.txtMAX_OB.Size = New System.Drawing.Size(47, 20)
        Me.txtMAX_OB.TabIndex = 59
        Me.txtMAX_OB.Text = "0.9"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(-2, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 58
        Me.Label1.Text = "Territory"
        '
        'cboTerritory
        '
        Me.cboTerritory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTerritory.FormattingEnabled = True
        Me.cboTerritory.Items.AddRange(New Object() {"SGN", "HAN", "ALL"})
        Me.cboTerritory.Location = New System.Drawing.Point(56, 52)
        Me.cboTerritory.Name = "cboTerritory"
        Me.cboTerritory.Size = New System.Drawing.Size(83, 21)
        Me.cboTerritory.TabIndex = 60
        '
        'cboSVC
        '
        Me.cboSVC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSVC.FormattingEnabled = True
        Me.cboSVC.Items.AddRange(New Object() {"RAS", "FLX", "ALL"})
        Me.cboSVC.Location = New System.Drawing.Point(194, 52)
        Me.cboSVC.Name = "cboSVC"
        Me.cboSVC.Size = New System.Drawing.Size(54, 21)
        Me.cboSVC.TabIndex = 62
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(142, 60)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(28, 13)
        Me.Label6.TabIndex = 61
        Me.Label6.Text = "SVC"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(254, 60)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(51, 13)
        Me.Label7.TabIndex = 63
        Me.Label7.Text = "MAX_OB"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(366, 60)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(60, 13)
        Me.Label8.TabIndex = 65
        Me.Label8.Text = "%Kickback"
        '
        'txtKickBack
        '
        Me.txtKickBack.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtKickBack.Location = New System.Drawing.Point(425, 57)
        Me.txtKickBack.Name = "txtKickBack"
        Me.txtKickBack.Size = New System.Drawing.Size(47, 20)
        Me.txtKickBack.TabIndex = 64
        Me.txtKickBack.Text = "0"
        '
        'DiscountManager
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtKickBack)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.cboSVC)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cboTerritory)
        Me.Controls.Add(Me.txtMAX_OB)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboFilter)
        Me.Controls.Add(Me.LbLFilter)
        Me.Controls.Add(Me.CmbNoValue)
        Me.Controls.Add(Me.LblBlock)
        Me.Controls.Add(Me.txtTenkhach)
        Me.Controls.Add(Me.TxtRMK)
        Me.Controls.Add(Me.CmbNoOfVCR)
        Me.Controls.Add(Me.txtPrjCode)
        Me.Controls.Add(Me.CmbType)
        Me.Controls.Add(Me.LblUpdateVCRAMT)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtCondi)
        Me.Controls.Add(Me.TxtValidThru)
        Me.Controls.Add(Me.GridVCR)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtBlockAmt)
        Me.Controls.Add(Me.TxtAmt)
        Me.Controls.Add(Me.CmbCurr)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DiscountManager"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS 12 :. Discount Manager"
        CType(Me.GridVCR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtCondi As System.Windows.Forms.RichTextBox
    Friend WithEvents TxtValidThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LblUpdateVCRAMT As System.Windows.Forms.LinkLabel
    Friend WithEvents GridVCR As System.Windows.Forms.DataGridView
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbNoOfVCR As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents CmbCurr As System.Windows.Forms.ComboBox
    Friend WithEvents TxtAmt As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CmbType As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrjCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtRMK As System.Windows.Forms.TextBox
    Friend WithEvents LblBlock As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtBlockAmt As System.Windows.Forms.TextBox
    Friend WithEvents CmbNoValue As System.Windows.Forms.ComboBox
    Friend WithEvents S As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents LbLFilter As System.Windows.Forms.LinkLabel
    Friend WithEvents txtTenkhach As System.Windows.Forms.TextBox
    Friend WithEvents cboFilter As System.Windows.Forms.ComboBox
    Friend WithEvents txtMAX_OB As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboTerritory As System.Windows.Forms.ComboBox
    Friend WithEvents cboSVC As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtKickBack As System.Windows.Forms.TextBox
End Class
