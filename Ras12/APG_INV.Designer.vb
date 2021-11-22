<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class APG_INV
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.txtCustAddrr = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.txtTaxCode = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.txtCustName = New System.Windows.Forms.TextBox()
        Me.CmbCustomer = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CmbAL = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.LblRePrint = New System.Windows.Forms.LinkLabel()
        Me.GridInvHandlerInv = New System.Windows.Forms.DataGridView()
        Me.LblPreviewInv = New System.Windows.Forms.LinkLabel()
        Me.LblSave = New System.Windows.Forms.LinkLabel()
        Me.TxtNoiDung = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtAmt = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TxtDOI = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CmbYear = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lbkReserveInvNo = New System.Windows.Forms.LinkLabel()
        Me.cboStatus = New System.Windows.Forms.ComboBox()
        Me.chkShowLast3Month = New System.Windows.Forms.CheckBox()
        CType(Me.GridInvHandlerInv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(75, 32)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(53, 13)
        Me.Label29.TabIndex = 53
        Me.Label29.Text = "Tax Code"
        '
        'txtCustAddrr
        '
        Me.txtCustAddrr.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustAddrr.Location = New System.Drawing.Point(295, 29)
        Me.txtCustAddrr.MaxLength = 0
        Me.txtCustAddrr.Name = "txtCustAddrr"
        Me.txtCustAddrr.ReadOnly = True
        Me.txtCustAddrr.Size = New System.Drawing.Size(492, 20)
        Me.txtCustAddrr.TabIndex = 4
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(244, 32)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(45, 13)
        Me.Label28.TabIndex = 52
        Me.Label28.Text = "Address"
        '
        'txtTaxCode
        '
        Me.txtTaxCode.Location = New System.Drawing.Point(133, 29)
        Me.txtTaxCode.MaxLength = 20
        Me.txtTaxCode.Name = "txtTaxCode"
        Me.txtTaxCode.ReadOnly = True
        Me.txtTaxCode.Size = New System.Drawing.Size(111, 20)
        Me.txtTaxCode.TabIndex = 3
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(244, 11)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(51, 13)
        Me.Label30.TabIndex = 54
        Me.Label30.Text = "FullName"
        '
        'txtCustName
        '
        Me.txtCustName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustName.Location = New System.Drawing.Point(295, 8)
        Me.txtCustName.MaxLength = 256
        Me.txtCustName.Name = "txtCustName"
        Me.txtCustName.ReadOnly = True
        Me.txtCustName.Size = New System.Drawing.Size(492, 20)
        Me.txtCustName.TabIndex = 2
        '
        'CmbCustomer
        '
        Me.CmbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCustomer.FormattingEnabled = True
        Me.CmbCustomer.Location = New System.Drawing.Point(133, 6)
        Me.CmbCustomer.Name = "CmbCustomer"
        Me.CmbCustomer.Size = New System.Drawing.Size(111, 23)
        Me.CmbCustomer.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(75, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 15)
        Me.Label6.TabIndex = 56
        Me.Label6.Text = "Customer"
        '
        'CmbAL
        '
        Me.CmbAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(32, 8)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(41, 21)
        Me.CmbAL.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(2, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(21, 15)
        Me.Label1.TabIndex = 56
        Me.Label1.Text = "AL"
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(554, 486)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 61
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        Me.LblDelete.Visible = False
        '
        'LblRePrint
        '
        Me.LblRePrint.AutoSize = True
        Me.LblRePrint.Location = New System.Drawing.Point(64, 484)
        Me.LblRePrint.Name = "LblRePrint"
        Me.LblRePrint.Size = New System.Drawing.Size(42, 13)
        Me.LblRePrint.TabIndex = 62
        Me.LblRePrint.TabStop = True
        Me.LblRePrint.Text = "RePrint"
        Me.LblRePrint.Visible = False
        '
        'GridInvHandlerInv
        '
        Me.GridInvHandlerInv.AllowUserToAddRows = False
        Me.GridInvHandlerInv.AllowUserToDeleteRows = False
        Me.GridInvHandlerInv.BackgroundColor = System.Drawing.Color.Beige
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridInvHandlerInv.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.GridInvHandlerInv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridInvHandlerInv.Location = New System.Drawing.Point(0, 73)
        Me.GridInvHandlerInv.Name = "GridInvHandlerInv"
        Me.GridInvHandlerInv.ReadOnly = True
        Me.GridInvHandlerInv.RowHeadersVisible = False
        Me.GridInvHandlerInv.Size = New System.Drawing.Size(788, 409)
        Me.GridInvHandlerInv.TabIndex = 5
        '
        'LblPreviewInv
        '
        Me.LblPreviewInv.AutoSize = True
        Me.LblPreviewInv.Location = New System.Drawing.Point(2, 484)
        Me.LblPreviewInv.Name = "LblPreviewInv"
        Me.LblPreviewInv.Size = New System.Drawing.Size(45, 13)
        Me.LblPreviewInv.TabIndex = 58
        Me.LblPreviewInv.TabStop = True
        Me.LblPreviewInv.Text = "Preview"
        Me.LblPreviewInv.Visible = False
        '
        'LblSave
        '
        Me.LblSave.AutoSize = True
        Me.LblSave.Location = New System.Drawing.Point(754, 53)
        Me.LblSave.Name = "LblSave"
        Me.LblSave.Size = New System.Drawing.Size(32, 13)
        Me.LblSave.TabIndex = 7
        Me.LblSave.TabStop = True
        Me.LblSave.Text = "Save"
        '
        'TxtNoiDung
        '
        Me.TxtNoiDung.Location = New System.Drawing.Point(295, 50)
        Me.TxtNoiDung.MaxLength = 56
        Me.TxtNoiDung.Name = "TxtNoiDung"
        Me.TxtNoiDung.Size = New System.Drawing.Size(333, 20)
        Me.TxtNoiDung.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(75, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 52
        Me.Label2.Text = "VND"
        '
        'TxtAmt
        '
        Me.TxtAmt.Location = New System.Drawing.Point(133, 50)
        Me.TxtAmt.MaxLength = 20
        Me.TxtAmt.Name = "TxtAmt"
        Me.TxtAmt.Size = New System.Drawing.Size(111, 20)
        Me.TxtAmt.TabIndex = 5
        Me.TxtAmt.Text = "0"
        Me.TxtAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(244, 53)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 13)
        Me.Label3.TabIndex = 52
        Me.Label3.Text = "Descrt."
        '
        'TxtDOI
        '
        Me.TxtDOI.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtDOI.Location = New System.Drawing.Point(664, 50)
        Me.TxtDOI.Name = "TxtDOI"
        Me.TxtDOI.Size = New System.Drawing.Size(84, 20)
        Me.TxtDOI.TabIndex = 65
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(635, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(28, 15)
        Me.Label4.TabIndex = 56
        Me.Label4.Text = "DOI"
        '
        'CmbYear
        '
        Me.CmbYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbYear.FormattingEnabled = True
        Me.CmbYear.Location = New System.Drawing.Point(32, 49)
        Me.CmbYear.Name = "CmbYear"
        Me.CmbYear.Size = New System.Drawing.Size(41, 21)
        Me.CmbYear.TabIndex = 66
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(2, 51)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(32, 15)
        Me.Label5.TabIndex = 56
        Me.Label5.Text = "Year"
        '
        'lbkReserveInvNo
        '
        Me.lbkReserveInvNo.AutoSize = True
        Me.lbkReserveInvNo.Location = New System.Drawing.Point(481, 486)
        Me.lbkReserveInvNo.Name = "lbkReserveInvNo"
        Me.lbkReserveInvNo.Size = New System.Drawing.Size(76, 13)
        Me.lbkReserveInvNo.TabIndex = 67
        Me.lbkReserveInvNo.TabStop = True
        Me.lbkReserveInvNo.Text = "ReserveInvNo"
        Me.lbkReserveInvNo.Visible = False
        '
        'cboStatus
        '
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Items.AddRange(New Object() {"OK", "XX", "QQ"})
        Me.cboStatus.Location = New System.Drawing.Point(729, 483)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(58, 21)
        Me.cboStatus.TabIndex = 68
        '
        'chkShowLast3Month
        '
        Me.chkShowLast3Month.AutoSize = True
        Me.chkShowLast3Month.Checked = True
        Me.chkShowLast3Month.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLast3Month.Location = New System.Drawing.Point(593, 485)
        Me.chkShowLast3Month.Name = "chkShowLast3Month"
        Me.chkShowLast3Month.Size = New System.Drawing.Size(130, 17)
        Me.chkShowLast3Month.TabIndex = 69
        Me.chkShowLast3Month.Text = "ShowLast3MonthOnly"
        Me.chkShowLast3Month.UseVisualStyleBackColor = True
        '
        'APG_INV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.chkShowLast3Month)
        Me.Controls.Add(Me.cboStatus)
        Me.Controls.Add(Me.lbkReserveInvNo)
        Me.Controls.Add(Me.CmbYear)
        Me.Controls.Add(Me.TxtDOI)
        Me.Controls.Add(Me.TxtNoiDung)
        Me.Controls.Add(Me.LblSave)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.LblRePrint)
        Me.Controls.Add(Me.GridInvHandlerInv)
        Me.Controls.Add(Me.LblPreviewInv)
        Me.Controls.Add(Me.CmbAL)
        Me.Controls.Add(Me.CmbCustomer)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.txtCustAddrr)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.TxtAmt)
        Me.Controls.Add(Me.txtTaxCode)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.txtCustName)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "APG_INV"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. APG Invoices"
        CType(Me.GridInvHandlerInv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents txtCustAddrr As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents txtTaxCode As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents txtCustName As System.Windows.Forms.TextBox
    Friend WithEvents CmbCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LblRePrint As System.Windows.Forms.LinkLabel
    Friend WithEvents GridInvHandlerInv As System.Windows.Forms.DataGridView
    Friend WithEvents LblPreviewInv As System.Windows.Forms.LinkLabel
    Friend WithEvents LblSave As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtNoiDung As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtAmt As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtDOI As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CmbYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lbkReserveInvNo As System.Windows.Forms.LinkLabel
    Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents chkShowLast3Month As System.Windows.Forms.CheckBox
End Class
