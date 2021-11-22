<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Payment4VCR
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.CmbVCRNo = New System.Windows.Forms.ComboBox()
        Me.GridFOP = New System.Windows.Forms.DataGridView()
        Me.FOP = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Currency = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ROE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Amount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Document = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LblUpdate = New System.Windows.Forms.LinkLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtINVND = New System.Windows.Forms.TextBox()
        Me.TxtAmt = New System.Windows.Forms.TextBox()
        Me.TxtPaid = New System.Windows.Forms.TextBox()
        Me.LblSearch = New System.Windows.Forms.LinkLabel()
        Me.LblCurr = New System.Windows.Forms.Label()
        Me.TxtPayer = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LblNotice = New System.Windows.Forms.Label()
        Me.CmbPrj = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.OptVCR = New System.Windows.Forms.RadioButton()
        Me.OptPrj = New System.Windows.Forms.RadioButton()
        Me.GrdVCRinPrj = New System.Windows.Forms.DataGridView()
        Me.LblNoOfVCR = New System.Windows.Forms.Label()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GridDEB = New System.Windows.Forms.DataGridView()
        Me.TxtNoteClearDeb = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.LblClearDEB = New System.Windows.Forms.LinkLabel()
        CType(Me.GridFOP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GrdVCRinPrj, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.GridDEB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CmbVCRNo
        '
        Me.CmbVCRNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbVCRNo.Enabled = False
        Me.CmbVCRNo.FormattingEnabled = True
        Me.CmbVCRNo.Location = New System.Drawing.Point(224, 24)
        Me.CmbVCRNo.Name = "CmbVCRNo"
        Me.CmbVCRNo.Size = New System.Drawing.Size(133, 21)
        Me.CmbVCRNo.TabIndex = 0
        '
        'GridFOP
        '
        Me.GridFOP.BackgroundColor = System.Drawing.Color.AliceBlue
        Me.GridFOP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridFOP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FOP, Me.Currency, Me.ROE, Me.Amount, Me.Document})
        Me.GridFOP.Location = New System.Drawing.Point(3, 32)
        Me.GridFOP.Name = "GridFOP"
        Me.GridFOP.RowHeadersWidth = 25
        Me.GridFOP.Size = New System.Drawing.Size(408, 135)
        Me.GridFOP.TabIndex = 33
        '
        'FOP
        '
        Me.FOP.HeaderText = "FOP"
        Me.FOP.Name = "FOP"
        Me.FOP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.FOP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.FOP.Width = 56
        '
        'Currency
        '
        Me.Currency.HeaderText = "Curr"
        Me.Currency.Name = "Currency"
        Me.Currency.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Currency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Currency.Width = 56
        '
        'ROE
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.Format = "#,##0"
        Me.ROE.DefaultCellStyle = DataGridViewCellStyle3
        Me.ROE.HeaderText = "ROE"
        Me.ROE.Name = "ROE"
        Me.ROE.Width = 75
        '
        'Amount
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.Format = "#,##0"
        Me.Amount.DefaultCellStyle = DataGridViewCellStyle4
        Me.Amount.HeaderText = "Amount"
        Me.Amount.Name = "Amount"
        Me.Amount.Width = 80
        '
        'Document
        '
        Me.Document.HeaderText = "Document"
        Me.Document.Name = "Document"
        '
        'LblUpdate
        '
        Me.LblUpdate.AutoSize = True
        Me.LblUpdate.Location = New System.Drawing.Point(369, 176)
        Me.LblUpdate.Name = "LblUpdate"
        Me.LblUpdate.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdate.TabIndex = 34
        Me.LblUpdate.TabStop = True
        Me.LblUpdate.Text = "Update"
        Me.LblUpdate.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(167, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "VCR No."
        '
        'TxtINVND
        '
        Me.TxtINVND.Location = New System.Drawing.Point(144, 6)
        Me.TxtINVND.Name = "TxtINVND"
        Me.TxtINVND.ReadOnly = True
        Me.TxtINVND.Size = New System.Drawing.Size(39, 20)
        Me.TxtINVND.TabIndex = 36
        Me.TxtINVND.Visible = False
        '
        'TxtAmt
        '
        Me.TxtAmt.Location = New System.Drawing.Point(39, 6)
        Me.TxtAmt.Name = "TxtAmt"
        Me.TxtAmt.ReadOnly = True
        Me.TxtAmt.Size = New System.Drawing.Size(99, 20)
        Me.TxtAmt.TabIndex = 36
        Me.TxtAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPaid
        '
        Me.TxtPaid.Location = New System.Drawing.Point(189, 6)
        Me.TxtPaid.Name = "TxtPaid"
        Me.TxtPaid.Size = New System.Drawing.Size(39, 20)
        Me.TxtPaid.TabIndex = 36
        Me.TxtPaid.Visible = False
        '
        'LblSearch
        '
        Me.LblSearch.AutoSize = True
        Me.LblSearch.Location = New System.Drawing.Point(227, 3)
        Me.LblSearch.Name = "LblSearch"
        Me.LblSearch.Size = New System.Drawing.Size(41, 13)
        Me.LblSearch.TabIndex = 37
        Me.LblSearch.TabStop = True
        Me.LblSearch.Text = "Search"
        '
        'LblCurr
        '
        Me.LblCurr.AutoSize = True
        Me.LblCurr.Location = New System.Drawing.Point(3, 9)
        Me.LblCurr.Name = "LblCurr"
        Me.LblCurr.Size = New System.Drawing.Size(30, 13)
        Me.LblCurr.TabIndex = 38
        Me.LblCurr.Text = "VND"
        '
        'TxtPayer
        '
        Me.TxtPayer.Location = New System.Drawing.Point(3, 173)
        Me.TxtPayer.Name = "TxtPayer"
        Me.TxtPayer.Size = New System.Drawing.Size(171, 20)
        Me.TxtPayer.TabIndex = 39
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 145)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 13)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Payer"
        '
        'LblNotice
        '
        Me.LblNotice.AutoSize = True
        Me.LblNotice.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblNotice.ForeColor = System.Drawing.Color.Red
        Me.LblNotice.Location = New System.Drawing.Point(6, 210)
        Me.LblNotice.Name = "LblNotice"
        Me.LblNotice.Size = New System.Drawing.Size(396, 13)
        Me.LblNotice.TabIndex = 40
        Me.LblNotice.Text = "Be Careful Before Clicking On UPDATE. You Cant Undo This Action!"
        Me.LblNotice.Visible = False
        '
        'CmbPrj
        '
        Me.CmbPrj.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbPrj.Enabled = False
        Me.CmbPrj.FormattingEnabled = True
        Me.CmbPrj.Location = New System.Drawing.Point(45, 24)
        Me.CmbPrj.Name = "CmbPrj"
        Me.CmbPrj.Size = New System.Drawing.Size(116, 21)
        Me.CmbPrj.TabIndex = 41
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(-1, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 13)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "Project"
        '
        'OptVCR
        '
        Me.OptVCR.AutoSize = True
        Me.OptVCR.Location = New System.Drawing.Point(121, 1)
        Me.OptVCR.Name = "OptVCR"
        Me.OptVCR.Size = New System.Drawing.Size(100, 17)
        Me.OptVCR.TabIndex = 42
        Me.OptVCR.TabStop = True
        Me.OptVCR.Text = "Pay4SingleVCR"
        Me.OptVCR.UseVisualStyleBackColor = True
        '
        'OptPrj
        '
        Me.OptPrj.AutoSize = True
        Me.OptPrj.Location = New System.Drawing.Point(2, 1)
        Me.OptPrj.Name = "OptPrj"
        Me.OptPrj.Size = New System.Drawing.Size(113, 17)
        Me.OptPrj.TabIndex = 42
        Me.OptPrj.TabStop = True
        Me.OptPrj.Text = "Pay4WholeProject"
        Me.OptPrj.UseVisualStyleBackColor = True
        '
        'GrdVCRinPrj
        '
        Me.GrdVCRinPrj.AllowUserToAddRows = False
        Me.GrdVCRinPrj.AllowUserToDeleteRows = False
        Me.GrdVCRinPrj.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GrdVCRinPrj.Location = New System.Drawing.Point(2, 51)
        Me.GrdVCRinPrj.Name = "GrdVCRinPrj"
        Me.GrdVCRinPrj.RowHeadersVisible = False
        Me.GrdVCRinPrj.Size = New System.Drawing.Size(358, 426)
        Me.GrdVCRinPrj.TabIndex = 43
        '
        'LblNoOfVCR
        '
        Me.LblNoOfVCR.AutoSize = True
        Me.LblNoOfVCR.Location = New System.Drawing.Point(2, 480)
        Me.LblNoOfVCR.Name = "LblNoOfVCR"
        Me.LblNoOfVCR.Size = New System.Drawing.Size(113, 13)
        Me.LblNoOfVCR.TabIndex = 44
        Me.LblNoOfVCR.Text = "No. Of VCR in Project:"
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(319, 480)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 45
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        Me.LblDelete.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(363, 3)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(419, 474)
        Me.TabControl1.TabIndex = 46
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GridFOP)
        Me.TabPage1.Controls.Add(Me.LblUpdate)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.TxtPayer)
        Me.TabPage1.Controls.Add(Me.TxtAmt)
        Me.TabPage1.Controls.Add(Me.TxtINVND)
        Me.TabPage1.Controls.Add(Me.TxtPaid)
        Me.TabPage1.Controls.Add(Me.LblNotice)
        Me.TabPage1.Controls.Add(Me.LblCurr)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(411, 448)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Get Payment"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.LblClearDEB)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.TxtNoteClearDeb)
        Me.TabPage2.Controls.Add(Me.GridDEB)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(411, 448)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Pending DEB"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GridDEB
        '
        Me.GridDEB.AllowUserToAddRows = False
        Me.GridDEB.AllowUserToDeleteRows = False
        Me.GridDEB.AllowUserToResizeColumns = False
        Me.GridDEB.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridDEB.Location = New System.Drawing.Point(3, 2)
        Me.GridDEB.Name = "GridDEB"
        Me.GridDEB.RowHeadersVisible = False
        Me.GridDEB.Size = New System.Drawing.Size(405, 277)
        Me.GridDEB.TabIndex = 0
        '
        'TxtNoteClearDeb
        '
        Me.TxtNoteClearDeb.Location = New System.Drawing.Point(32, 287)
        Me.TxtNoteClearDeb.Name = "TxtNoteClearDeb"
        Me.TxtNoteClearDeb.Size = New System.Drawing.Size(173, 20)
        Me.TxtNoteClearDeb.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(2, 290)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Note"
        '
        'LblClearDEB
        '
        Me.LblClearDEB.AutoSize = True
        Me.LblClearDEB.Location = New System.Drawing.Point(206, 290)
        Me.LblClearDEB.Name = "LblClearDEB"
        Me.LblClearDEB.Size = New System.Drawing.Size(53, 13)
        Me.LblClearDEB.TabIndex = 3
        Me.LblClearDEB.TabStop = True
        Me.LblClearDEB.Text = "ClearDEB"
        Me.LblClearDEB.Visible = False
        '
        'Payment4VCR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.LblNoOfVCR)
        Me.Controls.Add(Me.GrdVCRinPrj)
        Me.Controls.Add(Me.OptPrj)
        Me.Controls.Add(Me.OptVCR)
        Me.Controls.Add(Me.CmbPrj)
        Me.Controls.Add(Me.LblSearch)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmbVCRNo)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Payment4VCR"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: Flex Tour 10 :. Payment for Voucher"
        CType(Me.GridFOP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GrdVCRinPrj, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.GridDEB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CmbVCRNo As System.Windows.Forms.ComboBox
    Friend WithEvents GridFOP As System.Windows.Forms.DataGridView
    Friend WithEvents FOP As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Currency As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ROE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Document As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtINVND As System.Windows.Forms.TextBox
    Friend WithEvents TxtAmt As System.Windows.Forms.TextBox
    Friend WithEvents TxtPaid As System.Windows.Forms.TextBox
    Friend WithEvents LblSearch As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCurr As System.Windows.Forms.Label
    Friend WithEvents TxtPayer As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblNotice As System.Windows.Forms.Label
    Friend WithEvents CmbPrj As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents OptVCR As System.Windows.Forms.RadioButton
    Friend WithEvents OptPrj As System.Windows.Forms.RadioButton
    Friend WithEvents GrdVCRinPrj As System.Windows.Forms.DataGridView
    Friend WithEvents LblNoOfVCR As System.Windows.Forms.Label
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GridDEB As System.Windows.Forms.DataGridView
    Friend WithEvents LblClearDEB As System.Windows.Forms.LinkLabel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtNoteClearDeb As System.Windows.Forms.TextBox
End Class
