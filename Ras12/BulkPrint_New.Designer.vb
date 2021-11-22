<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BulkPrint_New
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
        Me.GridTKT = New System.Windows.Forms.DataGridView()
        Me.S = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.LblPreview = New System.Windows.Forms.LinkLabel()
        Me.cmbCust = New System.Windows.Forms.ComboBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.LblLoadTKT = New System.Windows.Forms.LinkLabel()
        Me.LblCheckAll_unCheckAll_AL = New System.Windows.Forms.LinkLabel()
        Me.TxtDOIto = New System.Windows.Forms.DateTimePicker()
        Me.txtDOIFrm = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtTCocde = New System.Windows.Forms.TextBox()
        Me.CmbCounter = New System.Windows.Forms.ComboBox()
        Me.LstAL = New System.Windows.Forms.CheckedListBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.LblCalcRASInvAmt = New System.Windows.Forms.LinkLabel()
        Me.LBlChk_UnChkAll_TKT = New System.Windows.Forms.LinkLabel()
        Me.txtSummary = New System.Windows.Forms.RichTextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtInvDate = New System.Windows.Forms.DateTimePicker()
        Me.CmbPrintToWhom = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        CType(Me.GridTKT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridTKT
        '
        Me.GridTKT.AllowUserToAddRows = False
        Me.GridTKT.AllowUserToDeleteRows = False
        Me.GridTKT.BackgroundColor = System.Drawing.Color.Azure
        Me.GridTKT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTKT.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.S})
        Me.GridTKT.Location = New System.Drawing.Point(3, 3)
        Me.GridTKT.Name = "GridTKT"
        Me.GridTKT.RowHeadersVisible = False
        Me.GridTKT.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridTKT.Size = New System.Drawing.Size(774, 409)
        Me.GridTKT.TabIndex = 2
        '
        'S
        '
        Me.S.HeaderText = "S"
        Me.S.Name = "S"
        Me.S.Width = 24
        '
        'LblPreview
        '
        Me.LblPreview.AutoSize = True
        Me.LblPreview.Location = New System.Drawing.Point(234, 22)
        Me.LblPreview.Name = "LblPreview"
        Me.LblPreview.Size = New System.Drawing.Size(69, 13)
        Me.LblPreview.TabIndex = 3
        Me.LblPreview.TabStop = True
        Me.LblPreview.Text = "Print Preview"
        '
        'cmbCust
        '
        Me.cmbCust.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCust.FormattingEnabled = True
        Me.cmbCust.Location = New System.Drawing.Point(319, 3)
        Me.cmbCust.Name = "cmbCust"
        Me.cmbCust.Size = New System.Drawing.Size(146, 21)
        Me.cmbCust.TabIndex = 6
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(2, 1)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(785, 501)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.LblLoadTKT)
        Me.TabPage1.Controls.Add(Me.LblCheckAll_unCheckAll_AL)
        Me.TabPage1.Controls.Add(Me.TxtDOIto)
        Me.TabPage1.Controls.Add(Me.txtDOIFrm)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.txtTCocde)
        Me.TabPage1.Controls.Add(Me.CmbCounter)
        Me.TabPage1.Controls.Add(Me.LstAL)
        Me.TabPage1.Controls.Add(Me.cmbCust)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(777, 475)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "What2Print"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'LblLoadTKT
        '
        Me.LblLoadTKT.AutoSize = True
        Me.LblLoadTKT.Location = New System.Drawing.Point(405, 232)
        Me.LblLoadTKT.Name = "LblLoadTKT"
        Me.LblLoadTKT.Size = New System.Drawing.Size(60, 13)
        Me.LblLoadTKT.TabIndex = 16
        Me.LblLoadTKT.TabStop = True
        Me.LblLoadTKT.Text = "Load TKTs"
        Me.LblLoadTKT.Visible = False
        '
        'LblCheckAll_unCheckAll_AL
        '
        Me.LblCheckAll_unCheckAll_AL.AutoSize = True
        Me.LblCheckAll_unCheckAll_AL.Location = New System.Drawing.Point(3, 232)
        Me.LblCheckAll_unCheckAll_AL.Name = "LblCheckAll_unCheckAll_AL"
        Me.LblCheckAll_unCheckAll_AL.Size = New System.Drawing.Size(63, 13)
        Me.LblCheckAll_unCheckAll_AL.TabIndex = 15
        Me.LblCheckAll_unCheckAll_AL.TabStop = True
        Me.LblCheckAll_unCheckAll_AL.Text = "UnCheckAll"
        '
        'TxtDOIto
        '
        Me.TxtDOIto.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtDOIto.Location = New System.Drawing.Point(135, 3)
        Me.TxtDOIto.Name = "TxtDOIto"
        Me.TxtDOIto.Size = New System.Drawing.Size(101, 20)
        Me.TxtDOIto.TabIndex = 13
        '
        'txtDOIFrm
        '
        Me.txtDOIFrm.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtDOIFrm.Location = New System.Drawing.Point(32, 3)
        Me.txtDOIFrm.Name = "txtDOIFrm"
        Me.txtDOIFrm.Size = New System.Drawing.Size(101, 20)
        Me.txtDOIFrm.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(26, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "DOI"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(316, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "T.Code"
        '
        'txtTCocde
        '
        Me.txtTCocde.Location = New System.Drawing.Point(364, 27)
        Me.txtTCocde.Name = "txtTCocde"
        Me.txtTCocde.Size = New System.Drawing.Size(101, 20)
        Me.txtTCocde.TabIndex = 11
        '
        'CmbCounter
        '
        Me.CmbCounter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCounter.FormattingEnabled = True
        Me.CmbCounter.Items.AddRange(New Object() {"", "CTS", "TVS"})
        Me.CmbCounter.Location = New System.Drawing.Point(260, 3)
        Me.CmbCounter.Name = "CmbCounter"
        Me.CmbCounter.Size = New System.Drawing.Size(53, 21)
        Me.CmbCounter.TabIndex = 10
        '
        'LstAL
        '
        Me.LstAL.CheckOnClick = True
        Me.LstAL.ColumnWidth = 36
        Me.LstAL.FormattingEnabled = True
        Me.LstAL.Location = New System.Drawing.Point(6, 30)
        Me.LstAL.MultiColumn = True
        Me.LstAL.Name = "LstAL"
        Me.LstAL.Size = New System.Drawing.Size(309, 199)
        Me.LstAL.TabIndex = 9
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.LblCalcRASInvAmt)
        Me.TabPage2.Controls.Add(Me.LBlChk_UnChkAll_TKT)
        Me.TabPage2.Controls.Add(Me.txtSummary)
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Controls.Add(Me.GridTKT)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(777, 475)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "TKT_List"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'LblCalcRASInvAmt
        '
        Me.LblCalcRASInvAmt.AutoSize = True
        Me.LblCalcRASInvAmt.Location = New System.Drawing.Point(68, 417)
        Me.LblCalcRASInvAmt.Name = "LblCalcRASInvAmt"
        Me.LblCalcRASInvAmt.Size = New System.Drawing.Size(95, 13)
        Me.LblCalcRASInvAmt.TabIndex = 17
        Me.LblCalcRASInvAmt.TabStop = True
        Me.LblCalcRASInvAmt.Text = "Calc. RAS Inv Amt"
        '
        'LBlChk_UnChkAll_TKT
        '
        Me.LBlChk_UnChkAll_TKT.AutoSize = True
        Me.LBlChk_UnChkAll_TKT.Location = New System.Drawing.Point(3, 417)
        Me.LBlChk_UnChkAll_TKT.Name = "LBlChk_UnChkAll_TKT"
        Me.LBlChk_UnChkAll_TKT.Size = New System.Drawing.Size(49, 13)
        Me.LBlChk_UnChkAll_TKT.TabIndex = 15
        Me.LBlChk_UnChkAll_TKT.TabStop = True
        Me.LBlChk_UnChkAll_TKT.Text = "CheckAll"
        '
        'txtSummary
        '
        Me.txtSummary.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSummary.Location = New System.Drawing.Point(245, 417)
        Me.txtSummary.Name = "txtSummary"
        Me.txtSummary.ReadOnly = True
        Me.txtSummary.Size = New System.Drawing.Size(219, 55)
        Me.txtSummary.TabIndex = 16
        Me.txtSummary.Text = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtInvDate)
        Me.GroupBox1.Controls.Add(Me.LblPreview)
        Me.GroupBox1.Controls.Add(Me.CmbPrintToWhom)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Location = New System.Drawing.Point(468, 417)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(305, 55)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Invoice"
        '
        'TxtInvDate
        '
        Me.TxtInvDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtInvDate.Location = New System.Drawing.Point(139, 20)
        Me.TxtInvDate.Name = "TxtInvDate"
        Me.TxtInvDate.Size = New System.Drawing.Size(89, 20)
        Me.TxtInvDate.TabIndex = 6
        '
        'CmbPrintToWhom
        '
        Me.CmbPrintToWhom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbPrintToWhom.FormattingEnabled = True
        Me.CmbPrintToWhom.Items.AddRange(New Object() {"TransViet", "Customer"})
        Me.CmbPrintToWhom.Location = New System.Drawing.Point(20, 19)
        Me.CmbPrintToWhom.Name = "CmbPrintToWhom"
        Me.CmbPrintToWhom.Size = New System.Drawing.Size(89, 21)
        Me.CmbPrintToWhom.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "To"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(110, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Date"
        '
        'BulkPrint_New
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BulkPrint_New"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Bulk INV Printing"
        CType(Me.GridTKT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GridTKT As System.Windows.Forms.DataGridView
    Friend WithEvents LblPreview As System.Windows.Forms.LinkLabel
    Friend WithEvents S As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents cmbCust As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents CmbCounter As System.Windows.Forms.ComboBox
    Friend WithEvents LstAL As System.Windows.Forms.CheckedListBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTCocde As System.Windows.Forms.TextBox
    Friend WithEvents CmbPrintToWhom As System.Windows.Forms.ComboBox
    Friend WithEvents TxtDOIto As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtDOIFrm As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblCheckAll_unCheckAll_AL As System.Windows.Forms.LinkLabel
    Friend WithEvents LblLoadTKT As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtInvDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtSummary As System.Windows.Forms.RichTextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LBlChk_UnChkAll_TKT As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCalcRASInvAmt As System.Windows.Forms.LinkLabel
End Class
