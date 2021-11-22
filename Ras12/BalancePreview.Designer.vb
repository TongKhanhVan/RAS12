<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BalancePreview
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
        Me.OptPPD = New System.Windows.Forms.RadioButton()
        Me.OptPSP = New System.Windows.Forms.RadioButton()
        Me.GridBalance = New System.Windows.Forms.DataGridView()
        Me.TxtEmail = New System.Windows.Forms.RichTextBox()
        Me.LstOffice = New System.Windows.Forms.ListBox()
        Me.LstTKT = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TxtOfficeID = New System.Windows.Forms.TextBox()
        Me.LblSearch = New System.Windows.Forms.LinkLabel()
        Me.LblRefresh = New System.Windows.Forms.LinkLabel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.OptALL = New System.Windows.Forms.RadioButton()
        Me.OptLC = New System.Windows.Forms.RadioButton()
        Me.OptFT = New System.Windows.Forms.RadioButton()
        Me.Opt1S = New System.Windows.Forms.RadioButton()
        Me.OptCWT = New System.Windows.Forms.RadioButton()
        Me.LblChickBalance = New System.Windows.Forms.Label()
        Me.GridDetail = New System.Windows.Forms.DataGridView()
        Me.PnlDetail = New System.Windows.Forms.Panel()
        Me.TxtTo = New System.Windows.Forms.DateTimePicker()
        Me.txtFrm = New System.Windows.Forms.DateTimePicker()
        Me.LblCloseDetail = New System.Windows.Forms.LinkLabel()
        Me.OptM1S = New System.Windows.Forms.RadioButton()
        Me.OptRAS = New System.Windows.Forms.RadioButton()
        CType(Me.GridBalance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.GridDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'OptPPD
        '
        Me.OptPPD.AutoSize = True
        Me.OptPPD.Location = New System.Drawing.Point(6, 9)
        Me.OptPPD.Name = "OptPPD"
        Me.OptPPD.Size = New System.Drawing.Size(47, 17)
        Me.OptPPD.TabIndex = 0
        Me.OptPPD.Text = "PPD"
        Me.OptPPD.UseVisualStyleBackColor = True
        '
        'OptPSP
        '
        Me.OptPSP.AutoSize = True
        Me.OptPSP.Checked = True
        Me.OptPSP.Location = New System.Drawing.Point(59, 9)
        Me.OptPSP.Name = "OptPSP"
        Me.OptPSP.Size = New System.Drawing.Size(46, 17)
        Me.OptPSP.TabIndex = 0
        Me.OptPSP.TabStop = True
        Me.OptPSP.Text = "PSP"
        Me.OptPSP.UseVisualStyleBackColor = True
        '
        'GridBalance
        '
        Me.GridBalance.AllowUserToAddRows = False
        Me.GridBalance.AllowUserToDeleteRows = False
        Me.GridBalance.BackgroundColor = System.Drawing.Color.FloralWhite
        Me.GridBalance.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridBalance.Location = New System.Drawing.Point(4, 41)
        Me.GridBalance.Name = "GridBalance"
        Me.GridBalance.ReadOnly = True
        Me.GridBalance.RowHeadersVisible = False
        Me.GridBalance.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridBalance.Size = New System.Drawing.Size(404, 458)
        Me.GridBalance.TabIndex = 1
        '
        'TxtEmail
        '
        Me.TxtEmail.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtEmail.Location = New System.Drawing.Point(409, 41)
        Me.TxtEmail.Name = "TxtEmail"
        Me.TxtEmail.Size = New System.Drawing.Size(420, 79)
        Me.TxtEmail.TabIndex = 2
        Me.TxtEmail.Text = ""
        '
        'LstOffice
        '
        Me.LstOffice.ColumnWidth = 75
        Me.LstOffice.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstOffice.FormattingEnabled = True
        Me.LstOffice.ItemHeight = 14
        Me.LstOffice.Location = New System.Drawing.Point(410, 139)
        Me.LstOffice.MultiColumn = True
        Me.LstOffice.Name = "LstOffice"
        Me.LstOffice.Size = New System.Drawing.Size(420, 60)
        Me.LstOffice.TabIndex = 3
        '
        'LstTKT
        '
        Me.LstTKT.ColumnWidth = 205
        Me.LstTKT.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstTKT.FormattingEnabled = True
        Me.LstTKT.ItemHeight = 14
        Me.LstTKT.Location = New System.Drawing.Point(409, 219)
        Me.LstTKT.MultiColumn = True
        Me.LstTKT.Name = "LstTKT"
        Me.LstTKT.Size = New System.Drawing.Size(420, 88)
        Me.LstTKT.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(408, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Emails"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(408, 203)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(119, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "TKT Beeing Held Credit"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(408, 123)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "OfficeIDs"
        '
        'TxtOfficeID
        '
        Me.TxtOfficeID.Location = New System.Drawing.Point(411, 479)
        Me.TxtOfficeID.Name = "TxtOfficeID"
        Me.TxtOfficeID.Size = New System.Drawing.Size(88, 20)
        Me.TxtOfficeID.TabIndex = 6
        Me.TxtOfficeID.Text = "SGNVM"
        '
        'LblSearch
        '
        Me.LblSearch.AutoSize = True
        Me.LblSearch.Location = New System.Drawing.Point(503, 482)
        Me.LblSearch.Name = "LblSearch"
        Me.LblSearch.Size = New System.Drawing.Size(78, 13)
        Me.LblSearch.TabIndex = 7
        Me.LblSearch.TabStop = True
        Me.LblSearch.Text = "Search By OID"
        '
        'LblRefresh
        '
        Me.LblRefresh.AutoSize = True
        Me.LblRefresh.Location = New System.Drawing.Point(746, 482)
        Me.LblRefresh.Name = "LblRefresh"
        Me.LblRefresh.Size = New System.Drawing.Size(83, 13)
        Me.LblRefresh.TabIndex = 8
        Me.LblRefresh.TabStop = True
        Me.LblRefresh.Text = "RefreshBalance"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.OptPPD)
        Me.GroupBox1.Controls.Add(Me.OptPSP)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(107, 32)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.OptALL)
        Me.GroupBox2.Controls.Add(Me.OptLC)
        Me.GroupBox2.Controls.Add(Me.OptFT)
        Me.GroupBox2.Controls.Add(Me.Opt1S)
        Me.GroupBox2.Controls.Add(Me.OptCWT)
        Me.GroupBox2.Location = New System.Drawing.Point(117, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(291, 32)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        '
        'OptALL
        '
        Me.OptALL.AutoSize = True
        Me.OptALL.Location = New System.Drawing.Point(241, 9)
        Me.OptALL.Name = "OptALL"
        Me.OptALL.Size = New System.Drawing.Size(44, 17)
        Me.OptALL.TabIndex = 0
        Me.OptALL.Text = "ALL"
        Me.OptALL.UseVisualStyleBackColor = True
        '
        'OptLC
        '
        Me.OptLC.AutoSize = True
        Me.OptLC.Location = New System.Drawing.Point(170, 9)
        Me.OptLC.Name = "OptLC"
        Me.OptLC.Size = New System.Drawing.Size(44, 17)
        Me.OptLC.TabIndex = 0
        Me.OptLC.Text = "LCL"
        Me.OptLC.UseVisualStyleBackColor = True
        '
        'OptFT
        '
        Me.OptFT.AutoSize = True
        Me.OptFT.Location = New System.Drawing.Point(126, 9)
        Me.OptFT.Name = "OptFT"
        Me.OptFT.Size = New System.Drawing.Size(38, 17)
        Me.OptFT.TabIndex = 0
        Me.OptFT.Text = "FT"
        Me.OptFT.UseVisualStyleBackColor = True
        '
        'Opt1S
        '
        Me.Opt1S.AutoSize = True
        Me.Opt1S.Checked = True
        Me.Opt1S.Location = New System.Drawing.Point(62, 9)
        Me.Opt1S.Name = "Opt1S"
        Me.Opt1S.Size = New System.Drawing.Size(58, 17)
        Me.Opt1S.TabIndex = 0
        Me.Opt1S.TabStop = True
        Me.Opt1S.Text = "VN/1S"
        Me.Opt1S.UseVisualStyleBackColor = True
        '
        'OptCWT
        '
        Me.OptCWT.AutoSize = True
        Me.OptCWT.Location = New System.Drawing.Point(6, 9)
        Me.OptCWT.Name = "OptCWT"
        Me.OptCWT.Size = New System.Drawing.Size(50, 17)
        Me.OptCWT.TabIndex = 0
        Me.OptCWT.Text = "CWT"
        Me.OptCWT.UseVisualStyleBackColor = True
        '
        'LblChickBalance
        '
        Me.LblChickBalance.AutoSize = True
        Me.LblChickBalance.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblChickBalance.ForeColor = System.Drawing.Color.Blue
        Me.LblChickBalance.Location = New System.Drawing.Point(408, 310)
        Me.LblChickBalance.Name = "LblChickBalance"
        Me.LblChickBalance.Size = New System.Drawing.Size(51, 15)
        Me.LblChickBalance.TabIndex = 11
        Me.LblChickBalance.Text = "Label4"
        Me.LblChickBalance.Visible = False
        '
        'GridDetail
        '
        Me.GridDetail.AllowUserToAddRows = False
        Me.GridDetail.AllowUserToDeleteRows = False
        Me.GridDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridDetail.BackgroundColor = System.Drawing.Color.AliceBlue
        Me.GridDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridDetail.Location = New System.Drawing.Point(3, 3)
        Me.GridDetail.Name = "GridDetail"
        Me.GridDetail.RowHeadersVisible = False
        Me.GridDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridDetail.Size = New System.Drawing.Size(408, 45)
        Me.GridDetail.TabIndex = 12
        '
        'PnlDetail
        '
        Me.PnlDetail.Controls.Add(Me.TxtTo)
        Me.PnlDetail.Controls.Add(Me.txtFrm)
        Me.PnlDetail.Controls.Add(Me.LblCloseDetail)
        Me.PnlDetail.Controls.Add(Me.OptM1S)
        Me.PnlDetail.Controls.Add(Me.OptRAS)
        Me.PnlDetail.Controls.Add(Me.GridDetail)
        Me.PnlDetail.Location = New System.Drawing.Point(414, 376)
        Me.PnlDetail.Name = "PnlDetail"
        Me.PnlDetail.Size = New System.Drawing.Size(414, 72)
        Me.PnlDetail.TabIndex = 13
        '
        'TxtTo
        '
        Me.TxtTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtTo.Location = New System.Drawing.Point(295, 50)
        Me.TxtTo.Name = "TxtTo"
        Me.TxtTo.Size = New System.Drawing.Size(73, 20)
        Me.TxtTo.TabIndex = 15
        '
        'txtFrm
        '
        Me.txtFrm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtFrm.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtFrm.Location = New System.Drawing.Point(220, 50)
        Me.txtFrm.Name = "txtFrm"
        Me.txtFrm.Size = New System.Drawing.Size(73, 20)
        Me.txtFrm.TabIndex = 15
        '
        'LblCloseDetail
        '
        Me.LblCloseDetail.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblCloseDetail.AutoSize = True
        Me.LblCloseDetail.Location = New System.Drawing.Point(378, 52)
        Me.LblCloseDetail.Name = "LblCloseDetail"
        Me.LblCloseDetail.Size = New System.Drawing.Size(33, 13)
        Me.LblCloseDetail.TabIndex = 14
        Me.LblCloseDetail.TabStop = True
        Me.LblCloseDetail.Text = "Close"
        '
        'OptM1S
        '
        Me.OptM1S.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.OptM1S.AutoSize = True
        Me.OptM1S.Location = New System.Drawing.Point(88, 50)
        Me.OptM1S.Name = "OptM1S"
        Me.OptM1S.Size = New System.Drawing.Size(127, 17)
        Me.OptM1S.TabIndex = 13
        Me.OptM1S.Text = "SabreMonitorBalance"
        Me.OptM1S.UseVisualStyleBackColor = True
        '
        'OptRAS
        '
        Me.OptRAS.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.OptRAS.AutoSize = True
        Me.OptRAS.Checked = True
        Me.OptRAS.Location = New System.Drawing.Point(3, 50)
        Me.OptRAS.Name = "OptRAS"
        Me.OptRAS.Size = New System.Drawing.Size(83, 17)
        Me.OptRAS.TabIndex = 13
        Me.OptRAS.TabStop = True
        Me.OptRAS.Text = "RasBalance"
        Me.OptRAS.UseVisualStyleBackColor = True
        '
        'BalancePreview
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(830, 501)
        Me.Controls.Add(Me.PnlDetail)
        Me.Controls.Add(Me.LblChickBalance)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.LblRefresh)
        Me.Controls.Add(Me.LblSearch)
        Me.Controls.Add(Me.TxtOfficeID)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LstTKT)
        Me.Controls.Add(Me.LstOffice)
        Me.Controls.Add(Me.TxtEmail)
        Me.Controls.Add(Me.GridBalance)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BalancePreview"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Balance Preview"
        CType(Me.GridBalance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.GridDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlDetail.ResumeLayout(False)
        Me.PnlDetail.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OptPPD As System.Windows.Forms.RadioButton
    Friend WithEvents OptPSP As System.Windows.Forms.RadioButton
    Friend WithEvents GridBalance As System.Windows.Forms.DataGridView
    Friend WithEvents TxtEmail As System.Windows.Forms.RichTextBox
    Friend WithEvents LstOffice As System.Windows.Forms.ListBox
    Friend WithEvents LstTKT As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtOfficeID As System.Windows.Forms.TextBox
    Friend WithEvents LblSearch As System.Windows.Forms.LinkLabel
    Friend WithEvents LblRefresh As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents OptALL As System.Windows.Forms.RadioButton
    Friend WithEvents OptFT As System.Windows.Forms.RadioButton
    Friend WithEvents Opt1S As System.Windows.Forms.RadioButton
    Friend WithEvents OptCWT As System.Windows.Forms.RadioButton
    Friend WithEvents LblChickBalance As System.Windows.Forms.Label
    Friend WithEvents OptLC As System.Windows.Forms.RadioButton
    Friend WithEvents GridDetail As System.Windows.Forms.DataGridView
    Friend WithEvents PnlDetail As System.Windows.Forms.Panel
    Friend WithEvents LblCloseDetail As System.Windows.Forms.LinkLabel
    Friend WithEvents OptM1S As System.Windows.Forms.RadioButton
    Friend WithEvents OptRAS As System.Windows.Forms.RadioButton
    Friend WithEvents txtFrm As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtTo As System.Windows.Forms.DateTimePicker
End Class
