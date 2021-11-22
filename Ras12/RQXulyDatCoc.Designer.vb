<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RQXulyDatCoc
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
        Me.TxtDocNo = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblFind = New System.Windows.Forms.LinkLabel()
        Me.LblSendRQ = New System.Windows.Forms.LinkLabel()
        Me.CmbRQType = New System.Windows.Forms.ComboBox()
        Me.GridPending = New System.Windows.Forms.DataGridView()
        Me.GridPreview = New System.Windows.Forms.DataGridView()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TxtNote = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TxtFee = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.LblNewFltDate = New System.Windows.Forms.Label()
        Me.TxtNewTL = New System.Windows.Forms.DateTimePicker()
        Me.LblProcess = New System.Windows.Forms.LinkLabel()
        Me.LblRefund = New System.Windows.Forms.LinkLabel()
        Me.chckOKonly = New System.Windows.Forms.CheckBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lbkReject = New System.Windows.Forms.LinkLabel()
        CType(Me.GridPending, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridPreview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtDocNo
        '
        Me.TxtDocNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDocNo.Location = New System.Drawing.Point(83, 15)
        Me.TxtDocNo.Name = "TxtDocNo"
        Me.TxtDocNo.Size = New System.Drawing.Size(151, 20)
        Me.TxtDocNo.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 26)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "PseudoMCO" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(start with GRP..)"
        '
        'LblFind
        '
        Me.LblFind.AutoSize = True
        Me.LblFind.Location = New System.Drawing.Point(266, 9)
        Me.LblFind.Name = "LblFind"
        Me.LblFind.Size = New System.Drawing.Size(27, 13)
        Me.LblFind.TabIndex = 2
        Me.LblFind.TabStop = True
        Me.LblFind.Text = "Find"
        '
        'LblSendRQ
        '
        Me.LblSendRQ.AutoSize = True
        Me.LblSendRQ.Location = New System.Drawing.Point(563, 33)
        Me.LblSendRQ.Name = "LblSendRQ"
        Me.LblSendRQ.Size = New System.Drawing.Size(51, 13)
        Me.LblSendRQ.TabIndex = 2
        Me.LblSendRQ.TabStop = True
        Me.LblSendRQ.Text = "Send RQ"
        '
        'CmbRQType
        '
        Me.CmbRQType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbRQType.FormattingEnabled = True
        Me.CmbRQType.Location = New System.Drawing.Point(355, 6)
        Me.CmbRQType.Name = "CmbRQType"
        Me.CmbRQType.Size = New System.Drawing.Size(84, 21)
        Me.CmbRQType.TabIndex = 3
        '
        'GridPending
        '
        Me.GridPending.AllowUserToAddRows = False
        Me.GridPending.AllowUserToDeleteRows = False
        Me.GridPending.BackgroundColor = System.Drawing.SystemColors.ControlLight
        Me.GridPending.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPending.GridColor = System.Drawing.SystemColors.ControlLight
        Me.GridPending.Location = New System.Drawing.Point(3, 148)
        Me.GridPending.Name = "GridPending"
        Me.GridPending.RowHeadersVisible = False
        Me.GridPending.Size = New System.Drawing.Size(764, 297)
        Me.GridPending.TabIndex = 4
        '
        'GridPreview
        '
        Me.GridPreview.AllowUserToAddRows = False
        Me.GridPreview.AllowUserToDeleteRows = False
        Me.GridPreview.BackgroundColor = System.Drawing.SystemColors.ControlLightLight
        Me.GridPreview.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPreview.GridColor = System.Drawing.SystemColors.ControlLight
        Me.GridPreview.Location = New System.Drawing.Point(3, 61)
        Me.GridPreview.Name = "GridPreview"
        Me.GridPreview.RowHeadersVisible = False
        Me.GridPreview.Size = New System.Drawing.Size(764, 68)
        Me.GridPreview.TabIndex = 4
        '
        'lbl2
        '
        Me.lbl2.AutoSize = True
        Me.lbl2.Location = New System.Drawing.Point(299, 9)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(50, 13)
        Me.lbl2.TabIndex = 1
        Me.lbl2.Text = "RQ Type"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(0, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "PreView"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1, 132)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Pending RQ"
        '
        'TxtNote
        '
        Me.TxtNote.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNote.Location = New System.Drawing.Point(473, 6)
        Me.TxtNote.Name = "TxtNote"
        Me.TxtNote.Size = New System.Drawing.Size(297, 20)
        Me.TxtNote.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(442, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Note"
        '
        'TxtFee
        '
        Me.TxtFee.Location = New System.Drawing.Point(473, 30)
        Me.TxtFee.Name = "TxtFee"
        Me.TxtFee.Size = New System.Drawing.Size(84, 20)
        Me.TxtFee.TabIndex = 5
        Me.TxtFee.Text = "0"
        Me.TxtFee.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(442, 33)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(25, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Fee"
        '
        'LblNewFltDate
        '
        Me.LblNewFltDate.AutoSize = True
        Me.LblNewFltDate.Location = New System.Drawing.Point(299, 33)
        Me.LblNewFltDate.Name = "LblNewFltDate"
        Me.LblNewFltDate.Size = New System.Drawing.Size(45, 13)
        Me.LblNewFltDate.TabIndex = 6
        Me.LblNewFltDate.Text = "New TL"
        '
        'TxtNewTL
        '
        Me.TxtNewTL.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtNewTL.Location = New System.Drawing.Point(355, 30)
        Me.TxtNewTL.Name = "TxtNewTL"
        Me.TxtNewTL.Size = New System.Drawing.Size(84, 20)
        Me.TxtNewTL.TabIndex = 7
        '
        'LblProcess
        '
        Me.LblProcess.AutoSize = True
        Me.LblProcess.Location = New System.Drawing.Point(1, 448)
        Me.LblProcess.Name = "LblProcess"
        Me.LblProcess.Size = New System.Drawing.Size(45, 13)
        Me.LblProcess.TabIndex = 8
        Me.LblProcess.TabStop = True
        Me.LblProcess.Text = "Process"
        Me.LblProcess.Visible = False
        '
        'LblRefund
        '
        Me.LblRefund.AutoSize = True
        Me.LblRefund.Location = New System.Drawing.Point(143, 448)
        Me.LblRefund.Name = "LblRefund"
        Me.LblRefund.Size = New System.Drawing.Size(42, 13)
        Me.LblRefund.TabIndex = 8
        Me.LblRefund.TabStop = True
        Me.LblRefund.Text = "Refund"
        Me.LblRefund.Visible = False
        '
        'chckOKonly
        '
        Me.chckOKonly.AutoSize = True
        Me.chckOKonly.Location = New System.Drawing.Point(688, 447)
        Me.chckOKonly.Name = "chckOKonly"
        Me.chckOKonly.Size = New System.Drawing.Size(82, 17)
        Me.chckOKonly.TabIndex = 9
        Me.chckOKonly.Text = "OK RQ only"
        Me.chckOKonly.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(4, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(778, 496)
        Me.TabControl1.TabIndex = 10
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lbkReject)
        Me.TabPage1.Controls.Add(Me.TxtDocNo)
        Me.TabPage1.Controls.Add(Me.chckOKonly)
        Me.TabPage1.Controls.Add(Me.TxtNote)
        Me.TabPage1.Controls.Add(Me.LblRefund)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.LblProcess)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.TxtNewTL)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.LblNewFltDate)
        Me.TabPage1.Controls.Add(Me.LblFind)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.lbl2)
        Me.TabPage1.Controls.Add(Me.TxtFee)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.GridPreview)
        Me.TabPage1.Controls.Add(Me.LblSendRQ)
        Me.TabPage1.Controls.Add(Me.GridPending)
        Me.TabPage1.Controls.Add(Me.CmbRQType)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(770, 470)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Single RQ"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lbkReject
        '
        Me.lbkReject.AutoSize = True
        Me.lbkReject.Location = New System.Drawing.Point(266, 448)
        Me.lbkReject.Name = "lbkReject"
        Me.lbkReject.Size = New System.Drawing.Size(38, 13)
        Me.lbkReject.TabIndex = 10
        Me.lbkReject.TabStop = True
        Me.lbkReject.Text = "Reject"
        Me.lbkReject.Visible = False
        '
        'RQXulyDatCoc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "RQXulyDatCoc"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Request for Deposit Handling"
        CType(Me.GridPending, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridPreview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TxtDocNo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblFind As System.Windows.Forms.LinkLabel
    Friend WithEvents LblSendRQ As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbRQType As System.Windows.Forms.ComboBox
    Friend WithEvents GridPending As System.Windows.Forms.DataGridView
    Friend WithEvents GridPreview As System.Windows.Forms.DataGridView
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtNote As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtFee As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents LblNewFltDate As System.Windows.Forms.Label
    Friend WithEvents TxtNewTL As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblProcess As System.Windows.Forms.LinkLabel
    Friend WithEvents LblRefund As System.Windows.Forms.LinkLabel
    Friend WithEvents chckOKonly As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents lbkReject As System.Windows.Forms.LinkLabel
End Class
