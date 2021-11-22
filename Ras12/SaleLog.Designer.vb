<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SaleLog
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SaleLog))
        Me.GridCust = New System.Windows.Forms.DataGridView()
        Me.TxtPre = New System.Windows.Forms.RichTextBox()
        Me.TxtPost = New System.Windows.Forms.RichTextBox()
        Me.txtStart = New System.Windows.Forms.DateTimePicker()
        Me.txtEnd = New System.Windows.Forms.DateTimePicker()
        Me.LblMakeAppt = New System.Windows.Forms.LinkLabel()
        Me.LblDone = New System.Windows.Forms.LinkLabel()
        Me.TxtSubj = New System.Windows.Forms.TextBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.CmbCustType = New System.Windows.Forms.ComboBox()
        Me.ChkByPhone = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.OptLead = New System.Windows.Forms.RadioButton()
        Me.OptCust = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.LblChange = New System.Windows.Forms.LinkLabel()
        Me.TxtNewEnd = New System.Windows.Forms.DateTimePicker()
        Me.txtNewStart = New System.Windows.Forms.DateTimePicker()
        Me.LblCancel = New System.Windows.Forms.LinkLabel()
        Me.LblUpdatePrepar = New System.Windows.Forms.LinkLabel()
        Me.TxtPreView = New System.Windows.Forms.RichTextBox()
        Me.GridPendingCall = New System.Windows.Forms.DataGridView()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.LckLblDelete = New System.Windows.Forms.LinkLabel()
        Me.LblRefresh = New System.Windows.Forms.LinkLabel()
        Me.ChkSelectCust = New System.Windows.Forms.CheckBox()
        Me.txtRptTimeThru = New System.Windows.Forms.DateTimePicker()
        Me.txtRptTimeFrm = New System.Windows.Forms.DateTimePicker()
        Me.LblSum = New System.Windows.Forms.Label()
        Me.txtPostView = New System.Windows.Forms.RichTextBox()
        Me.GridDoneCall = New System.Windows.Forms.DataGridView()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.TB_Color = New System.Windows.Forms.ToolStripButton()
        Me.TB_Bold = New System.Windows.Forms.ToolStripButton()
        Me.TB_Underline = New System.Windows.Forms.ToolStripButton()
        Me.TB_Bullet = New System.Windows.Forms.ToolStripButton()
        CType(Me.GridCust, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.GridPendingCall, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.GridDoneCall, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridCust
        '
        Me.GridCust.AllowUserToAddRows = False
        Me.GridCust.AllowUserToDeleteRows = False
        Me.GridCust.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridCust.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCust.Location = New System.Drawing.Point(0, 0)
        Me.GridCust.Name = "GridCust"
        Me.GridCust.RowHeadersVisible = False
        Me.GridCust.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridCust.Size = New System.Drawing.Size(236, 332)
        Me.GridCust.TabIndex = 0
        '
        'TxtPre
        '
        Me.TxtPre.AcceptsTab = True
        Me.TxtPre.BulletIndent = 4
        Me.TxtPre.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtPre.Location = New System.Drawing.Point(238, 0)
        Me.TxtPre.Name = "TxtPre"
        Me.TxtPre.Size = New System.Drawing.Size(530, 461)
        Me.TxtPre.TabIndex = 1
        Me.TxtPre.Text = ""
        '
        'TxtPost
        '
        Me.TxtPost.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtPost.Location = New System.Drawing.Point(3, 263)
        Me.TxtPost.Name = "TxtPost"
        Me.TxtPost.Size = New System.Drawing.Size(769, 183)
        Me.TxtPost.TabIndex = 1
        Me.TxtPost.Text = ""
        '
        'txtStart
        '
        Me.txtStart.CustomFormat = "dd-MMM-yy HH:mm"
        Me.txtStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtStart.Location = New System.Drawing.Point(38, 373)
        Me.txtStart.Name = "txtStart"
        Me.txtStart.Size = New System.Drawing.Size(117, 20)
        Me.txtStart.TabIndex = 3
        '
        'txtEnd
        '
        Me.txtEnd.CustomFormat = "dd-MMM-yy HH:mm"
        Me.txtEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtEnd.Location = New System.Drawing.Point(38, 395)
        Me.txtEnd.Name = "txtEnd"
        Me.txtEnd.Size = New System.Drawing.Size(117, 20)
        Me.txtEnd.TabIndex = 3
        '
        'LblMakeAppt
        '
        Me.LblMakeAppt.AutoSize = True
        Me.LblMakeAppt.Location = New System.Drawing.Point(167, 401)
        Me.LblMakeAppt.Name = "LblMakeAppt"
        Me.LblMakeAppt.Size = New System.Drawing.Size(56, 13)
        Me.LblMakeAppt.TabIndex = 4
        Me.LblMakeAppt.TabStop = True
        Me.LblMakeAppt.Text = "MakeAppt"
        '
        'LblDone
        '
        Me.LblDone.AutoSize = True
        Me.LblDone.Location = New System.Drawing.Point(732, 449)
        Me.LblDone.Name = "LblDone"
        Me.LblDone.Size = New System.Drawing.Size(33, 13)
        Me.LblDone.TabIndex = 4
        Me.LblDone.TabStop = True
        Me.LblDone.Text = "Done"
        '
        'TxtSubj
        '
        Me.TxtSubj.Enabled = False
        Me.TxtSubj.Location = New System.Drawing.Point(2, 349)
        Me.TxtSubj.Name = "TxtSubj"
        Me.TxtSubj.Size = New System.Drawing.Size(232, 20)
        Me.TxtSubj.TabIndex = 5
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(0, 25)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(780, 490)
        Me.TabControl1.TabIndex = 6
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TxtPre)
        Me.TabPage1.Controls.Add(Me.CmbCustType)
        Me.TabPage1.Controls.Add(Me.ChkByPhone)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.GridCust)
        Me.TabPage1.Controls.Add(Me.LblMakeAppt)
        Me.TabPage1.Controls.Add(Me.TxtSubj)
        Me.TabPage1.Controls.Add(Me.txtStart)
        Me.TabPage1.Controls.Add(Me.txtEnd)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(772, 464)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Pre Call"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'CmbCustType
        '
        Me.CmbCustType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCustType.FormattingEnabled = True
        Me.CmbCustType.Location = New System.Drawing.Point(38, 435)
        Me.CmbCustType.Name = "CmbCustType"
        Me.CmbCustType.Size = New System.Drawing.Size(62, 21)
        Me.CmbCustType.TabIndex = 7
        '
        'ChkByPhone
        '
        Me.ChkByPhone.AutoSize = True
        Me.ChkByPhone.Location = New System.Drawing.Point(155, 375)
        Me.ChkByPhone.Name = "ChkByPhone"
        Me.ChkByPhone.Size = New System.Drawing.Size(80, 17)
        Me.ChkByPhone.TabIndex = 8
        Me.ChkByPhone.Text = "OverPhone"
        Me.ChkByPhone.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.OptLead)
        Me.GroupBox2.Controls.Add(Me.OptCust)
        Me.GroupBox2.Location = New System.Drawing.Point(106, 427)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(129, 32)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        '
        'OptLead
        '
        Me.OptLead.AutoSize = True
        Me.OptLead.Location = New System.Drawing.Point(61, 9)
        Me.OptLead.Name = "OptLead"
        Me.OptLead.Size = New System.Drawing.Size(49, 17)
        Me.OptLead.TabIndex = 8
        Me.OptLead.Text = "Lead"
        Me.OptLead.UseVisualStyleBackColor = True
        '
        'OptCust
        '
        Me.OptCust.AutoSize = True
        Me.OptCust.Checked = True
        Me.OptCust.Location = New System.Drawing.Point(6, 9)
        Me.OptCust.Name = "OptCust"
        Me.OptCust.Size = New System.Drawing.Size(49, 17)
        Me.OptCust.TabIndex = 8
        Me.OptCust.TabStop = True
        Me.OptCust.Text = "Cust."
        Me.OptCust.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 377)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Time"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 335)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Purpose"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.TxtPost)
        Me.TabPage2.Controls.Add(Me.LblChange)
        Me.TabPage2.Controls.Add(Me.TxtNewEnd)
        Me.TabPage2.Controls.Add(Me.txtNewStart)
        Me.TabPage2.Controls.Add(Me.LblCancel)
        Me.TabPage2.Controls.Add(Me.LblUpdatePrepar)
        Me.TabPage2.Controls.Add(Me.TxtPreView)
        Me.TabPage2.Controls.Add(Me.GridPendingCall)
        Me.TabPage2.Controls.Add(Me.LblDone)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(772, 464)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "After Call"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'LblChange
        '
        Me.LblChange.AutoSize = True
        Me.LblChange.Location = New System.Drawing.Point(586, 239)
        Me.LblChange.Name = "LblChange"
        Me.LblChange.Size = New System.Drawing.Size(69, 13)
        Me.LblChange.TabIndex = 10
        Me.LblChange.TabStop = True
        Me.LblChange.Text = "ChangeAppt."
        '
        'TxtNewEnd
        '
        Me.TxtNewEnd.CustomFormat = "dd-MMM-yy HH:mm"
        Me.TxtNewEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TxtNewEnd.Location = New System.Drawing.Point(457, 237)
        Me.TxtNewEnd.Name = "TxtNewEnd"
        Me.TxtNewEnd.Size = New System.Drawing.Size(123, 20)
        Me.TxtNewEnd.TabIndex = 9
        '
        'txtNewStart
        '
        Me.txtNewStart.CustomFormat = "dd-MMM-yy HH:mm"
        Me.txtNewStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtNewStart.Location = New System.Drawing.Point(328, 237)
        Me.txtNewStart.Name = "txtNewStart"
        Me.txtNewStart.Size = New System.Drawing.Size(123, 20)
        Me.txtNewStart.TabIndex = 9
        '
        'LblCancel
        '
        Me.LblCancel.AutoSize = True
        Me.LblCancel.Location = New System.Drawing.Point(3, 449)
        Me.LblCancel.Name = "LblCancel"
        Me.LblCancel.Size = New System.Drawing.Size(81, 13)
        Me.LblCancel.TabIndex = 8
        Me.LblCancel.TabStop = True
        Me.LblCancel.Text = "Cancel Meeting"
        '
        'LblUpdatePrepar
        '
        Me.LblUpdatePrepar.AutoSize = True
        Me.LblUpdatePrepar.Location = New System.Drawing.Point(723, 239)
        Me.LblUpdatePrepar.Name = "LblUpdatePrepar"
        Me.LblUpdatePrepar.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdatePrepar.TabIndex = 7
        Me.LblUpdatePrepar.TabStop = True
        Me.LblUpdatePrepar.Text = "Update"
        '
        'TxtPreView
        '
        Me.TxtPreView.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtPreView.Location = New System.Drawing.Point(327, 4)
        Me.TxtPreView.Name = "TxtPreView"
        Me.TxtPreView.Size = New System.Drawing.Size(442, 229)
        Me.TxtPreView.TabIndex = 6
        Me.TxtPreView.Text = ""
        '
        'GridPendingCall
        '
        Me.GridPendingCall.AllowUserToAddRows = False
        Me.GridPendingCall.AllowUserToDeleteRows = False
        Me.GridPendingCall.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridPendingCall.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPendingCall.Location = New System.Drawing.Point(3, 3)
        Me.GridPendingCall.Name = "GridPendingCall"
        Me.GridPendingCall.ReadOnly = True
        Me.GridPendingCall.RowHeadersVisible = False
        Me.GridPendingCall.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridPendingCall.Size = New System.Drawing.Size(319, 254)
        Me.GridPendingCall.TabIndex = 5
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.LckLblDelete)
        Me.TabPage3.Controls.Add(Me.LblRefresh)
        Me.TabPage3.Controls.Add(Me.ChkSelectCust)
        Me.TabPage3.Controls.Add(Me.txtRptTimeThru)
        Me.TabPage3.Controls.Add(Me.txtRptTimeFrm)
        Me.TabPage3.Controls.Add(Me.LblSum)
        Me.TabPage3.Controls.Add(Me.txtPostView)
        Me.TabPage3.Controls.Add(Me.GridDoneCall)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(772, 464)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Review"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'LckLblDelete
        '
        Me.LckLblDelete.AutoSize = True
        Me.LckLblDelete.Location = New System.Drawing.Point(322, 442)
        Me.LckLblDelete.Name = "LckLblDelete"
        Me.LckLblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LckLblDelete.TabIndex = 12
        Me.LckLblDelete.TabStop = True
        Me.LckLblDelete.Text = "Delete"
        Me.LckLblDelete.Visible = False
        '
        'LblRefresh
        '
        Me.LblRefresh.AutoSize = True
        Me.LblRefresh.Location = New System.Drawing.Point(366, 442)
        Me.LblRefresh.Name = "LblRefresh"
        Me.LblRefresh.Size = New System.Drawing.Size(44, 13)
        Me.LblRefresh.TabIndex = 11
        Me.LblRefresh.TabStop = True
        Me.LblRefresh.Text = "Refresh"
        '
        'ChkSelectCust
        '
        Me.ChkSelectCust.AutoSize = True
        Me.ChkSelectCust.Location = New System.Drawing.Point(175, 441)
        Me.ChkSelectCust.Name = "ChkSelectCust"
        Me.ChkSelectCust.Size = New System.Drawing.Size(110, 17)
        Me.ChkSelectCust.TabIndex = 10
        Me.ChkSelectCust.Text = "SelectedCustOnly"
        Me.ChkSelectCust.UseVisualStyleBackColor = True
        '
        'txtRptTimeThru
        '
        Me.txtRptTimeThru.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtRptTimeThru.Location = New System.Drawing.Point(89, 439)
        Me.txtRptTimeThru.Name = "txtRptTimeThru"
        Me.txtRptTimeThru.Size = New System.Drawing.Size(82, 20)
        Me.txtRptTimeThru.TabIndex = 9
        '
        'txtRptTimeFrm
        '
        Me.txtRptTimeFrm.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtRptTimeFrm.Location = New System.Drawing.Point(3, 439)
        Me.txtRptTimeFrm.Name = "txtRptTimeFrm"
        Me.txtRptTimeFrm.Size = New System.Drawing.Size(82, 20)
        Me.txtRptTimeFrm.TabIndex = 9
        '
        'LblSum
        '
        Me.LblSum.AutoSize = True
        Me.LblSum.Location = New System.Drawing.Point(3, 9)
        Me.LblSum.Name = "LblSum"
        Me.LblSum.Size = New System.Drawing.Size(39, 13)
        Me.LblSum.TabIndex = 8
        Me.LblSum.Text = "Label3"
        '
        'txtPostView
        '
        Me.txtPostView.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPostView.Location = New System.Drawing.Point(412, 30)
        Me.txtPostView.Name = "txtPostView"
        Me.txtPostView.Size = New System.Drawing.Size(356, 429)
        Me.txtPostView.TabIndex = 7
        Me.txtPostView.Text = ""
        '
        'GridDoneCall
        '
        Me.GridDoneCall.AllowUserToAddRows = False
        Me.GridDoneCall.AllowUserToDeleteRows = False
        Me.GridDoneCall.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridDoneCall.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridDoneCall.Location = New System.Drawing.Point(3, 30)
        Me.GridDoneCall.Name = "GridDoneCall"
        Me.GridDoneCall.RowHeadersVisible = False
        Me.GridDoneCall.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridDoneCall.Size = New System.Drawing.Size(407, 403)
        Me.GridDoneCall.TabIndex = 1
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TB_Color, Me.TB_Bold, Me.TB_Underline, Me.TB_Bullet})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(781, 25)
        Me.ToolStrip1.TabIndex = 9
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'TB_Color
        '
        Me.TB_Color.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TB_Color.Image = CType(resources.GetObject("TB_Color.Image"), System.Drawing.Image)
        Me.TB_Color.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TB_Color.Name = "TB_Color"
        Me.TB_Color.Size = New System.Drawing.Size(23, 22)
        Me.TB_Color.Text = "Font Color"
        '
        'TB_Bold
        '
        Me.TB_Bold.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TB_Bold.Image = CType(resources.GetObject("TB_Bold.Image"), System.Drawing.Image)
        Me.TB_Bold.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TB_Bold.Name = "TB_Bold"
        Me.TB_Bold.Size = New System.Drawing.Size(23, 22)
        Me.TB_Bold.Text = "Bold"
        '
        'TB_Underline
        '
        Me.TB_Underline.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TB_Underline.Image = CType(resources.GetObject("TB_Underline.Image"), System.Drawing.Image)
        Me.TB_Underline.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TB_Underline.Name = "TB_Underline"
        Me.TB_Underline.Size = New System.Drawing.Size(23, 22)
        Me.TB_Underline.Text = "Underline"
        '
        'TB_Bullet
        '
        Me.TB_Bullet.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TB_Bullet.Image = CType(resources.GetObject("TB_Bullet.Image"), System.Drawing.Image)
        Me.TB_Bullet.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TB_Bullet.Name = "TB_Bullet"
        Me.TB_Bullet.Size = New System.Drawing.Size(23, 22)
        Me.TB_Bullet.Text = "Bullets"
        '
        'SaleLog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 515)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SaleLog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: Ras 12 :. Sales Log"
        CType(Me.GridCust, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.GridPendingCall, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.GridDoneCall, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridCust As System.Windows.Forms.DataGridView
    Friend WithEvents TxtPre As System.Windows.Forms.RichTextBox
    Friend WithEvents TxtPost As System.Windows.Forms.RichTextBox
    Friend WithEvents txtStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblMakeAppt As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDone As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtSubj As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TxtPreView As System.Windows.Forms.RichTextBox
    Friend WithEvents GridPendingCall As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents OptLead As System.Windows.Forms.RadioButton
    Friend WithEvents OptCust As System.Windows.Forms.RadioButton
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents LblSum As System.Windows.Forms.Label
    Friend WithEvents txtPostView As System.Windows.Forms.RichTextBox
    Friend WithEvents GridDoneCall As System.Windows.Forms.DataGridView
    Friend WithEvents LblUpdatePrepar As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents LblChange As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtNewEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtNewStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents ChkByPhone As System.Windows.Forms.CheckBox
    Friend WithEvents ChkSelectCust As System.Windows.Forms.CheckBox
    Friend WithEvents txtRptTimeThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtRptTimeFrm As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblRefresh As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbCustType As System.Windows.Forms.ComboBox
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents TB_Color As System.Windows.Forms.ToolStripButton
    Friend WithEvents TB_Bold As System.Windows.Forms.ToolStripButton
    Friend WithEvents TB_Underline As System.Windows.Forms.ToolStripButton
    Friend WithEvents TB_Bullet As System.Windows.Forms.ToolStripButton
    Friend WithEvents LckLblDelete As System.Windows.Forms.LinkLabel
End Class
