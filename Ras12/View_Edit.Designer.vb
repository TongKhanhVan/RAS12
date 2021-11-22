<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class View
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
        Me.OptBySearch = New System.Windows.Forms.RadioButton()
        Me.OptByDate = New System.Windows.Forms.RadioButton()
        Me.CmbFindWhat = New System.Windows.Forms.ComboBox()
        Me.txtFrmDate = New System.Windows.Forms.DateTimePicker()
        Me.txtToDate = New System.Windows.Forms.DateTimePicker()
        Me.CmdSearch = New System.Windows.Forms.Button()
        Me.CmbAL = New System.Windows.Forms.ComboBox()
        Me.LblAL = New System.Windows.Forms.Label()
        Me.GridRCP = New System.Windows.Forms.DataGridView()
        Me.GridFOP = New System.Windows.Forms.DataGridView()
        Me.LblReprint = New System.Windows.Forms.LinkLabel()
        Me.TxtRPTNo = New System.Windows.Forms.TextBox()
        Me.cmbValueToSearch = New System.Windows.Forms.ComboBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TxtTC = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblSearchTC = New System.Windows.Forms.LinkLabel()
        Me.GridTKTinTC = New System.Windows.Forms.DataGridView()
        CType(Me.GridRCP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridFOP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.GridTKTinTC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OptBySearch
        '
        Me.OptBySearch.AutoSize = True
        Me.OptBySearch.Location = New System.Drawing.Point(150, 6)
        Me.OptBySearch.Name = "OptBySearch"
        Me.OptBySearch.Size = New System.Drawing.Size(61, 17)
        Me.OptBySearch.TabIndex = 515
        Me.OptBySearch.Text = "Specific"
        Me.OptBySearch.UseVisualStyleBackColor = True
        '
        'OptByDate
        '
        Me.OptByDate.AutoSize = True
        Me.OptByDate.Checked = True
        Me.OptByDate.Location = New System.Drawing.Point(59, 6)
        Me.OptByDate.Name = "OptByDate"
        Me.OptByDate.Size = New System.Drawing.Size(93, 17)
        Me.OptByDate.TabIndex = 515
        Me.OptByDate.TabStop = True
        Me.OptByDate.Text = "Date between"
        Me.OptByDate.UseVisualStyleBackColor = True
        '
        'CmbFindWhat
        '
        Me.CmbFindWhat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbFindWhat.FormattingEnabled = True
        Me.CmbFindWhat.Items.AddRange(New Object() {"TRX No.", "TKT No.", "RPT No.", "RLOC", "Policy", "TCODE"})
        Me.CmbFindWhat.Location = New System.Drawing.Point(212, 4)
        Me.CmbFindWhat.Name = "CmbFindWhat"
        Me.CmbFindWhat.Size = New System.Drawing.Size(69, 21)
        Me.CmbFindWhat.TabIndex = 515
        Me.CmbFindWhat.Visible = False
        '
        'txtFrmDate
        '
        Me.txtFrmDate.CustomFormat = "dd-MMM-yy"
        Me.txtFrmDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtFrmDate.Location = New System.Drawing.Point(212, 4)
        Me.txtFrmDate.Name = "txtFrmDate"
        Me.txtFrmDate.Size = New System.Drawing.Size(77, 21)
        Me.txtFrmDate.TabIndex = 30
        '
        'txtToDate
        '
        Me.txtToDate.CustomFormat = "dd-MMM-yy"
        Me.txtToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtToDate.Location = New System.Drawing.Point(291, 4)
        Me.txtToDate.Name = "txtToDate"
        Me.txtToDate.Size = New System.Drawing.Size(73, 21)
        Me.txtToDate.TabIndex = 30
        '
        'CmdSearch
        '
        Me.CmdSearch.Location = New System.Drawing.Point(422, 4)
        Me.CmdSearch.Name = "CmdSearch"
        Me.CmdSearch.Size = New System.Drawing.Size(38, 22)
        Me.CmdSearch.TabIndex = 25
        Me.CmdSearch.Text = "GO"
        Me.CmdSearch.UseVisualStyleBackColor = True
        '
        'CmbAL
        '
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(19, 4)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(38, 21)
        Me.CmbAL.TabIndex = 517
        '
        'LblAL
        '
        Me.LblAL.AutoSize = True
        Me.LblAL.Location = New System.Drawing.Point(0, 7)
        Me.LblAL.Name = "LblAL"
        Me.LblAL.Size = New System.Drawing.Size(19, 13)
        Me.LblAL.TabIndex = 518
        Me.LblAL.Text = "AL"
        '
        'GridRCP
        '
        Me.GridRCP.AllowUserToAddRows = False
        Me.GridRCP.AllowUserToDeleteRows = False
        Me.GridRCP.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridRCP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridRCP.Location = New System.Drawing.Point(2, 29)
        Me.GridRCP.Name = "GridRCP"
        Me.GridRCP.RowHeadersVisible = False
        Me.GridRCP.Size = New System.Drawing.Size(776, 284)
        Me.GridRCP.TabIndex = 519
        '
        'GridFOP
        '
        Me.GridFOP.AllowUserToAddRows = False
        Me.GridFOP.AllowUserToDeleteRows = False
        Me.GridFOP.BackgroundColor = System.Drawing.Color.Azure
        Me.GridFOP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridFOP.Location = New System.Drawing.Point(2, 314)
        Me.GridFOP.Name = "GridFOP"
        Me.GridFOP.RowHeadersVisible = False
        Me.GridFOP.Size = New System.Drawing.Size(776, 131)
        Me.GridFOP.TabIndex = 519
        '
        'LblReprint
        '
        Me.LblReprint.AutoSize = True
        Me.LblReprint.Enabled = False
        Me.LblReprint.Location = New System.Drawing.Point(581, 451)
        Me.LblReprint.Name = "LblReprint"
        Me.LblReprint.Size = New System.Drawing.Size(94, 13)
        Me.LblReprint.TabIndex = 521
        Me.LblReprint.TabStop = True
        Me.LblReprint.Text = "RePrintThisReport"
        '
        'TxtRPTNo
        '
        Me.TxtRPTNo.Enabled = False
        Me.TxtRPTNo.Location = New System.Drawing.Point(681, 448)
        Me.TxtRPTNo.Name = "TxtRPTNo"
        Me.TxtRPTNo.Size = New System.Drawing.Size(96, 21)
        Me.TxtRPTNo.TabIndex = 522
        '
        'cmbValueToSearch
        '
        Me.cmbValueToSearch.FormattingEnabled = True
        Me.cmbValueToSearch.Location = New System.Drawing.Point(365, 4)
        Me.cmbValueToSearch.Name = "cmbValueToSearch"
        Me.cmbValueToSearch.Size = New System.Drawing.Size(53, 21)
        Me.cmbValueToSearch.TabIndex = 523
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(0, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(788, 498)
        Me.TabControl1.TabIndex = 524
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.LblReprint)
        Me.TabPage1.Controls.Add(Me.CmdSearch)
        Me.TabPage1.Controls.Add(Me.TxtRPTNo)
        Me.TabPage1.Controls.Add(Me.CmbFindWhat)
        Me.TabPage1.Controls.Add(Me.GridFOP)
        Me.TabPage1.Controls.Add(Me.txtToDate)
        Me.TabPage1.Controls.Add(Me.GridRCP)
        Me.TabPage1.Controls.Add(Me.txtFrmDate)
        Me.TabPage1.Controls.Add(Me.OptBySearch)
        Me.TabPage1.Controls.Add(Me.cmbValueToSearch)
        Me.TabPage1.Controls.Add(Me.LblAL)
        Me.TabPage1.Controls.Add(Me.OptByDate)
        Me.TabPage1.Controls.Add(Me.CmbAL)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(780, 472)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Nostalgia"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GridTKTinTC)
        Me.TabPage2.Controls.Add(Me.LblSearchTC)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.TxtTC)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(780, 472)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "TourCode"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TxtTC
        '
        Me.TxtTC.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTC.Location = New System.Drawing.Point(66, 2)
        Me.TxtTC.Name = "TxtTC"
        Me.TxtTC.Size = New System.Drawing.Size(126, 21)
        Me.TxtTC.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "TourCode"
        '
        'LblSearchTC
        '
        Me.LblSearchTC.AutoSize = True
        Me.LblSearchTC.Location = New System.Drawing.Point(198, 5)
        Me.LblSearchTC.Name = "LblSearchTC"
        Me.LblSearchTC.Size = New System.Drawing.Size(40, 13)
        Me.LblSearchTC.TabIndex = 2
        Me.LblSearchTC.TabStop = True
        Me.LblSearchTC.Text = "Search"
        '
        'GridTKTinTC
        '
        Me.GridTKTinTC.AllowUserToAddRows = False
        Me.GridTKTinTC.AllowUserToDeleteRows = False
        Me.GridTKTinTC.BackgroundColor = System.Drawing.Color.Azure
        Me.GridTKTinTC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTKTinTC.Location = New System.Drawing.Point(3, 27)
        Me.GridTKTinTC.Name = "GridTKTinTC"
        Me.GridTKTinTC.RowHeadersVisible = False
        Me.GridTKTinTC.Size = New System.Drawing.Size(777, 442)
        Me.GridTKTinTC.TabIndex = 3
        '
        'View
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(791, 501)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "View"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS12 :. Transaction Listing and View"
        CType(Me.GridRCP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridFOP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.GridTKTinTC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CmdSearch As System.Windows.Forms.Button
    Friend WithEvents CmbFindWhat As System.Windows.Forms.ComboBox
    Friend WithEvents OptBySearch As System.Windows.Forms.RadioButton
    Friend WithEvents OptByDate As System.Windows.Forms.RadioButton
    Friend WithEvents txtFrmDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents LblAL As System.Windows.Forms.Label
    Friend WithEvents GridRCP As System.Windows.Forms.DataGridView
    Friend WithEvents GridFOP As System.Windows.Forms.DataGridView
    Friend WithEvents LblReprint As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtRPTNo As System.Windows.Forms.TextBox
    Friend WithEvents cmbValueToSearch As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GridTKTinTC As System.Windows.Forms.DataGridView
    Friend WithEvents LblSearchTC As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtTC As System.Windows.Forms.TextBox
End Class
