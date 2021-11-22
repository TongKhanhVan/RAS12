<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CustChannelLevel
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.ChkListAvail = New System.Windows.Forms.CheckedListBox
        Me.ChkActiveOnly = New System.Windows.Forms.CheckBox
        Me.CmbSBU = New System.Windows.Forms.ComboBox
        Me.TxtValidFrm = New System.Windows.Forms.DateTimePicker
        Me.TxtValidThru = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.CmbAL = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.CmbChannel = New System.Windows.Forms.ComboBox
        Me.LblChannel = New System.Windows.Forms.Label
        Me.CmbLevel = New System.Windows.Forms.ComboBox
        Me.LblUnAssigned = New System.Windows.Forms.LinkLabel
        Me.LblMemberOf = New System.Windows.Forms.LinkLabel
        Me.LblToRight = New System.Windows.Forms.LinkLabel
        Me.GridCurrent = New System.Windows.Forms.DataGridView
        Me.LblDeactive = New System.Windows.Forms.LinkLabel
        Me.ChkQOnly = New System.Windows.Forms.CheckBox
        Me.LckLblApprove = New System.Windows.Forms.LinkLabel
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.LblChangeHL = New System.Windows.Forms.LinkLabel
        Me.LblCheckUnCheckALL = New System.Windows.Forms.LinkLabel
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.CmbFilter = New System.Windows.Forms.ComboBox
        Me.GridSF = New System.Windows.Forms.DataGridView
        Me.GridDiscount = New System.Windows.Forms.DataGridView
        Me.ChkBulkEnxtend = New System.Windows.Forms.CheckBox
        Me.LblExtend = New System.Windows.Forms.LinkLabel
        CType(Me.GridCurrent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.GridSF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ChkListAvail
        '
        Me.ChkListAvail.CheckOnClick = True
        Me.ChkListAvail.ColumnWidth = 90
        Me.ChkListAvail.FormattingEnabled = True
        Me.ChkListAvail.Location = New System.Drawing.Point(0, 25)
        Me.ChkListAvail.MultiColumn = True
        Me.ChkListAvail.Name = "ChkListAvail"
        Me.ChkListAvail.Size = New System.Drawing.Size(384, 394)
        Me.ChkListAvail.TabIndex = 7
        '
        'ChkActiveOnly
        '
        Me.ChkActiveOnly.AutoSize = True
        Me.ChkActiveOnly.Location = New System.Drawing.Point(384, 425)
        Me.ChkActiveOnly.Name = "ChkActiveOnly"
        Me.ChkActiveOnly.Size = New System.Drawing.Size(77, 17)
        Me.ChkActiveOnly.TabIndex = 9
        Me.ChkActiveOnly.Text = "ActiveOnly"
        Me.ChkActiveOnly.UseVisualStyleBackColor = True
        '
        'CmbSBU
        '
        Me.CmbSBU.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbSBU.FormattingEnabled = True
        Me.CmbSBU.Items.AddRange(New Object() {"GSA", "TVS", "EDU"})
        Me.CmbSBU.Location = New System.Drawing.Point(31, 5)
        Me.CmbSBU.Name = "CmbSBU"
        Me.CmbSBU.Size = New System.Drawing.Size(48, 21)
        Me.CmbSBU.TabIndex = 0
        '
        'TxtValidFrm
        '
        Me.TxtValidFrm.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtValidFrm.Location = New System.Drawing.Point(358, 4)
        Me.TxtValidFrm.Name = "TxtValidFrm"
        Me.TxtValidFrm.Size = New System.Drawing.Size(81, 20)
        Me.TxtValidFrm.TabIndex = 4
        '
        'TxtValidThru
        '
        Me.TxtValidThru.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtValidThru.Location = New System.Drawing.Point(498, 4)
        Me.TxtValidThru.Name = "TxtValidThru"
        Me.TxtValidThru.Size = New System.Drawing.Size(77, 20)
        Me.TxtValidThru.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "SBU"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(79, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "AL"
        '
        'CmbAL
        '
        Me.CmbAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(102, 5)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(41, 21)
        Me.CmbAL.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(300, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Valid From"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(441, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Valid Thru"
        '
        'CmbChannel
        '
        Me.CmbChannel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbChannel.FormattingEnabled = True
        Me.CmbChannel.Location = New System.Drawing.Point(191, 5)
        Me.CmbChannel.Name = "CmbChannel"
        Me.CmbChannel.Size = New System.Drawing.Size(44, 21)
        Me.CmbChannel.TabIndex = 2
        '
        'LblChannel
        '
        Me.LblChannel.AutoSize = True
        Me.LblChannel.Location = New System.Drawing.Point(144, 9)
        Me.LblChannel.Name = "LblChannel"
        Me.LblChannel.Size = New System.Drawing.Size(46, 13)
        Me.LblChannel.TabIndex = 5
        Me.LblChannel.Text = "Channel"
        '
        'CmbLevel
        '
        Me.CmbLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbLevel.FormattingEnabled = True
        Me.CmbLevel.Location = New System.Drawing.Point(237, 5)
        Me.CmbLevel.Name = "CmbLevel"
        Me.CmbLevel.Size = New System.Drawing.Size(57, 21)
        Me.CmbLevel.TabIndex = 3
        '
        'LblUnAssigned
        '
        Me.LblUnAssigned.AutoSize = True
        Me.LblUnAssigned.Location = New System.Drawing.Point(2, 8)
        Me.LblUnAssigned.Name = "LblUnAssigned"
        Me.LblUnAssigned.Size = New System.Drawing.Size(116, 13)
        Me.LblUnAssigned.TabIndex = 10
        Me.LblUnAssigned.TabStop = True
        Me.LblUnAssigned.Text = "UnAssigned Customers"
        '
        'LblMemberOf
        '
        Me.LblMemberOf.AutoSize = True
        Me.LblMemberOf.Location = New System.Drawing.Point(381, 10)
        Me.LblMemberOf.Name = "LblMemberOf"
        Me.LblMemberOf.Size = New System.Drawing.Size(62, 13)
        Me.LblMemberOf.TabIndex = 10
        Me.LblMemberOf.TabStop = True
        Me.LblMemberOf.Text = "Member Of "
        '
        'LblToRight
        '
        Me.LblToRight.AutoSize = True
        Me.LblToRight.Location = New System.Drawing.Point(353, 425)
        Me.LblToRight.Name = "LblToRight"
        Me.LblToRight.Size = New System.Drawing.Size(25, 13)
        Me.LblToRight.TabIndex = 11
        Me.LblToRight.TabStop = True
        Me.LblToRight.Text = ">>>"
        '
        'GridCurrent
        '
        Me.GridCurrent.AllowUserToAddRows = False
        Me.GridCurrent.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GridCurrent.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.GridCurrent.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridCurrent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCurrent.Location = New System.Drawing.Point(385, 25)
        Me.GridCurrent.Name = "GridCurrent"
        Me.GridCurrent.RowHeadersVisible = False
        Me.GridCurrent.RowHeadersWidth = 25
        Me.GridCurrent.Size = New System.Drawing.Size(397, 394)
        Me.GridCurrent.TabIndex = 12
        '
        'LblDeactive
        '
        Me.LblDeactive.AutoSize = True
        Me.LblDeactive.Location = New System.Drawing.Point(467, 425)
        Me.LblDeactive.Name = "LblDeactive"
        Me.LblDeactive.Size = New System.Drawing.Size(38, 13)
        Me.LblDeactive.TabIndex = 13
        Me.LblDeactive.TabStop = True
        Me.LblDeactive.Text = "Delete"
        '
        'ChkQOnly
        '
        Me.ChkQOnly.AutoSize = True
        Me.ChkQOnly.Location = New System.Drawing.Point(721, 424)
        Me.ChkQOnly.Name = "ChkQOnly"
        Me.ChkQOnly.Size = New System.Drawing.Size(58, 17)
        Me.ChkQOnly.TabIndex = 14
        Me.ChkQOnly.Text = "Q Only"
        Me.ChkQOnly.UseVisualStyleBackColor = True
        '
        'LckLblApprove
        '
        Me.LckLblApprove.AutoSize = True
        Me.LckLblApprove.Location = New System.Drawing.Point(738, 8)
        Me.LckLblApprove.Name = "LckLblApprove"
        Me.LckLblApprove.Size = New System.Drawing.Size(47, 13)
        Me.LckLblApprove.TabIndex = 13
        Me.LckLblApprove.TabStop = True
        Me.LckLblApprove.Text = "Approve"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(-1, 32)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(790, 473)
        Me.TabControl1.TabIndex = 15
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.LblChangeHL)
        Me.TabPage1.Controls.Add(Me.GridCurrent)
        Me.TabPage1.Controls.Add(Me.LblMemberOf)
        Me.TabPage1.Controls.Add(Me.ChkQOnly)
        Me.TabPage1.Controls.Add(Me.LblUnAssigned)
        Me.TabPage1.Controls.Add(Me.ChkListAvail)
        Me.TabPage1.Controls.Add(Me.LblCheckUnCheckALL)
        Me.TabPage1.Controls.Add(Me.ChkActiveOnly)
        Me.TabPage1.Controls.Add(Me.LblDeactive)
        Me.TabPage1.Controls.Add(Me.LblToRight)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(782, 447)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Customer Assignment"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'LblChangeHL
        '
        Me.LblChangeHL.AutoSize = True
        Me.LblChangeHL.Location = New System.Drawing.Point(511, 425)
        Me.LblChangeHL.Name = "LblChangeHL"
        Me.LblChangeHL.Size = New System.Drawing.Size(77, 13)
        Me.LblChangeHL.TabIndex = 15
        Me.LblChangeHL.TabStop = True
        Me.LblChangeHL.Text = "ChangeValidity"
        '
        'LblCheckUnCheckALL
        '
        Me.LblCheckUnCheckALL.AutoSize = True
        Me.LblCheckUnCheckALL.Location = New System.Drawing.Point(6, 425)
        Me.LblCheckUnCheckALL.Name = "LblCheckUnCheckALL"
        Me.LblCheckUnCheckALL.Size = New System.Drawing.Size(49, 13)
        Me.LblCheckUnCheckALL.TabIndex = 13
        Me.LblCheckUnCheckALL.TabStop = True
        Me.LblCheckUnCheckALL.Tag = "CHK"
        Me.LblCheckUnCheckALL.Text = "CheckAll"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.CmbFilter)
        Me.TabPage2.Controls.Add(Me.GridSF)
        Me.TabPage2.Controls.Add(Me.GridDiscount)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(782, 447)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Current Policy"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'CmbFilter
        '
        Me.CmbFilter.FormattingEnabled = True
        Me.CmbFilter.Items.AddRange(New Object() {"1-Current", "2-Future", "3-Expired", "4-Cancelled"})
        Me.CmbFilter.Location = New System.Drawing.Point(0, 423)
        Me.CmbFilter.Name = "CmbFilter"
        Me.CmbFilter.Size = New System.Drawing.Size(100, 21)
        Me.CmbFilter.TabIndex = 1
        '
        'GridSF
        '
        Me.GridSF.AllowUserToAddRows = False
        Me.GridSF.AllowUserToDeleteRows = False
        Me.GridSF.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridSF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridSF.Location = New System.Drawing.Point(0, 161)
        Me.GridSF.Name = "GridSF"
        Me.GridSF.RowHeadersVisible = False
        Me.GridSF.Size = New System.Drawing.Size(782, 260)
        Me.GridSF.TabIndex = 0
        '
        'GridDiscount
        '
        Me.GridDiscount.AllowUserToAddRows = False
        Me.GridDiscount.AllowUserToDeleteRows = False
        Me.GridDiscount.BackgroundColor = System.Drawing.Color.Azure
        Me.GridDiscount.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridDiscount.Location = New System.Drawing.Point(0, 3)
        Me.GridDiscount.Name = "GridDiscount"
        Me.GridDiscount.RowHeadersVisible = False
        Me.GridDiscount.Size = New System.Drawing.Size(782, 156)
        Me.GridDiscount.TabIndex = 0
        '
        'ChkBulkEnxtend
        '
        Me.ChkBulkEnxtend.AutoSize = True
        Me.ChkBulkEnxtend.Location = New System.Drawing.Point(581, 7)
        Me.ChkBulkEnxtend.Name = "ChkBulkEnxtend"
        Me.ChkBulkEnxtend.Size = New System.Drawing.Size(91, 17)
        Me.ChkBulkEnxtend.TabIndex = 6
        Me.ChkBulkEnxtend.Text = "Bulk Xtension"
        Me.ChkBulkEnxtend.UseVisualStyleBackColor = True
        '
        'LblExtend
        '
        Me.LblExtend.AutoSize = True
        Me.LblExtend.Location = New System.Drawing.Point(687, 8)
        Me.LblExtend.Name = "LblExtend"
        Me.LblExtend.Size = New System.Drawing.Size(40, 13)
        Me.LblExtend.TabIndex = 7
        Me.LblExtend.TabStop = True
        Me.LblExtend.Text = "Extend"
        Me.LblExtend.Visible = False
        '
        'CustChannelLevel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(786, 501)
        Me.Controls.Add(Me.LblExtend)
        Me.Controls.Add(Me.ChkBulkEnxtend)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.CmbLevel)
        Me.Controls.Add(Me.CmbAL)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LckLblApprove)
        Me.Controls.Add(Me.LblChannel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtValidThru)
        Me.Controls.Add(Me.CmbChannel)
        Me.Controls.Add(Me.TxtValidFrm)
        Me.Controls.Add(Me.CmbSBU)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CustChannelLevel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Customer Channel and Level by Airlines"
        CType(Me.GridCurrent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.GridSF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ChkListAvail As System.Windows.Forms.CheckedListBox
    Friend WithEvents ChkActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents CmbSBU As System.Windows.Forms.ComboBox
    Friend WithEvents TxtValidFrm As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtValidThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CmbChannel As System.Windows.Forms.ComboBox
    Friend WithEvents LblChannel As System.Windows.Forms.Label
    Friend WithEvents CmbLevel As System.Windows.Forms.ComboBox
    Friend WithEvents LblUnAssigned As System.Windows.Forms.LinkLabel
    Friend WithEvents LblMemberOf As System.Windows.Forms.LinkLabel
    Friend WithEvents LblToRight As System.Windows.Forms.LinkLabel
    Friend WithEvents GridCurrent As System.Windows.Forms.DataGridView
    Friend WithEvents LblDeactive As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkQOnly As System.Windows.Forms.CheckBox
    Friend WithEvents LckLblApprove As System.Windows.Forms.LinkLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GridSF As System.Windows.Forms.DataGridView
    Friend WithEvents GridDiscount As System.Windows.Forms.DataGridView
    Friend WithEvents LblChangeHL As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCheckUnCheckALL As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbFilter As System.Windows.Forms.ComboBox
    Friend WithEvents ChkBulkEnxtend As System.Windows.Forms.CheckBox
    Friend WithEvents LblExtend As System.Windows.Forms.LinkLabel
End Class
