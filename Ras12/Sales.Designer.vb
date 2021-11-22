<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Sales
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
        Me.PnlFB_RBD = New System.Windows.Forms.Panel()
        Me.CmdFilter = New System.Windows.Forms.Button()
        Me.txtAL = New System.Windows.Forms.TextBox()
        Me.CmdSave = New System.Windows.Forms.Button()
        Me.GridFB_RBD = New System.Windows.Forms.DataGridView()
        Me.GrpReport = New System.Windows.Forms.GroupBox()
        Me.ChkCompare = New System.Windows.Forms.CheckBox()
        Me.CmbTime = New System.Windows.Forms.ComboBox()
        Me.LblCancel = New System.Windows.Forms.LinkLabel()
        Me.LblRun = New System.Windows.Forms.LinkLabel()
        Me.ChkCust = New System.Windows.Forms.CheckBox()
        Me.ChkAL = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CmbCust = New System.Windows.Forms.ComboBox()
        Me.CmbCounter = New System.Windows.Forms.ComboBox()
        Me.txtThru = New System.Windows.Forms.DateTimePicker()
        Me.TxtFrm = New System.Windows.Forms.DateTimePicker()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.StatusFeedback = New System.Windows.Forms.ToolStripStatusLabel()
        Me.PnlFB_RBD.SuspendLayout()
        CType(Me.GridFB_RBD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GrpReport.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PnlFB_RBD
        '
        Me.PnlFB_RBD.Controls.Add(Me.CmdFilter)
        Me.PnlFB_RBD.Controls.Add(Me.txtAL)
        Me.PnlFB_RBD.Controls.Add(Me.CmdSave)
        Me.PnlFB_RBD.Controls.Add(Me.GridFB_RBD)
        Me.PnlFB_RBD.Location = New System.Drawing.Point(2, 3)
        Me.PnlFB_RBD.Name = "PnlFB_RBD"
        Me.PnlFB_RBD.Size = New System.Drawing.Size(304, 60)
        Me.PnlFB_RBD.TabIndex = 0
        Me.PnlFB_RBD.Visible = False
        '
        'CmdFilter
        '
        Me.CmdFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdFilter.Location = New System.Drawing.Point(243, 3)
        Me.CmdFilter.Name = "CmdFilter"
        Me.CmdFilter.Size = New System.Drawing.Size(55, 25)
        Me.CmdFilter.TabIndex = 3
        Me.CmdFilter.Text = "Filter"
        Me.CmdFilter.UseVisualStyleBackColor = True
        '
        'txtAL
        '
        Me.txtAL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAL.Location = New System.Drawing.Point(211, 6)
        Me.txtAL.Name = "txtAL"
        Me.txtAL.Size = New System.Drawing.Size(26, 20)
        Me.txtAL.TabIndex = 2
        '
        'CmdSave
        '
        Me.CmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdSave.Location = New System.Drawing.Point(209, 32)
        Me.CmdSave.Name = "CmdSave"
        Me.CmdSave.Size = New System.Drawing.Size(89, 25)
        Me.CmdSave.TabIndex = 1
        Me.CmdSave.Text = "Save"
        Me.CmdSave.UseVisualStyleBackColor = True
        '
        'GridFB_RBD
        '
        Me.GridFB_RBD.AllowUserToAddRows = False
        Me.GridFB_RBD.AllowUserToDeleteRows = False
        Me.GridFB_RBD.AllowUserToResizeColumns = False
        Me.GridFB_RBD.AllowUserToResizeRows = False
        Me.GridFB_RBD.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridFB_RBD.BackgroundColor = System.Drawing.Color.Azure
        Me.GridFB_RBD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridFB_RBD.Location = New System.Drawing.Point(2, 0)
        Me.GridFB_RBD.Name = "GridFB_RBD"
        Me.GridFB_RBD.Size = New System.Drawing.Size(204, 58)
        Me.GridFB_RBD.TabIndex = 0
        '
        'GrpReport
        '
        Me.GrpReport.Controls.Add(Me.ChkCompare)
        Me.GrpReport.Controls.Add(Me.CmbTime)
        Me.GrpReport.Controls.Add(Me.LblCancel)
        Me.GrpReport.Controls.Add(Me.LblRun)
        Me.GrpReport.Controls.Add(Me.ChkCust)
        Me.GrpReport.Controls.Add(Me.ChkAL)
        Me.GrpReport.Controls.Add(Me.Label2)
        Me.GrpReport.Controls.Add(Me.Label3)
        Me.GrpReport.Controls.Add(Me.Label1)
        Me.GrpReport.Controls.Add(Me.CmbCust)
        Me.GrpReport.Controls.Add(Me.CmbCounter)
        Me.GrpReport.Controls.Add(Me.txtThru)
        Me.GrpReport.Controls.Add(Me.TxtFrm)
        Me.GrpReport.Location = New System.Drawing.Point(4, 89)
        Me.GrpReport.Name = "GrpReport"
        Me.GrpReport.Size = New System.Drawing.Size(368, 119)
        Me.GrpReport.TabIndex = 1
        Me.GrpReport.TabStop = False
        Me.GrpReport.Text = "Report"
        Me.GrpReport.Visible = False
        '
        'ChkCompare
        '
        Me.ChkCompare.AutoSize = True
        Me.ChkCompare.Checked = True
        Me.ChkCompare.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkCompare.Location = New System.Drawing.Point(169, 19)
        Me.ChkCompare.Name = "ChkCompare"
        Me.ChkCompare.Size = New System.Drawing.Size(153, 17)
        Me.ChkCompare.TabIndex = 8
        Me.ChkCompare.Text = "vs Performance in the Past"
        Me.ChkCompare.UseVisualStyleBackColor = True
        '
        'CmbTime
        '
        Me.CmbTime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbTime.FormattingEnabled = True
        Me.CmbTime.Items.AddRange(New Object() {"This MTD", "This QTD", "This YTD", "Last Month", "Last Quarter", "Last Quarter", "Custome"})
        Me.CmbTime.Location = New System.Drawing.Point(61, 17)
        Me.CmbTime.Name = "CmbTime"
        Me.CmbTime.Size = New System.Drawing.Size(102, 21)
        Me.CmbTime.TabIndex = 6
        '
        'LblCancel
        '
        Me.LblCancel.AutoSize = True
        Me.LblCancel.Location = New System.Drawing.Point(236, 97)
        Me.LblCancel.Name = "LblCancel"
        Me.LblCancel.Size = New System.Drawing.Size(33, 13)
        Me.LblCancel.TabIndex = 5
        Me.LblCancel.TabStop = True
        Me.LblCancel.Text = "Close"
        '
        'LblRun
        '
        Me.LblRun.AutoSize = True
        Me.LblRun.Location = New System.Drawing.Point(58, 97)
        Me.LblRun.Name = "LblRun"
        Me.LblRun.Size = New System.Drawing.Size(62, 13)
        Me.LblRun.TabIndex = 5
        Me.LblRun.TabStop = True
        Me.LblRun.Text = "Run Report"
        '
        'ChkCust
        '
        Me.ChkCust.AutoSize = True
        Me.ChkCust.Location = New System.Drawing.Point(169, 46)
        Me.ChkCust.Name = "ChkCust"
        Me.ChkCust.Size = New System.Drawing.Size(70, 17)
        Me.ChkCust.TabIndex = 4
        Me.ChkCust.Text = "Customer"
        Me.ChkCust.UseVisualStyleBackColor = True
        '
        'ChkAL
        '
        Me.ChkAL.AutoSize = True
        Me.ChkAL.Location = New System.Drawing.Point(169, 68)
        Me.ChkAL.Name = "ChkAL"
        Me.ChkAL.Size = New System.Drawing.Size(63, 17)
        Me.ChkAL.TabIndex = 4
        Me.ChkAL.Text = "Counter"
        Me.ChkAL.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Thru"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "TimeInput"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "From"
        '
        'CmbCust
        '
        Me.CmbCust.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCust.Enabled = False
        Me.CmbCust.FormattingEnabled = True
        Me.CmbCust.Location = New System.Drawing.Point(239, 42)
        Me.CmbCust.Name = "CmbCust"
        Me.CmbCust.Size = New System.Drawing.Size(119, 21)
        Me.CmbCust.TabIndex = 2
        '
        'CmbCounter
        '
        Me.CmbCounter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCounter.Enabled = False
        Me.CmbCounter.FormattingEnabled = True
        Me.CmbCounter.Location = New System.Drawing.Point(239, 65)
        Me.CmbCounter.Name = "CmbCounter"
        Me.CmbCounter.Size = New System.Drawing.Size(43, 21)
        Me.CmbCounter.TabIndex = 1
        '
        'txtThru
        '
        Me.txtThru.Enabled = False
        Me.txtThru.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtThru.Location = New System.Drawing.Point(61, 66)
        Me.txtThru.Name = "txtThru"
        Me.txtThru.Size = New System.Drawing.Size(102, 20)
        Me.txtThru.TabIndex = 0
        '
        'TxtFrm
        '
        Me.TxtFrm.Enabled = False
        Me.TxtFrm.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtFrm.Location = New System.Drawing.Point(61, 44)
        Me.TxtFrm.Name = "TxtFrm"
        Me.TxtFrm.Size = New System.Drawing.Size(102, 20)
        Me.TxtFrm.TabIndex = 0
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusFeedback})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 422)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(812, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'StatusFeedback
        '
        Me.StatusFeedback.BackColor = System.Drawing.Color.Transparent
        Me.StatusFeedback.ForeColor = System.Drawing.Color.Crimson
        Me.StatusFeedback.Name = "StatusFeedback"
        Me.StatusFeedback.Size = New System.Drawing.Size(0, 17)
        '
        'Sales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(812, 444)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.GrpReport)
        Me.Controls.Add(Me.PnlFB_RBD)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Sales"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Sales"
        Me.PnlFB_RBD.ResumeLayout(False)
        Me.PnlFB_RBD.PerformLayout()
        CType(Me.GridFB_RBD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GrpReport.ResumeLayout(False)
        Me.GrpReport.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PnlFB_RBD As System.Windows.Forms.Panel
    Friend WithEvents GridFB_RBD As System.Windows.Forms.DataGridView
    Friend WithEvents CmdSave As System.Windows.Forms.Button
    Friend WithEvents CmdFilter As System.Windows.Forms.Button
    Friend WithEvents txtAL As System.Windows.Forms.TextBox
    Friend WithEvents GrpReport As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CmbCust As System.Windows.Forms.ComboBox
    Friend WithEvents CmbCounter As System.Windows.Forms.ComboBox
    Friend WithEvents txtThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtFrm As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents LblRun As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkCust As System.Windows.Forms.CheckBox
    Friend WithEvents ChkAL As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CmbTime As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ChkCompare As System.Windows.Forms.CheckBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusFeedback As System.Windows.Forms.ToolStripStatusLabel
End Class
