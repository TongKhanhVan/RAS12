<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ClearPendingPmt
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
        Me.GrpPaymentDetails = New System.Windows.Forms.GroupBox()
        Me.GridFOP = New System.Windows.Forms.DataGridView()
        Me.FOP = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.RCPCurrency = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.RCPROE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Amount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Document = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CmdApply = New System.Windows.Forms.Button()
        Me.GridPendingRCP = New System.Windows.Forms.DataGridView()
        Me.CmbRCVDby = New System.Windows.Forms.ComboBox()
        Me.TxtDate = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CmdClearAnother = New System.Windows.Forms.Button()
        Me.PnlClearPendingPmt = New System.Windows.Forms.Panel()
        Me.lblAddCc = New System.Windows.Forms.LinkLabel()
        Me.LblViewTRX = New System.Windows.Forms.LinkLabel()
        Me.LblLoad = New System.Windows.Forms.LinkLabel()
        Me.LblAL = New System.Windows.Forms.Label()
        Me.CmbAL = New System.Windows.Forms.ComboBox()
        Me.OptALL = New System.Windows.Forms.RadioButton()
        Me.OptDEB = New System.Windows.Forms.RadioButton()
        Me.OptCRD = New System.Windows.Forms.RadioButton()
        Me.GrpPaymentDetails.SuspendLayout()
        CType(Me.GridFOP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridPendingRCP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlClearPendingPmt.SuspendLayout()
        Me.SuspendLayout()
        '
        'GrpPaymentDetails
        '
        Me.GrpPaymentDetails.Controls.Add(Me.GridFOP)
        Me.GrpPaymentDetails.Controls.Add(Me.CmdApply)
        Me.GrpPaymentDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GrpPaymentDetails.Location = New System.Drawing.Point(3, 234)
        Me.GrpPaymentDetails.Name = "GrpPaymentDetails"
        Me.GrpPaymentDetails.Size = New System.Drawing.Size(883, 231)
        Me.GrpPaymentDetails.TabIndex = 8
        Me.GrpPaymentDetails.TabStop = False
        Me.GrpPaymentDetails.Text = "Form Of Payment"
        '
        'GridFOP
        '
        Me.GridFOP.BackgroundColor = System.Drawing.SystemColors.ControlLight
        Me.GridFOP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.GridFOP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridFOP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FOP, Me.RCPCurrency, Me.RCPROE, Me.Amount, Me.Document})
        Me.GridFOP.Location = New System.Drawing.Point(-1, 30)
        Me.GridFOP.Name = "GridFOP"
        Me.GridFOP.RowHeadersWidth = 30
        Me.GridFOP.Size = New System.Drawing.Size(535, 159)
        Me.GridFOP.TabIndex = 5
        '
        'FOP
        '
        Me.FOP.HeaderText = "FOP"
        Me.FOP.Name = "FOP"
        Me.FOP.Width = 70
        '
        'RCPCurrency
        '
        Me.RCPCurrency.HeaderText = "Currency"
        Me.RCPCurrency.Name = "RCPCurrency"
        Me.RCPCurrency.Width = 70
        '
        'RCPROE
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.RCPROE.DefaultCellStyle = DataGridViewCellStyle3
        Me.RCPROE.HeaderText = "ROE"
        Me.RCPROE.Name = "RCPROE"
        Me.RCPROE.ReadOnly = True
        Me.RCPROE.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.RCPROE.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.RCPROE.Width = 75
        '
        'Amount
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.Format = "#,##0.00"
        DataGridViewCellStyle4.NullValue = "0"
        Me.Amount.DefaultCellStyle = DataGridViewCellStyle4
        Me.Amount.HeaderText = "Amount"
        Me.Amount.Name = "Amount"
        '
        'Document
        '
        Me.Document.HeaderText = "Document"
        Me.Document.Name = "Document"
        '
        'CmdApply
        '
        Me.CmdApply.Enabled = False
        Me.CmdApply.Location = New System.Drawing.Point(540, 159)
        Me.CmdApply.Name = "CmdApply"
        Me.CmdApply.Size = New System.Drawing.Size(94, 30)
        Me.CmdApply.TabIndex = 6
        Me.CmdApply.Text = "Apply"
        Me.CmdApply.UseVisualStyleBackColor = True
        '
        'GridPendingRCP
        '
        Me.GridPendingRCP.AllowUserToAddRows = False
        Me.GridPendingRCP.AllowUserToDeleteRows = False
        Me.GridPendingRCP.BackgroundColor = System.Drawing.Color.AliceBlue
        Me.GridPendingRCP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPendingRCP.Location = New System.Drawing.Point(3, 3)
        Me.GridPendingRCP.MultiSelect = False
        Me.GridPendingRCP.Name = "GridPendingRCP"
        Me.GridPendingRCP.RowHeadersVisible = False
        Me.GridPendingRCP.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridPendingRCP.Size = New System.Drawing.Size(762, 196)
        Me.GridPendingRCP.TabIndex = 9
        '
        'CmbRCVDby
        '
        Me.CmbRCVDby.Enabled = False
        Me.CmbRCVDby.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbRCVDby.FormattingEnabled = True
        Me.CmbRCVDby.Location = New System.Drawing.Point(252, 208)
        Me.CmbRCVDby.Name = "CmbRCVDby"
        Me.CmbRCVDby.Size = New System.Drawing.Size(112, 24)
        Me.CmbRCVDby.TabIndex = 11
        '
        'TxtDate
        '
        Me.TxtDate.CustomFormat = "dd-MMM-yy"
        Me.TxtDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TxtDate.Location = New System.Drawing.Point(43, 208)
        Me.TxtDate.Name = "TxtDate"
        Me.TxtDate.Size = New System.Drawing.Size(94, 22)
        Me.TxtDate.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(4, 212)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 16)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Date"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(160, 212)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Received By"
        '
        'CmdClearAnother
        '
        Me.CmdClearAnother.Location = New System.Drawing.Point(770, 33)
        Me.CmdClearAnother.Name = "CmdClearAnother"
        Me.CmdClearAnother.Size = New System.Drawing.Size(110, 21)
        Me.CmdClearAnother.TabIndex = 14
        Me.CmdClearAnother.Text = "Clear Another TRX"
        Me.CmdClearAnother.UseVisualStyleBackColor = True
        '
        'PnlClearPendingPmt
        '
        Me.PnlClearPendingPmt.Controls.Add(Me.lblAddCc)
        Me.PnlClearPendingPmt.Controls.Add(Me.LblViewTRX)
        Me.PnlClearPendingPmt.Controls.Add(Me.LblLoad)
        Me.PnlClearPendingPmt.Controls.Add(Me.GridPendingRCP)
        Me.PnlClearPendingPmt.Controls.Add(Me.CmdClearAnother)
        Me.PnlClearPendingPmt.Controls.Add(Me.LblAL)
        Me.PnlClearPendingPmt.Controls.Add(Me.CmbAL)
        Me.PnlClearPendingPmt.Location = New System.Drawing.Point(0, 1)
        Me.PnlClearPendingPmt.Name = "PnlClearPendingPmt"
        Me.PnlClearPendingPmt.Size = New System.Drawing.Size(886, 202)
        Me.PnlClearPendingPmt.TabIndex = 15
        '
        'lblAddCc
        '
        Me.lblAddCc.AutoSize = True
        Me.lblAddCc.Location = New System.Drawing.Point(771, 158)
        Me.lblAddCc.Name = "lblAddCc"
        Me.lblAddCc.Size = New System.Drawing.Size(92, 13)
        Me.lblAddCc.TabIndex = 19
        Me.lblAddCc.TabStop = True
        Me.lblAddCc.Text = "AddCreditCardNbr"
        '
        'LblViewTRX
        '
        Me.LblViewTRX.AutoSize = True
        Me.LblViewTRX.Location = New System.Drawing.Point(771, 181)
        Me.LblViewTRX.Name = "LblViewTRX"
        Me.LblViewTRX.Size = New System.Drawing.Size(55, 13)
        Me.LblViewTRX.TabIndex = 18
        Me.LblViewTRX.TabStop = True
        Me.LblViewTRX.Text = "View TRX"
        '
        'LblLoad
        '
        Me.LblLoad.AutoSize = True
        Me.LblLoad.Location = New System.Drawing.Point(852, 9)
        Me.LblLoad.Name = "LblLoad"
        Me.LblLoad.Size = New System.Drawing.Size(31, 13)
        Me.LblLoad.TabIndex = 17
        Me.LblLoad.TabStop = True
        Me.LblLoad.Text = "Load"
        '
        'LblAL
        '
        Me.LblAL.AutoSize = True
        Me.LblAL.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblAL.Location = New System.Drawing.Point(771, 9)
        Me.LblAL.Name = "LblAL"
        Me.LblAL.Size = New System.Drawing.Size(22, 13)
        Me.LblAL.TabIndex = 9
        Me.LblAL.Text = "AL"
        '
        'CmbAL
        '
        Me.CmbAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(799, 6)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(52, 21)
        Me.CmbAL.TabIndex = 12
        '
        'OptALL
        '
        Me.OptALL.AutoSize = True
        Me.OptALL.Checked = True
        Me.OptALL.Location = New System.Drawing.Point(370, 211)
        Me.OptALL.Name = "OptALL"
        Me.OptALL.Size = New System.Drawing.Size(44, 17)
        Me.OptALL.TabIndex = 16
        Me.OptALL.TabStop = True
        Me.OptALL.Text = "ALL"
        Me.OptALL.UseVisualStyleBackColor = True
        '
        'OptDEB
        '
        Me.OptDEB.AutoSize = True
        Me.OptDEB.Location = New System.Drawing.Point(420, 211)
        Me.OptDEB.Name = "OptDEB"
        Me.OptDEB.Size = New System.Drawing.Size(47, 17)
        Me.OptDEB.TabIndex = 16
        Me.OptDEB.Text = "DEB"
        Me.OptDEB.UseVisualStyleBackColor = True
        '
        'OptCRD
        '
        Me.OptCRD.AutoSize = True
        Me.OptCRD.Location = New System.Drawing.Point(473, 212)
        Me.OptCRD.Name = "OptCRD"
        Me.OptCRD.Size = New System.Drawing.Size(48, 17)
        Me.OptCRD.TabIndex = 16
        Me.OptCRD.Text = "CRD"
        Me.OptCRD.UseVisualStyleBackColor = True
        '
        'ClearPendingPmt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(892, 423)
        Me.Controls.Add(Me.OptCRD)
        Me.Controls.Add(Me.OptDEB)
        Me.Controls.Add(Me.OptALL)
        Me.Controls.Add(Me.PnlClearPendingPmt)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtDate)
        Me.Controls.Add(Me.CmbRCVDby)
        Me.Controls.Add(Me.GrpPaymentDetails)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ClearPendingPmt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Clear Pending Payment"
        Me.GrpPaymentDetails.ResumeLayout(False)
        CType(Me.GridFOP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridPendingRCP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlClearPendingPmt.ResumeLayout(False)
        Me.PnlClearPendingPmt.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GrpPaymentDetails As System.Windows.Forms.GroupBox
    Friend WithEvents GridFOP As System.Windows.Forms.DataGridView
    Friend WithEvents GridPendingRCP As System.Windows.Forms.DataGridView
    Friend WithEvents CmdApply As System.Windows.Forms.Button
    Friend WithEvents CmbRCVDby As System.Windows.Forms.ComboBox
    Friend WithEvents TxtDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CmdClearAnother As System.Windows.Forms.Button
    Friend WithEvents PnlClearPendingPmt As System.Windows.Forms.Panel
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents LblAL As System.Windows.Forms.Label
    Friend WithEvents LblLoad As System.Windows.Forms.LinkLabel
    Friend WithEvents FOP As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents RCPCurrency As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents RCPROE As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Document As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OptALL As System.Windows.Forms.RadioButton
    Friend WithEvents OptDEB As System.Windows.Forms.RadioButton
    Friend WithEvents OptCRD As System.Windows.Forms.RadioButton
    Friend WithEvents LblViewTRX As System.Windows.Forms.LinkLabel
    Friend WithEvents lblAddCc As System.Windows.Forms.LinkLabel
End Class
