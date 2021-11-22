<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SCB
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
        Me.GridUNC = New System.Windows.Forms.DataGridView()
        Me.S = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.LblExport = New System.Windows.Forms.LinkLabel()
        Me.LblUpdateBatchNo = New System.Windows.Forms.LinkLabel()
        Me.ChkPending = New System.Windows.Forms.CheckBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.CmbAcctToExp = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtLotNo = New System.Windows.Forms.TextBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.GridUNCinBatch = New System.Windows.Forms.DataGridView()
        Me.GridBatch = New System.Windows.Forms.DataGridView()
        CType(Me.GridUNC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.GridUNCinBatch, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridBatch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridUNC
        '
        Me.GridUNC.AllowUserToAddRows = False
        Me.GridUNC.AllowUserToDeleteRows = False
        Me.GridUNC.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridUNC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridUNC.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.S})
        Me.GridUNC.Location = New System.Drawing.Point(5, 6)
        Me.GridUNC.Name = "GridUNC"
        Me.GridUNC.RowHeadersVisible = False
        Me.GridUNC.Size = New System.Drawing.Size(760, 437)
        Me.GridUNC.TabIndex = 0
        '
        'S
        '
        Me.S.HeaderText = "S"
        Me.S.Name = "S"
        Me.S.Width = 32
        '
        'LblExport
        '
        Me.LblExport.AutoSize = True
        Me.LblExport.Location = New System.Drawing.Point(6, 450)
        Me.LblExport.Name = "LblExport"
        Me.LblExport.Size = New System.Drawing.Size(37, 13)
        Me.LblExport.TabIndex = 1
        Me.LblExport.TabStop = True
        Me.LblExport.Text = "Export"
        '
        'LblUpdateBatchNo
        '
        Me.LblUpdateBatchNo.AutoSize = True
        Me.LblUpdateBatchNo.Enabled = False
        Me.LblUpdateBatchNo.Location = New System.Drawing.Point(626, 450)
        Me.LblUpdateBatchNo.Name = "LblUpdateBatchNo"
        Me.LblUpdateBatchNo.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdateBatchNo.TabIndex = 1
        Me.LblUpdateBatchNo.TabStop = True
        Me.LblUpdateBatchNo.Text = "Update"
        '
        'ChkPending
        '
        Me.ChkPending.AutoSize = True
        Me.ChkPending.Location = New System.Drawing.Point(669, 449)
        Me.ChkPending.Name = "ChkPending"
        Me.ChkPending.Size = New System.Drawing.Size(95, 17)
        Me.ChkPending.TabIndex = 4
        Me.ChkPending.Text = "Show Pending"
        Me.ChkPending.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(2, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(778, 495)
        Me.TabControl1.TabIndex = 5
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.CmbAcctToExp)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.txtAmount)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.txtLotNo)
        Me.TabPage1.Controls.Add(Me.GridUNC)
        Me.TabPage1.Controls.Add(Me.LblUpdateBatchNo)
        Me.TabPage1.Controls.Add(Me.ChkPending)
        Me.TabPage1.Controls.Add(Me.LblExport)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(770, 469)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Export"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'CmbAcctToExp
        '
        Me.CmbAcctToExp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAcctToExp.FormattingEnabled = True
        Me.CmbAcctToExp.Location = New System.Drawing.Point(49, 447)
        Me.CmbAcctToExp.Name = "CmbAcctToExp"
        Me.CmbAcctToExp.Size = New System.Drawing.Size(94, 21)
        Me.CmbAcctToExp.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(153, 450)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "SelectedAmt"
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(223, 447)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(107, 20)
        Me.txtAmount.TabIndex = 7
        Me.txtAmount.Text = "0"
        Me.txtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(456, 450)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "BatchNo."
        '
        'txtLotNo
        '
        Me.txtLotNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtLotNo.Location = New System.Drawing.Point(514, 447)
        Me.txtLotNo.Name = "txtLotNo"
        Me.txtLotNo.Size = New System.Drawing.Size(106, 20)
        Me.txtLotNo.TabIndex = 5
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.LblDelete)
        Me.TabPage2.Controls.Add(Me.GridUNCinBatch)
        Me.TabPage2.Controls.Add(Me.GridBatch)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(770, 469)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "View_Edit"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(6, 442)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 4
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        '
        'GridUNCinBatch
        '
        Me.GridUNCinBatch.AllowUserToAddRows = False
        Me.GridUNCinBatch.AllowUserToDeleteRows = False
        Me.GridUNCinBatch.BackgroundColor = System.Drawing.Color.Azure
        Me.GridUNCinBatch.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridUNCinBatch.Location = New System.Drawing.Point(151, 6)
        Me.GridUNCinBatch.Name = "GridUNCinBatch"
        Me.GridUNCinBatch.ReadOnly = True
        Me.GridUNCinBatch.RowHeadersVisible = False
        Me.GridUNCinBatch.Size = New System.Drawing.Size(616, 433)
        Me.GridUNCinBatch.TabIndex = 3
        '
        'GridBatch
        '
        Me.GridBatch.AllowUserToAddRows = False
        Me.GridBatch.AllowUserToDeleteRows = False
        Me.GridBatch.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridBatch.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridBatch.Location = New System.Drawing.Point(7, 6)
        Me.GridBatch.Name = "GridBatch"
        Me.GridBatch.ReadOnly = True
        Me.GridBatch.RowHeadersVisible = False
        Me.GridBatch.Size = New System.Drawing.Size(138, 433)
        Me.GridBatch.TabIndex = 2
        '
        'SCB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SCB"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS 12 :. SCB Straight2Bank"
        CType(Me.GridUNC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.GridUNCinBatch, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridBatch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GridUNC As System.Windows.Forms.DataGridView
    Friend WithEvents S As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents LblExport As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUpdateBatchNo As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkPending As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents GridUNCinBatch As System.Windows.Forms.DataGridView
    Friend WithEvents GridBatch As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtLotNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents CmbAcctToExp As System.Windows.Forms.ComboBox
End Class
