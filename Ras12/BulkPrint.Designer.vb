<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BulkPrint
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
        Me.OptGSA = New System.Windows.Forms.RadioButton()
        Me.OptTVS = New System.Windows.Forms.RadioButton()
        Me.GridTRX = New System.Windows.Forms.DataGridView()
        Me.S = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.LblPreview = New System.Windows.Forms.LinkLabel()
        Me.OptCorp = New System.Windows.Forms.RadioButton()
        Me.CmbMonth = New System.Windows.Forms.ComboBox()
        Me.cmbCust = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ChkByCust = New System.Windows.Forms.CheckBox()
        CType(Me.GridTRX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OptGSA
        '
        Me.OptGSA.AutoSize = True
        Me.OptGSA.Checked = True
        Me.OptGSA.Location = New System.Drawing.Point(58, 7)
        Me.OptGSA.Name = "OptGSA"
        Me.OptGSA.Size = New System.Drawing.Size(47, 17)
        Me.OptGSA.TabIndex = 0
        Me.OptGSA.TabStop = True
        Me.OptGSA.Text = "GSA"
        Me.OptGSA.UseVisualStyleBackColor = True
        '
        'OptTVS
        '
        Me.OptTVS.AutoSize = True
        Me.OptTVS.Location = New System.Drawing.Point(114, 7)
        Me.OptTVS.Name = "OptTVS"
        Me.OptTVS.Size = New System.Drawing.Size(103, 17)
        Me.OptTVS.TabIndex = 0
        Me.OptTVS.Text = "TVS (Excl. Corp)"
        Me.OptTVS.UseVisualStyleBackColor = True
        '
        'GridTRX
        '
        Me.GridTRX.AllowUserToAddRows = False
        Me.GridTRX.AllowUserToDeleteRows = False
        Me.GridTRX.BackgroundColor = System.Drawing.Color.Azure
        Me.GridTRX.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTRX.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.S})
        Me.GridTRX.Location = New System.Drawing.Point(2, 33)
        Me.GridTRX.Name = "GridTRX"
        Me.GridTRX.RowHeadersVisible = False
        Me.GridTRX.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridTRX.Size = New System.Drawing.Size(785, 467)
        Me.GridTRX.TabIndex = 2
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
        Me.LblPreview.Location = New System.Drawing.Point(749, 9)
        Me.LblPreview.Name = "LblPreview"
        Me.LblPreview.Size = New System.Drawing.Size(28, 13)
        Me.LblPreview.TabIndex = 3
        Me.LblPreview.TabStop = True
        Me.LblPreview.Text = "Print"
        '
        'OptCorp
        '
        Me.OptCorp.AutoSize = True
        Me.OptCorp.Location = New System.Drawing.Point(224, 7)
        Me.OptCorp.Name = "OptCorp"
        Me.OptCorp.Size = New System.Drawing.Size(77, 17)
        Me.OptCorp.TabIndex = 0
        Me.OptCorp.Text = "Corp (BSP)"
        Me.OptCorp.UseVisualStyleBackColor = True
        '
        'CmbMonth
        '
        Me.CmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbMonth.FormattingEnabled = True
        Me.CmbMonth.Location = New System.Drawing.Point(2, 6)
        Me.CmbMonth.Name = "CmbMonth"
        Me.CmbMonth.Size = New System.Drawing.Size(53, 21)
        Me.CmbMonth.TabIndex = 4
        '
        'cmbCust
        '
        Me.cmbCust.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCust.FormattingEnabled = True
        Me.cmbCust.Location = New System.Drawing.Point(434, 6)
        Me.cmbCust.Name = "cmbCust"
        Me.cmbCust.Size = New System.Drawing.Size(146, 21)
        Me.cmbCust.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(403, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Cust."
        '
        'ChkByCust
        '
        Me.ChkByCust.AutoSize = True
        Me.ChkByCust.Location = New System.Drawing.Point(586, 8)
        Me.ChkByCust.Name = "ChkByCust"
        Me.ChkByCust.Size = New System.Drawing.Size(91, 17)
        Me.ChkByCust.TabIndex = 8
        Me.ChkByCust.Text = "Load by Cust."
        Me.ChkByCust.UseVisualStyleBackColor = True
        '
        'BulkPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.ChkByCust)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbCust)
        Me.Controls.Add(Me.CmbMonth)
        Me.Controls.Add(Me.LblPreview)
        Me.Controls.Add(Me.GridTRX)
        Me.Controls.Add(Me.OptCorp)
        Me.Controls.Add(Me.OptTVS)
        Me.Controls.Add(Me.OptGSA)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BulkPrint"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Bulk INV Printing"
        CType(Me.GridTRX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OptGSA As System.Windows.Forms.RadioButton
    Friend WithEvents OptTVS As System.Windows.Forms.RadioButton
    Friend WithEvents GridTRX As System.Windows.Forms.DataGridView
    Friend WithEvents LblPreview As System.Windows.Forms.LinkLabel
    Friend WithEvents OptCorp As System.Windows.Forms.RadioButton
    Friend WithEvents S As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CmbMonth As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCust As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ChkByCust As System.Windows.Forms.CheckBox
End Class
