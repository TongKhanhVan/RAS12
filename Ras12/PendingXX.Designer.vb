<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PendingXX
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GridPendingXX = New System.Windows.Forms.DataGridView()
        Me.OptTVS = New System.Windows.Forms.RadioButton()
        Me.OptGSA = New System.Windows.Forms.RadioButton()
        Me.LblInvalidTKNO = New System.Windows.Forms.LinkLabel()
        Me.TxtLstDay = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GridTKNO = New System.Windows.Forms.DataGridView()
        Me.TxtLyDo = New System.Windows.Forms.TextBox()
        Me.LblClear = New System.Windows.Forms.LinkLabel()
        Me.LblChecked = New System.Windows.Forms.LinkLabel()
        CType(Me.GridPendingXX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridTKNO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridPendingXX
        '
        Me.GridPendingXX.AllowUserToAddRows = False
        Me.GridPendingXX.AllowUserToDeleteRows = False
        Me.GridPendingXX.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridPendingXX.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPendingXX.Location = New System.Drawing.Point(2, 29)
        Me.GridPendingXX.Name = "GridPendingXX"
        Me.GridPendingXX.RowHeadersVisible = False
        Me.GridPendingXX.Size = New System.Drawing.Size(389, 371)
        Me.GridPendingXX.TabIndex = 0
        '
        'OptTVS
        '
        Me.OptTVS.AutoSize = True
        Me.OptTVS.Checked = True
        Me.OptTVS.Location = New System.Drawing.Point(2, 6)
        Me.OptTVS.Name = "OptTVS"
        Me.OptTVS.Size = New System.Drawing.Size(46, 17)
        Me.OptTVS.TabIndex = 1
        Me.OptTVS.TabStop = True
        Me.OptTVS.Text = "TVS"
        Me.OptTVS.UseVisualStyleBackColor = True
        '
        'OptGSA
        '
        Me.OptGSA.AutoSize = True
        Me.OptGSA.Location = New System.Drawing.Point(47, 6)
        Me.OptGSA.Name = "OptGSA"
        Me.OptGSA.Size = New System.Drawing.Size(47, 17)
        Me.OptGSA.TabIndex = 1
        Me.OptGSA.Text = "GSA"
        Me.OptGSA.UseVisualStyleBackColor = True
        '
        'LblInvalidTKNO
        '
        Me.LblInvalidTKNO.AutoSize = True
        Me.LblInvalidTKNO.Location = New System.Drawing.Point(656, 383)
        Me.LblInvalidTKNO.Name = "LblInvalidTKNO"
        Me.LblInvalidTKNO.Size = New System.Drawing.Size(68, 13)
        Me.LblInvalidTKNO.TabIndex = 2
        Me.LblInvalidTKNO.TabStop = True
        Me.LblInvalidTKNO.Text = "InvalidTKNO"
        Me.LblInvalidTKNO.Visible = False
        '
        'TxtLstDay
        '
        Me.TxtLstDay.Location = New System.Drawing.Point(123, 5)
        Me.TxtLstDay.Name = "TxtLstDay"
        Me.TxtLstDay.Size = New System.Drawing.Size(33, 20)
        Me.TxtLstDay.TabIndex = 3
        Me.TxtLstDay.Text = "16"
        Me.TxtLstDay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(96, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Last"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(156, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "days"
        '
        'GridTKNO
        '
        Me.GridTKNO.AllowUserToAddRows = False
        Me.GridTKNO.AllowUserToDeleteRows = False
        Me.GridTKNO.BackgroundColor = System.Drawing.SystemColors.Control
        Me.GridTKNO.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.AppWorkspace
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.GridTKNO.DefaultCellStyle = DataGridViewCellStyle2
        Me.GridTKNO.Enabled = False
        Me.GridTKNO.Location = New System.Drawing.Point(394, 29)
        Me.GridTKNO.Name = "GridTKNO"
        Me.GridTKNO.RowHeadersVisible = False
        Me.GridTKNO.Size = New System.Drawing.Size(331, 348)
        Me.GridTKNO.TabIndex = 5
        '
        'TxtLyDo
        '
        Me.TxtLyDo.Location = New System.Drawing.Point(394, 380)
        Me.TxtLyDo.Name = "TxtLyDo"
        Me.TxtLyDo.Size = New System.Drawing.Size(177, 20)
        Me.TxtLyDo.TabIndex = 6
        '
        'LblClear
        '
        Me.LblClear.AutoSize = True
        Me.LblClear.Location = New System.Drawing.Point(577, 383)
        Me.LblClear.Name = "LblClear"
        Me.LblClear.Size = New System.Drawing.Size(31, 13)
        Me.LblClear.TabIndex = 7
        Me.LblClear.TabStop = True
        Me.LblClear.Tag = "Clear"
        Me.LblClear.Text = "Clear"
        Me.LblClear.Visible = False
        '
        'LblChecked
        '
        Me.LblChecked.AutoSize = True
        Me.LblChecked.Location = New System.Drawing.Point(607, 383)
        Me.LblChecked.Name = "LblChecked"
        Me.LblChecked.Size = New System.Drawing.Size(50, 13)
        Me.LblChecked.TabIndex = 8
        Me.LblChecked.TabStop = True
        Me.LblChecked.Tag = "Checked"
        Me.LblChecked.Text = "Checked"
        '
        'PendingXX
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(725, 402)
        Me.Controls.Add(Me.LblChecked)
        Me.Controls.Add(Me.LblClear)
        Me.Controls.Add(Me.TxtLyDo)
        Me.Controls.Add(Me.GridTKNO)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtLstDay)
        Me.Controls.Add(Me.LblInvalidTKNO)
        Me.Controls.Add(Me.OptGSA)
        Me.Controls.Add(Me.OptTVS)
        Me.Controls.Add(Me.GridPendingXX)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PendingXX"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Pending XX Alert"
        CType(Me.GridPendingXX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridTKNO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridPendingXX As System.Windows.Forms.DataGridView
    Friend WithEvents OptTVS As System.Windows.Forms.RadioButton
    Friend WithEvents OptGSA As System.Windows.Forms.RadioButton
    Friend WithEvents LblInvalidTKNO As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtLstDay As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GridTKNO As System.Windows.Forms.DataGridView
    Friend WithEvents TxtLyDo As System.Windows.Forms.TextBox
    Friend WithEvents LblClear As System.Windows.Forms.LinkLabel
    Friend WithEvents LblChecked As System.Windows.Forms.LinkLabel
End Class
