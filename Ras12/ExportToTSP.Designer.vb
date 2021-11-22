<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExportToTSP
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GridTKT = New System.Windows.Forms.DataGridView()
        Me.C = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.GridTC = New System.Windows.Forms.DataGridView()
        Me.LblCheck = New System.Windows.Forms.LinkLabel()
        Me.LblCheckAll = New System.Windows.Forms.LinkLabel()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.StatusADL = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusCHD = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusINF = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusTTL = New System.Windows.Forms.ToolStripStatusLabel()
        Me.LblExport = New System.Windows.Forms.LinkLabel()
        Me.txtUser = New System.Windows.Forms.TextBox()
        CType(Me.GridTKT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridTC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridTKT
        '
        Me.GridTKT.AllowUserToAddRows = False
        Me.GridTKT.AllowUserToDeleteRows = False
        Me.GridTKT.BackgroundColor = System.Drawing.Color.Ivory
        Me.GridTKT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTKT.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.C})
        Me.GridTKT.Location = New System.Drawing.Point(119, 3)
        Me.GridTKT.Name = "GridTKT"
        Me.GridTKT.RowHeadersVisible = False
        Me.GridTKT.Size = New System.Drawing.Size(659, 445)
        Me.GridTKT.TabIndex = 1
        '
        'C
        '
        Me.C.Frozen = True
        Me.C.HeaderText = "C"
        Me.C.Name = "C"
        Me.C.Width = 24
        '
        'GridTC
        '
        Me.GridTC.AllowUserToAddRows = False
        Me.GridTC.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.Azure
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        Me.GridTC.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.GridTC.BackgroundColor = System.Drawing.Color.AliceBlue
        Me.GridTC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridTC.Location = New System.Drawing.Point(2, 3)
        Me.GridTC.Name = "GridTC"
        Me.GridTC.RowHeadersVisible = False
        Me.GridTC.Size = New System.Drawing.Size(116, 445)
        Me.GridTC.TabIndex = 2
        '
        'LblCheck
        '
        Me.LblCheck.AutoSize = True
        Me.LblCheck.Location = New System.Drawing.Point(697, 454)
        Me.LblCheck.Name = "LblCheck"
        Me.LblCheck.Size = New System.Drawing.Size(38, 13)
        Me.LblCheck.TabIndex = 3
        Me.LblCheck.TabStop = True
        Me.LblCheck.Text = "Check"
        '
        'LblCheckAll
        '
        Me.LblCheckAll.AutoSize = True
        Me.LblCheckAll.Location = New System.Drawing.Point(124, 454)
        Me.LblCheckAll.Name = "LblCheckAll"
        Me.LblCheckAll.Size = New System.Drawing.Size(48, 13)
        Me.LblCheckAll.TabIndex = 5
        Me.LblCheckAll.TabStop = True
        Me.LblCheckAll.Text = "SelectAll"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusADL, Me.StatusCHD, Me.StatusINF, Me.StatusTTL})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 474)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(781, 22)
        Me.StatusStrip1.TabIndex = 6
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'StatusADL
        '
        Me.StatusADL.ForeColor = System.Drawing.Color.Blue
        Me.StatusADL.Name = "StatusADL"
        Me.StatusADL.Size = New System.Drawing.Size(0, 17)
        '
        'StatusCHD
        '
        Me.StatusCHD.ForeColor = System.Drawing.Color.DodgerBlue
        Me.StatusCHD.Name = "StatusCHD"
        Me.StatusCHD.Size = New System.Drawing.Size(0, 17)
        '
        'StatusINF
        '
        Me.StatusINF.ForeColor = System.Drawing.Color.Blue
        Me.StatusINF.Name = "StatusINF"
        Me.StatusINF.Size = New System.Drawing.Size(0, 17)
        '
        'StatusTTL
        '
        Me.StatusTTL.Name = "StatusTTL"
        Me.StatusTTL.Size = New System.Drawing.Size(0, 17)
        '
        'LblExport
        '
        Me.LblExport.AutoSize = True
        Me.LblExport.Location = New System.Drawing.Point(741, 454)
        Me.LblExport.Name = "LblExport"
        Me.LblExport.Size = New System.Drawing.Size(37, 13)
        Me.LblExport.TabIndex = 7
        Me.LblExport.TabStop = True
        Me.LblExport.Text = "Export"
        '
        'txtUser
        '
        Me.txtUser.BackColor = System.Drawing.SystemColors.MenuBar
        Me.txtUser.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtUser.Location = New System.Drawing.Point(2, 451)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(116, 20)
        Me.txtUser.TabIndex = 10
        Me.txtUser.Text = "NNK, AUT,NTT,LTH"
        '
        'ExportToTSP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.txtUser)
        Me.Controls.Add(Me.LblExport)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.LblCheckAll)
        Me.Controls.Add(Me.LblCheck)
        Me.Controls.Add(Me.GridTC)
        Me.Controls.Add(Me.GridTKT)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ExportToTSP"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS 12 :. Export To TSP"
        CType(Me.GridTKT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridTC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridTKT As System.Windows.Forms.DataGridView
    Friend WithEvents GridTC As System.Windows.Forms.DataGridView
    Friend WithEvents C As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents LblCheck As System.Windows.Forms.LinkLabel
    Friend WithEvents LblCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusADL As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusCHD As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusINF As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusTTL As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents LblExport As System.Windows.Forms.LinkLabel
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
End Class
