<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CustInfor
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtSearch = New System.Windows.Forms.TextBox
        Me.LblShowFull = New System.Windows.Forms.LinkLabel
        Me.LblFindShort = New System.Windows.Forms.LinkLabel
        Me.GridCust = New System.Windows.Forms.DataGridView
        Me.LblViewPolicy = New System.Windows.Forms.LinkLabel
        Me.GridDiscount = New System.Windows.Forms.DataGridView
        Me.GridSF = New System.Windows.Forms.DataGridView
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        CType(Me.GridCust, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridSF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(-1, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(189, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Type Short Name or Part of Full Name "
        '
        'TxtSearch
        '
        Me.TxtSearch.Location = New System.Drawing.Point(1, 19)
        Me.TxtSearch.Name = "TxtSearch"
        Me.TxtSearch.Size = New System.Drawing.Size(95, 20)
        Me.TxtSearch.TabIndex = 2
        '
        'LblShowFull
        '
        Me.LblShowFull.AutoSize = True
        Me.LblShowFull.Location = New System.Drawing.Point(96, 22)
        Me.LblShowFull.Name = "LblShowFull"
        Me.LblShowFull.Size = New System.Drawing.Size(84, 13)
        Me.LblShowFull.TabIndex = 3
        Me.LblShowFull.TabStop = True
        Me.LblShowFull.Text = "Show Full Name"
        '
        'LblFindShort
        '
        Me.LblFindShort.AutoSize = True
        Me.LblFindShort.Location = New System.Drawing.Point(179, 22)
        Me.LblFindShort.Name = "LblFindShort"
        Me.LblFindShort.Size = New System.Drawing.Size(86, 13)
        Me.LblFindShort.TabIndex = 3
        Me.LblFindShort.TabStop = True
        Me.LblFindShort.Text = "Find Short Name"
        '
        'GridCust
        '
        Me.GridCust.AllowUserToAddRows = False
        Me.GridCust.AllowUserToDeleteRows = False
        Me.GridCust.BackgroundColor = System.Drawing.Color.Ivory
        Me.GridCust.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCust.GridColor = System.Drawing.SystemColors.ButtonFace
        Me.GridCust.Location = New System.Drawing.Point(1, 40)
        Me.GridCust.Name = "GridCust"
        Me.GridCust.RowHeadersVisible = False
        Me.GridCust.Size = New System.Drawing.Size(263, 460)
        Me.GridCust.TabIndex = 4
        '
        'LblViewPolicy
        '
        Me.LblViewPolicy.AutoSize = True
        Me.LblViewPolicy.Enabled = False
        Me.LblViewPolicy.Location = New System.Drawing.Point(271, 22)
        Me.LblViewPolicy.Name = "LblViewPolicy"
        Me.LblViewPolicy.Size = New System.Drawing.Size(61, 13)
        Me.LblViewPolicy.TabIndex = 5
        Me.LblViewPolicy.TabStop = True
        Me.LblViewPolicy.Text = "View Policy"
        '
        'GridDiscount
        '
        Me.GridDiscount.AllowUserToAddRows = False
        Me.GridDiscount.AllowUserToDeleteRows = False
        Me.GridDiscount.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridDiscount.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridDiscount.Location = New System.Drawing.Point(265, 40)
        Me.GridDiscount.Name = "GridDiscount"
        Me.GridDiscount.RowHeadersVisible = False
        Me.GridDiscount.Size = New System.Drawing.Size(522, 162)
        Me.GridDiscount.TabIndex = 6
        '
        'GridSF
        '
        Me.GridSF.AllowUserToAddRows = False
        Me.GridSF.AllowUserToDeleteRows = False
        Me.GridSF.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridSF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridSF.Location = New System.Drawing.Point(266, 221)
        Me.GridSF.Name = "GridSF"
        Me.GridSF.RowHeadersVisible = False
        Me.GridSF.Size = New System.Drawing.Size(522, 279)
        Me.GridSF.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(735, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Discount "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(724, 205)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Service Fee"
        '
        'CustInfor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 501)
        Me.Controls.Add(Me.GridSF)
        Me.Controls.Add(Me.GridDiscount)
        Me.Controls.Add(Me.LblViewPolicy)
        Me.Controls.Add(Me.GridCust)
        Me.Controls.Add(Me.LblFindShort)
        Me.Controls.Add(Me.LblShowFull)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtSearch)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CustInfor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Customer Infor"
        CType(Me.GridCust, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridSF, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblShowFull As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtSearch As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblFindShort As System.Windows.Forms.LinkLabel
    Friend WithEvents GridCust As System.Windows.Forms.DataGridView
    Friend WithEvents LblViewPolicy As System.Windows.Forms.LinkLabel
    Friend WithEvents GridDiscount As System.Windows.Forms.DataGridView
    Friend WithEvents GridSF As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
