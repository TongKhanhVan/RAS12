<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OverDue_CreditReport
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GridOverLimit = New System.Windows.Forms.DataGridView()
        Me.GridOverDue = New System.Windows.Forms.DataGridView()
        Me.LblGo = New System.Windows.Forms.LinkLabel()
        Me.GridUser = New System.Windows.Forms.DataGridView()
        Me._Item = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Qty = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtThru = New System.Windows.Forms.DateTimePicker()
        Me.txtFrm = New System.Windows.Forms.DateTimePicker()
        Me.CmbUser = New System.Windows.Forms.ComboBox()
        CType(Me.GridOverLimit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridOverDue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridOverLimit
        '
        Me.GridOverLimit.AllowUserToAddRows = False
        Me.GridOverLimit.AllowUserToDeleteRows = False
        Me.GridOverLimit.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridOverLimit.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridOverLimit.Location = New System.Drawing.Point(1, 3)
        Me.GridOverLimit.Name = "GridOverLimit"
        Me.GridOverLimit.RowHeadersVisible = False
        Me.GridOverLimit.Size = New System.Drawing.Size(456, 489)
        Me.GridOverLimit.TabIndex = 0
        '
        'GridOverDue
        '
        Me.GridOverDue.AllowUserToAddRows = False
        Me.GridOverDue.AllowUserToDeleteRows = False
        Me.GridOverDue.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridOverDue.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridOverDue.Location = New System.Drawing.Point(460, 139)
        Me.GridOverDue.Name = "GridOverDue"
        Me.GridOverDue.RowHeadersVisible = False
        Me.GridOverDue.Size = New System.Drawing.Size(318, 353)
        Me.GridOverDue.TabIndex = 0
        '
        'LblGo
        '
        Me.LblGo.AutoSize = True
        Me.LblGo.Location = New System.Drawing.Point(755, 6)
        Me.LblGo.Name = "LblGo"
        Me.LblGo.Size = New System.Drawing.Size(23, 13)
        Me.LblGo.TabIndex = 3
        Me.LblGo.TabStop = True
        Me.LblGo.Text = "GO"
        '
        'GridUser
        '
        Me.GridUser.AllowUserToAddRows = False
        Me.GridUser.AllowUserToDeleteRows = False
        Me.GridUser.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me._Item, Me.Qty})
        Me.GridUser.Location = New System.Drawing.Point(460, 27)
        Me.GridUser.Name = "GridUser"
        Me.GridUser.RowHeadersVisible = False
        Me.GridUser.Size = New System.Drawing.Size(318, 106)
        Me.GridUser.TabIndex = 3
        '
        '_Item
        '
        Me._Item.HeaderText = "Item"
        Me._Item.Name = "_Item"
        Me._Item.Width = 128
        '
        'Qty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.Qty.DefaultCellStyle = DataGridViewCellStyle3
        Me.Qty.HeaderText = "Qty"
        Me.Qty.Name = "Qty"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(577, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Thru"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(458, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "From"
        '
        'txtThru
        '
        Me.txtThru.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtThru.Location = New System.Drawing.Point(606, 3)
        Me.txtThru.Name = "txtThru"
        Me.txtThru.Size = New System.Drawing.Size(89, 20)
        Me.txtThru.TabIndex = 4
        '
        'txtFrm
        '
        Me.txtFrm.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.txtFrm.Location = New System.Drawing.Point(488, 3)
        Me.txtFrm.Name = "txtFrm"
        Me.txtFrm.Size = New System.Drawing.Size(89, 20)
        Me.txtFrm.TabIndex = 5
        '
        'CmbUser
        '
        Me.CmbUser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbUser.FormattingEnabled = True
        Me.CmbUser.Items.AddRange(New Object() {"LTR", "NKN"})
        Me.CmbUser.Location = New System.Drawing.Point(697, 2)
        Me.CmbUser.Name = "CmbUser"
        Me.CmbUser.Size = New System.Drawing.Size(54, 21)
        Me.CmbUser.TabIndex = 8
        '
        'OverDue_CreditReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.CmbUser)
        Me.Controls.Add(Me.LblGo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.GridUser)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.GridOverDue)
        Me.Controls.Add(Me.txtThru)
        Me.Controls.Add(Me.GridOverLimit)
        Me.Controls.Add(Me.txtFrm)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "OverDue_CreditReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: Ras12:. Over Due_Over Credit Report"
        CType(Me.GridOverLimit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridOverDue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridOverDue As System.Windows.Forms.DataGridView
    Friend WithEvents GridOverLimit As System.Windows.Forms.DataGridView
    Friend WithEvents LblGo As System.Windows.Forms.LinkLabel
    Friend WithEvents GridUser As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtFrm As System.Windows.Forms.DateTimePicker
    Friend WithEvents CmbUser As System.Windows.Forms.ComboBox
    Friend WithEvents _Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Qty As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
