<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class B2BUserMap
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
        Me.GridUser = New System.Windows.Forms.DataGridView()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.LblMap = New System.Windows.Forms.LinkLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CmbCust = New System.Windows.Forms.ComboBox()
        Me.ChkXXOnly = New System.Windows.Forms.CheckBox()
        Me.LblReinstate = New System.Windows.Forms.LinkLabel()
        CType(Me.GridUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridUser
        '
        Me.GridUser.AllowUserToAddRows = False
        Me.GridUser.AllowUserToDeleteRows = False
        Me.GridUser.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridUser.Location = New System.Drawing.Point(1, 1)
        Me.GridUser.Name = "GridUser"
        Me.GridUser.RowHeadersVisible = False
        Me.GridUser.Size = New System.Drawing.Size(578, 347)
        Me.GridUser.TabIndex = 0
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(0, 357)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 1
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        '
        'LblMap
        '
        Me.LblMap.AutoSize = True
        Me.LblMap.Location = New System.Drawing.Point(537, 357)
        Me.LblMap.Name = "LblMap"
        Me.LblMap.Size = New System.Drawing.Size(42, 13)
        Me.LblMap.TabIndex = 1
        Me.LblMap.TabStop = True
        Me.LblMap.Text = "Update"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(315, 357)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "ShortName"
        '
        'CmbCust
        '
        Me.CmbCust.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCust.FormattingEnabled = True
        Me.CmbCust.Location = New System.Drawing.Point(377, 354)
        Me.CmbCust.Name = "CmbCust"
        Me.CmbCust.Size = New System.Drawing.Size(158, 21)
        Me.CmbCust.TabIndex = 3
        '
        'ChkXXOnly
        '
        Me.ChkXXOnly.AutoSize = True
        Me.ChkXXOnly.Location = New System.Drawing.Point(44, 356)
        Me.ChkXXOnly.Name = "ChkXXOnly"
        Me.ChkXXOnly.Size = New System.Drawing.Size(64, 17)
        Me.ChkXXOnly.TabIndex = 5
        Me.ChkXXOnly.Text = "XX Only"
        Me.ChkXXOnly.UseVisualStyleBackColor = True
        '
        'LblReinstate
        '
        Me.LblReinstate.AutoSize = True
        Me.LblReinstate.Location = New System.Drawing.Point(114, 357)
        Me.LblReinstate.Name = "LblReinstate"
        Me.LblReinstate.Size = New System.Drawing.Size(52, 13)
        Me.LblReinstate.TabIndex = 6
        Me.LblReinstate.TabStop = True
        Me.LblReinstate.Text = "Reinstate"
        '
        'B2BUserMap
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(579, 379)
        Me.Controls.Add(Me.LblReinstate)
        Me.Controls.Add(Me.ChkXXOnly)
        Me.Controls.Add(Me.CmbCust)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LblMap)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.GridUser)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "B2BUserMap"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS 12 :. TransViet.vn User Manager"
        CType(Me.GridUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridUser As System.Windows.Forms.DataGridView
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LblMap As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CmbCust As System.Windows.Forms.ComboBox
    Friend WithEvents ChkXXOnly As System.Windows.Forms.CheckBox
    Friend WithEvents LblReinstate As System.Windows.Forms.LinkLabel
End Class
