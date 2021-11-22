<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PromoCode
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
        Me.GridPromoCode = New System.Windows.Forms.DataGridView
        Me.LblAdd = New System.Windows.Forms.LinkLabel
        Me.LblDeactivate = New System.Windows.Forms.LinkLabel
        Me.LblViewFull = New System.Windows.Forms.LinkLabel
        Me.ChkActiveOnly = New System.Windows.Forms.CheckBox
        Me.LblRefresh = New System.Windows.Forms.LinkLabel
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.GridPromoCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridPromoCode
        '
        Me.GridPromoCode.AllowUserToDeleteRows = False
        Me.GridPromoCode.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridPromoCode.BackgroundColor = System.Drawing.Color.SeaShell
        Me.GridPromoCode.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPromoCode.Location = New System.Drawing.Point(2, 25)
        Me.GridPromoCode.Name = "GridPromoCode"
        Me.GridPromoCode.RowHeadersVisible = False
        Me.GridPromoCode.Size = New System.Drawing.Size(785, 429)
        Me.GridPromoCode.TabIndex = 8
        '
        'LblAdd
        '
        Me.LblAdd.AutoSize = True
        Me.LblAdd.Location = New System.Drawing.Point(731, 9)
        Me.LblAdd.Name = "LblAdd"
        Me.LblAdd.Size = New System.Drawing.Size(26, 13)
        Me.LblAdd.TabIndex = 9
        Me.LblAdd.TabStop = True
        Me.LblAdd.Text = "Add"
        '
        'LblDeactivate
        '
        Me.LblDeactivate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblDeactivate.AutoSize = True
        Me.LblDeactivate.Location = New System.Drawing.Point(607, 461)
        Me.LblDeactivate.Name = "LblDeactivate"
        Me.LblDeactivate.Size = New System.Drawing.Size(51, 13)
        Me.LblDeactivate.TabIndex = 9
        Me.LblDeactivate.TabStop = True
        Me.LblDeactivate.Text = "DeActive"
        Me.LblDeactivate.Visible = False
        '
        'LblViewFull
        '
        Me.LblViewFull.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblViewFull.AutoSize = True
        Me.LblViewFull.Location = New System.Drawing.Point(716, 461)
        Me.LblViewFull.Name = "LblViewFull"
        Me.LblViewFull.Size = New System.Drawing.Size(71, 13)
        Me.LblViewFull.TabIndex = 9
        Me.LblViewFull.TabStop = True
        Me.LblViewFull.Text = "MoreColumns"
        '
        'ChkActiveOnly
        '
        Me.ChkActiveOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkActiveOnly.AutoSize = True
        Me.ChkActiveOnly.Checked = True
        Me.ChkActiveOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkActiveOnly.Location = New System.Drawing.Point(2, 460)
        Me.ChkActiveOnly.Name = "ChkActiveOnly"
        Me.ChkActiveOnly.Size = New System.Drawing.Size(77, 17)
        Me.ChkActiveOnly.TabIndex = 10
        Me.ChkActiveOnly.Text = "ActiveOnly"
        Me.ChkActiveOnly.UseVisualStyleBackColor = True
        '
        'LblRefresh
        '
        Me.LblRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblRefresh.AutoSize = True
        Me.LblRefresh.Location = New System.Drawing.Point(85, 461)
        Me.LblRefresh.Name = "LblRefresh"
        Me.LblRefresh.Size = New System.Drawing.Size(44, 13)
        Me.LblRefresh.TabIndex = 11
        Me.LblRefresh.TabStop = True
        Me.LblRefresh.Text = "Refresh"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(-1, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(221, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "PromoCode Should Start With AL Code or TV"
        '
        'PromoCode
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(789, 478)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LblRefresh)
        Me.Controls.Add(Me.ChkActiveOnly)
        Me.Controls.Add(Me.LblViewFull)
        Me.Controls.Add(Me.LblDeactivate)
        Me.Controls.Add(Me.LblAdd)
        Me.Controls.Add(Me.GridPromoCode)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PromoCode"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airline :: RAS :. PromoCode"
        CType(Me.GridPromoCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridPromoCode As System.Windows.Forms.DataGridView
    Friend WithEvents LblAdd As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDeactivate As System.Windows.Forms.LinkLabel
    Friend WithEvents LblViewFull As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents LblRefresh As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
