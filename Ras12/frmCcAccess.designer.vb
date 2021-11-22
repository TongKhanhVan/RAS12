<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCcAccess
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
        Me.LblSave = New System.Windows.Forms.LinkLabel()
        Me.dgCompanies = New System.Windows.Forms.DataGridView()
        Me.dgUsers = New System.Windows.Forms.DataGridView()
        Me.lbkSelectAll = New System.Windows.Forms.LinkLabel()
        Me.lbkUnselectAll = New System.Windows.Forms.LinkLabel()
        CType(Me.dgCompanies, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgUsers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblSave
        '
        Me.LblSave.AutoSize = True
        Me.LblSave.Location = New System.Drawing.Point(919, 601)
        Me.LblSave.Name = "LblSave"
        Me.LblSave.Size = New System.Drawing.Size(32, 13)
        Me.LblSave.TabIndex = 55
        Me.LblSave.TabStop = True
        Me.LblSave.Text = "Save"
        '
        'dgCompanies
        '
        Me.dgCompanies.AllowUserToAddRows = False
        Me.dgCompanies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCompanies.Location = New System.Drawing.Point(488, 11)
        Me.dgCompanies.Name = "dgCompanies"
        Me.dgCompanies.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgCompanies.Size = New System.Drawing.Size(516, 587)
        Me.dgCompanies.TabIndex = 56
        '
        'dgUsers
        '
        Me.dgUsers.AllowUserToAddRows = False
        Me.dgUsers.AllowUserToDeleteRows = False
        Me.dgUsers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgUsers.Location = New System.Drawing.Point(12, 11)
        Me.dgUsers.Name = "dgUsers"
        Me.dgUsers.ReadOnly = True
        Me.dgUsers.Size = New System.Drawing.Size(460, 587)
        Me.dgUsers.TabIndex = 57
        '
        'lbkSelectAll
        '
        Me.lbkSelectAll.AutoSize = True
        Me.lbkSelectAll.Location = New System.Drawing.Point(485, 601)
        Me.lbkSelectAll.Name = "lbkSelectAll"
        Me.lbkSelectAll.Size = New System.Drawing.Size(48, 13)
        Me.lbkSelectAll.TabIndex = 58
        Me.lbkSelectAll.TabStop = True
        Me.lbkSelectAll.Text = "SelectAll"
        '
        'lbkUnselectAll
        '
        Me.lbkUnselectAll.AutoSize = True
        Me.lbkUnselectAll.Location = New System.Drawing.Point(539, 601)
        Me.lbkUnselectAll.Name = "lbkUnselectAll"
        Me.lbkUnselectAll.Size = New System.Drawing.Size(62, 13)
        Me.lbkUnselectAll.TabIndex = 59
        Me.lbkUnselectAll.TabStop = True
        Me.lbkUnselectAll.Text = "UnSelectAll"
        '
        'frmCcAccess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 623)
        Me.Controls.Add(Me.lbkUnselectAll)
        Me.Controls.Add(Me.lbkSelectAll)
        Me.Controls.Add(Me.dgUsers)
        Me.Controls.Add(Me.dgCompanies)
        Me.Controls.Add(Me.LblSave)
        Me.Name = "frmCcAccess"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CreditCardcAccess"
        CType(Me.dgCompanies, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgUsers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblSave As System.Windows.Forms.LinkLabel
    Friend WithEvents dgCompanies As System.Windows.Forms.DataGridView
    Friend WithEvents dgUsers As System.Windows.Forms.DataGridView
    Friend WithEvents lbkSelectAll As System.Windows.Forms.LinkLabel
    Friend WithEvents lbkUnselectAll As System.Windows.Forms.LinkLabel
End Class
