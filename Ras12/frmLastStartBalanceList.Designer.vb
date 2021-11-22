<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStartBalanceList
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
        Me.cboCustomers = New System.Windows.Forms.ComboBox()
        Me.dgStartBalance = New System.Windows.Forms.DataGridView()
        Me.lbkAdd = New System.Windows.Forms.LinkLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.dgStartBalance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboCustomers
        '
        Me.cboCustomers.FormattingEnabled = True
        Me.cboCustomers.Location = New System.Drawing.Point(94, 12)
        Me.cboCustomers.Name = "cboCustomers"
        Me.cboCustomers.Size = New System.Drawing.Size(121, 21)
        Me.cboCustomers.TabIndex = 0
        '
        'dgStartBalance
        '
        Me.dgStartBalance.AllowUserToAddRows = False
        Me.dgStartBalance.AllowUserToDeleteRows = False
        Me.dgStartBalance.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgStartBalance.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgStartBalance.Location = New System.Drawing.Point(12, 39)
        Me.dgStartBalance.Name = "dgStartBalance"
        Me.dgStartBalance.ReadOnly = True
        Me.dgStartBalance.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgStartBalance.Size = New System.Drawing.Size(674, 622)
        Me.dgStartBalance.TabIndex = 1
        '
        'lbkAdd
        '
        Me.lbkAdd.AutoSize = True
        Me.lbkAdd.Location = New System.Drawing.Point(621, 20)
        Me.lbkAdd.Name = "lbkAdd"
        Me.lbkAdd.Size = New System.Drawing.Size(26, 13)
        Me.lbkAdd.TabIndex = 2
        Me.lbkAdd.TabStop = True
        Me.lbkAdd.Text = "Add"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Customer"
        '
        'frmStartBalanceList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(692, 673)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lbkAdd)
        Me.Controls.Add(Me.dgStartBalance)
        Me.Controls.Add(Me.cboCustomers)
        Me.Name = "frmStartBalanceList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "LastStartBalanceList"
        CType(Me.dgStartBalance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboCustomers As System.Windows.Forms.ComboBox
    Friend WithEvents dgStartBalance As System.Windows.Forms.DataGridView
    Friend WithEvents lbkAdd As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
