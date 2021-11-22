<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMassExtendServiceFee
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
        Me.lbkExtend = New System.Windows.Forms.LinkLabel()
        Me.GridExpireDate = New System.Windows.Forms.DataGridView()
        Me.dtpNewValidThru = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.GridExpireDate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbkExtend
        '
        Me.lbkExtend.AutoSize = True
        Me.lbkExtend.Location = New System.Drawing.Point(186, 451)
        Me.lbkExtend.Name = "lbkExtend"
        Me.lbkExtend.Size = New System.Drawing.Size(40, 13)
        Me.lbkExtend.TabIndex = 0
        Me.lbkExtend.TabStop = True
        Me.lbkExtend.Text = "Extend"
        '
        'GridExpireDate
        '
        Me.GridExpireDate.AllowUserToAddRows = False
        Me.GridExpireDate.AllowUserToDeleteRows = False
        Me.GridExpireDate.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridExpireDate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridExpireDate.Location = New System.Drawing.Point(3, 12)
        Me.GridExpireDate.Name = "GridExpireDate"
        Me.GridExpireDate.RowHeadersVisible = False
        Me.GridExpireDate.Size = New System.Drawing.Size(488, 423)
        Me.GridExpireDate.TabIndex = 22
        '
        'dtpNewValidThru
        '
        Me.dtpNewValidThru.CustomFormat = "dd MMM yy"
        Me.dtpNewValidThru.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpNewValidThru.Location = New System.Drawing.Point(92, 445)
        Me.dtpNewValidThru.Name = "dtpNewValidThru"
        Me.dtpNewValidThru.Size = New System.Drawing.Size(88, 20)
        Me.dtpNewValidThru.TabIndex = 23
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 451)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "NewValidThru"
        '
        'frmMassExtendServiceFee
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 473)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpNewValidThru)
        Me.Controls.Add(Me.GridExpireDate)
        Me.Controls.Add(Me.lbkExtend)
        Me.Name = "frmMassExtendServiceFee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MassExtendServiceFee"
        CType(Me.GridExpireDate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbkExtend As System.Windows.Forms.LinkLabel
    Friend WithEvents GridExpireDate As System.Windows.Forms.DataGridView
    Friend WithEvents dtpNewValidThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
