<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStartBalanceEdit
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboCustomers = New System.Windows.Forms.ComboBox()
        Me.lbkSave = New System.Windows.Forms.LinkLabel()
        Me.dtpAsOf = New System.Windows.Forms.DateTimePicker()
        Me.txtVND_Avail = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Customer"
        '
        'cboCustomers
        '
        Me.cboCustomers.FormattingEnabled = True
        Me.cboCustomers.Location = New System.Drawing.Point(97, 13)
        Me.cboCustomers.Name = "cboCustomers"
        Me.cboCustomers.Size = New System.Drawing.Size(121, 21)
        Me.cboCustomers.TabIndex = 0
        '
        'lbkSave
        '
        Me.lbkSave.AutoSize = True
        Me.lbkSave.Location = New System.Drawing.Point(186, 140)
        Me.lbkSave.Name = "lbkSave"
        Me.lbkSave.Size = New System.Drawing.Size(32, 13)
        Me.lbkSave.TabIndex = 3
        Me.lbkSave.TabStop = True
        Me.lbkSave.Text = "Save"
        '
        'dtpAsOf
        '
        Me.dtpAsOf.CustomFormat = "dd MMM yy"
        Me.dtpAsOf.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAsOf.Location = New System.Drawing.Point(97, 50)
        Me.dtpAsOf.Name = "dtpAsOf"
        Me.dtpAsOf.Size = New System.Drawing.Size(121, 20)
        Me.dtpAsOf.TabIndex = 1
        '
        'txtVND_Avail
        '
        Me.txtVND_Avail.Location = New System.Drawing.Point(97, 88)
        Me.txtVND_Avail.Name = "txtVND_Avail"
        Me.txtVND_Avail.Size = New System.Drawing.Size(121, 20)
        Me.txtVND_Avail.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(33, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "As Of"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 95)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "VND_Avail"
        '
        'frmStartBalanceEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtVND_Avail)
        Me.Controls.Add(Me.dtpAsOf)
        Me.Controls.Add(Me.lbkSave)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboCustomers)
        Me.Name = "frmStartBalanceEdit"
        Me.Text = "StartBalanceAdd"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboCustomers As System.Windows.Forms.ComboBox
    Friend WithEvents lbkSave As System.Windows.Forms.LinkLabel
    Friend WithEvents dtpAsOf As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtVND_Avail As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
