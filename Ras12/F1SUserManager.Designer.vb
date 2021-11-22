<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F1SUserManager
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
        Me.components = New System.ComponentModel.Container()
        Me.CmbCustomer = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtMobi = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LblCreate = New System.Windows.Forms.LinkLabel()
        Me.GridAgent = New System.Windows.Forms.DataGridView()
        Me.LblDeactive = New System.Windows.Forms.LinkLabel()
        Me.LblResetPSW = New System.Windows.Forms.LinkLabel()
        Me.TxtEmail = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LblUpdate = New System.Windows.Forms.LinkLabel()
        Me.TxtWait = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TxtOwnerName = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TxtMinBLC = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        CType(Me.GridAgent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CmbCustomer
        '
        Me.CmbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCustomer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCustomer.FormattingEnabled = True
        Me.CmbCustomer.Location = New System.Drawing.Point(537, 14)
        Me.CmbCustomer.Name = "CmbCustomer"
        Me.CmbCustomer.Size = New System.Drawing.Size(124, 23)
        Me.CmbCustomer.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(458, 71)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Owner's Mobi"
        '
        'TxtMobi
        '
        Me.TxtMobi.Location = New System.Drawing.Point(537, 68)
        Me.TxtMobi.Name = "TxtMobi"
        Me.TxtMobi.Size = New System.Drawing.Size(124, 20)
        Me.TxtMobi.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(458, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(68, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Travel Agent"
        '
        'LblCreate
        '
        Me.LblCreate.AutoSize = True
        Me.LblCreate.Location = New System.Drawing.Point(563, 153)
        Me.LblCreate.Name = "LblCreate"
        Me.LblCreate.Size = New System.Drawing.Size(98, 13)
        Me.LblCreate.TabIndex = 4
        Me.LblCreate.TabStop = True
        Me.LblCreate.Text = "Create Master User"
        '
        'GridAgent
        '
        Me.GridAgent.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridAgent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridAgent.Location = New System.Drawing.Point(1, 1)
        Me.GridAgent.Name = "GridAgent"
        Me.GridAgent.RowHeadersVisible = False
        Me.GridAgent.Size = New System.Drawing.Size(454, 350)
        Me.GridAgent.TabIndex = 6
        '
        'LblDeactive
        '
        Me.LblDeactive.AutoSize = True
        Me.LblDeactive.Location = New System.Drawing.Point(602, 337)
        Me.LblDeactive.Name = "LblDeactive"
        Me.LblDeactive.Size = New System.Drawing.Size(59, 13)
        Me.LblDeactive.TabIndex = 7
        Me.LblDeactive.TabStop = True
        Me.LblDeactive.Text = "Deactivate"
        '
        'LblResetPSW
        '
        Me.LblResetPSW.AutoSize = True
        Me.LblResetPSW.Location = New System.Drawing.Point(458, 337)
        Me.LblResetPSW.Name = "LblResetPSW"
        Me.LblResetPSW.Size = New System.Drawing.Size(98, 13)
        Me.LblResetPSW.TabIndex = 8
        Me.LblResetPSW.TabStop = True
        Me.LblResetPSW.Text = "Reset Master PSW"
        '
        'TxtEmail
        '
        Me.TxtEmail.Location = New System.Drawing.Point(461, 130)
        Me.TxtEmail.Name = "TxtEmail"
        Me.TxtEmail.Size = New System.Drawing.Size(200, 20)
        Me.TxtEmail.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(458, 114)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Owner's email"
        '
        'LblUpdate
        '
        Me.LblUpdate.AutoSize = True
        Me.LblUpdate.Location = New System.Drawing.Point(458, 315)
        Me.LblUpdate.Name = "LblUpdate"
        Me.LblUpdate.Size = New System.Drawing.Size(82, 13)
        Me.LblUpdate.TabIndex = 10
        Me.LblUpdate.TabStop = True
        Me.LblUpdate.Text = "Update Contact"
        Me.LblUpdate.Visible = False
        '
        'TxtWait
        '
        Me.TxtWait.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtWait.Font = New System.Drawing.Font("Segoe Script", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtWait.ForeColor = System.Drawing.Color.DodgerBlue
        Me.TxtWait.Location = New System.Drawing.Point(135, 129)
        Me.TxtWait.Name = "TxtWait"
        Me.TxtWait.Size = New System.Drawing.Size(342, 24)
        Me.TxtWait.TabIndex = 11
        Me.TxtWait.Text = "Connecting to F1S Server...... Please Wait"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtMinBLC)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TxtOwnerName)
        Me.GroupBox1.Controls.Add(Me.GridAgent)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.LblUpdate)
        Me.GroupBox1.Controls.Add(Me.LblResetPSW)
        Me.GroupBox1.Controls.Add(Me.CmbCustomer)
        Me.GroupBox1.Controls.Add(Me.LblDeactive)
        Me.GroupBox1.Controls.Add(Me.TxtEmail)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtMobi)
        Me.GroupBox1.Controls.Add(Me.LblCreate)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(667, 356)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(458, 47)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Owner's Name"
        '
        'TxtOwnerName
        '
        Me.TxtOwnerName.Location = New System.Drawing.Point(537, 44)
        Me.TxtOwnerName.Name = "TxtOwnerName"
        Me.TxtOwnerName.Size = New System.Drawing.Size(123, 20)
        Me.TxtOwnerName.TabIndex = 1
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 128
        '
        'TxtMinBLC
        '
        Me.TxtMinBLC.Location = New System.Drawing.Point(537, 91)
        Me.TxtMinBLC.Name = "TxtMinBLC"
        Me.TxtMinBLC.Size = New System.Drawing.Size(124, 20)
        Me.TxtMinBLC.TabIndex = 13
        Me.TxtMinBLC.Text = "4000000"
        Me.TxtMinBLC.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(459, 94)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Min Balance"
        '
        'F1SUserManager
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(673, 359)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TxtWait)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "F1SUserManager"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: Sabre Simulator :. F1S User Manager"
        CType(Me.GridAgent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CmbCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtMobi As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblCreate As System.Windows.Forms.LinkLabel
    Friend WithEvents GridAgent As System.Windows.Forms.DataGridView
    Friend WithEvents LblDeactive As System.Windows.Forms.LinkLabel
    Friend WithEvents LblResetPSW As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtWait As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtOwnerName As System.Windows.Forms.TextBox
    Friend WithEvents TxtMinBLC As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
