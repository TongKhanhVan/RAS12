<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PivotUser
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
        Me.CmbUser = New System.Windows.Forms.ComboBox
        Me.LstAgent = New System.Windows.Forms.CheckedListBox
        Me.LblUpdate = New System.Windows.Forms.LinkLabel
        Me.Label1 = New System.Windows.Forms.Label
        Me.LblCheckAll = New System.Windows.Forms.LinkLabel
        Me.LblUnCheckAll = New System.Windows.Forms.LinkLabel
        Me.OptTA = New System.Windows.Forms.RadioButton
        Me.OptTO = New System.Windows.Forms.RadioButton
        Me.OptCA = New System.Windows.Forms.RadioButton
        Me.ChkUnAssign = New System.Windows.Forms.CheckBox
        Me.OptCS = New System.Windows.Forms.RadioButton
        Me.SuspendLayout()
        '
        'CmbUser
        '
        Me.CmbUser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbUser.FormattingEnabled = True
        Me.CmbUser.Location = New System.Drawing.Point(50, 2)
        Me.CmbUser.Name = "CmbUser"
        Me.CmbUser.Size = New System.Drawing.Size(148, 21)
        Me.CmbUser.TabIndex = 0
        '
        'LstAgent
        '
        Me.LstAgent.CheckOnClick = True
        Me.LstAgent.ColumnWidth = 75
        Me.LstAgent.FormattingEnabled = True
        Me.LstAgent.Location = New System.Drawing.Point(1, 24)
        Me.LstAgent.MultiColumn = True
        Me.LstAgent.Name = "LstAgent"
        Me.LstAgent.Size = New System.Drawing.Size(890, 544)
        Me.LstAgent.TabIndex = 1
        '
        'LblUpdate
        '
        Me.LblUpdate.AutoSize = True
        Me.LblUpdate.Enabled = False
        Me.LblUpdate.Location = New System.Drawing.Point(204, 5)
        Me.LblUpdate.Name = "LblUpdate"
        Me.LblUpdate.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdate.TabIndex = 2
        Me.LblUpdate.TabStop = True
        Me.LblUpdate.Text = "Update"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "User"
        '
        'LblCheckAll
        '
        Me.LblCheckAll.AutoSize = True
        Me.LblCheckAll.Location = New System.Drawing.Point(724, 5)
        Me.LblCheckAll.Name = "LblCheckAll"
        Me.LblCheckAll.Size = New System.Drawing.Size(49, 13)
        Me.LblCheckAll.TabIndex = 4
        Me.LblCheckAll.TabStop = True
        Me.LblCheckAll.Text = "CheckAll"
        '
        'LblUnCheckAll
        '
        Me.LblUnCheckAll.AutoSize = True
        Me.LblUnCheckAll.Location = New System.Drawing.Point(800, 5)
        Me.LblUnCheckAll.Name = "LblUnCheckAll"
        Me.LblUnCheckAll.Size = New System.Drawing.Size(63, 13)
        Me.LblUnCheckAll.TabIndex = 4
        Me.LblUnCheckAll.TabStop = True
        Me.LblUnCheckAll.Text = "UnCheckAll"
        '
        'OptTA
        '
        Me.OptTA.AutoSize = True
        Me.OptTA.Checked = True
        Me.OptTA.Location = New System.Drawing.Point(298, 3)
        Me.OptTA.Name = "OptTA"
        Me.OptTA.Size = New System.Drawing.Size(39, 17)
        Me.OptTA.TabIndex = 5
        Me.OptTA.TabStop = True
        Me.OptTA.Text = "TA"
        Me.OptTA.UseVisualStyleBackColor = True
        '
        'OptTO
        '
        Me.OptTO.AutoSize = True
        Me.OptTO.Location = New System.Drawing.Point(343, 3)
        Me.OptTO.Name = "OptTO"
        Me.OptTO.Size = New System.Drawing.Size(40, 17)
        Me.OptTO.TabIndex = 5
        Me.OptTO.Text = "TO"
        Me.OptTO.UseVisualStyleBackColor = True
        '
        'OptCA
        '
        Me.OptCA.AutoSize = True
        Me.OptCA.Location = New System.Drawing.Point(389, 3)
        Me.OptCA.Name = "OptCA"
        Me.OptCA.Size = New System.Drawing.Size(39, 17)
        Me.OptCA.TabIndex = 5
        Me.OptCA.Text = "CA"
        Me.OptCA.UseVisualStyleBackColor = True
        '
        'ChkUnAssign
        '
        Me.ChkUnAssign.AutoSize = True
        Me.ChkUnAssign.Location = New System.Drawing.Point(515, 4)
        Me.ChkUnAssign.Name = "ChkUnAssign"
        Me.ChkUnAssign.Size = New System.Drawing.Size(104, 17)
        Me.ChkUnAssign.TabIndex = 6
        Me.ChkUnAssign.Text = "UnAssignedOnly"
        Me.ChkUnAssign.UseVisualStyleBackColor = True
        '
        'OptCS
        '
        Me.OptCS.AutoSize = True
        Me.OptCS.Location = New System.Drawing.Point(434, 3)
        Me.OptCS.Name = "OptCS"
        Me.OptCS.Size = New System.Drawing.Size(39, 17)
        Me.OptCS.TabIndex = 5
        Me.OptCS.Text = "CS"
        Me.OptCS.UseVisualStyleBackColor = True
        '
        'PivotUser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(892, 566)
        Me.Controls.Add(Me.ChkUnAssign)
        Me.Controls.Add(Me.OptCS)
        Me.Controls.Add(Me.OptCA)
        Me.Controls.Add(Me.OptTO)
        Me.Controls.Add(Me.OptTA)
        Me.Controls.Add(Me.LblUnCheckAll)
        Me.Controls.Add(Me.LblCheckAll)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LblUpdate)
        Me.Controls.Add(Me.LstAgent)
        Me.Controls.Add(Me.CmbUser)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PivotUser"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Pivot User Manager"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CmbUser As System.Windows.Forms.ComboBox
    Friend WithEvents LstAgent As System.Windows.Forms.CheckedListBox
    Friend WithEvents LblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUnCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents OptTA As System.Windows.Forms.RadioButton
    Friend WithEvents OptTO As System.Windows.Forms.RadioButton
    Friend WithEvents OptCA As System.Windows.Forms.RadioButton
    Friend WithEvents ChkUnAssign As System.Windows.Forms.CheckBox
    Friend WithEvents OptCS As System.Windows.Forms.RadioButton
End Class
