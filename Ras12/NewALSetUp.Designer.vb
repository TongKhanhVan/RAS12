<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NewALSetUp
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
        Me.TxtAL = New System.Windows.Forms.TextBox()
        Me.TxtDocCode = New System.Windows.Forms.TextBox()
        Me.TxtALName = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ChkBSP = New System.Windows.Forms.CheckBox()
        Me.OptNonGSA = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OptGSA = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TxtMS = New System.Windows.Forms.TextBox()
        Me.TxtKH = New System.Windows.Forms.TextBox()
        Me.CmbCurr = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.LblAdd = New System.Windows.Forms.LinkLabel()
        Me.CmbTVCompany = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.LblSave = New System.Windows.Forms.LinkLabel()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtAL
        '
        Me.TxtAL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAL.Location = New System.Drawing.Point(40, 6)
        Me.TxtAL.MaxLength = 2
        Me.TxtAL.Name = "TxtAL"
        Me.TxtAL.Size = New System.Drawing.Size(28, 20)
        Me.TxtAL.TabIndex = 0
        '
        'TxtDocCode
        '
        Me.TxtDocCode.Enabled = False
        Me.TxtDocCode.Location = New System.Drawing.Point(322, 6)
        Me.TxtDocCode.Name = "TxtDocCode"
        Me.TxtDocCode.ReadOnly = True
        Me.TxtDocCode.Size = New System.Drawing.Size(39, 20)
        Me.TxtDocCode.TabIndex = 1
        '
        'TxtALName
        '
        Me.TxtALName.Enabled = False
        Me.TxtALName.Location = New System.Drawing.Point(70, 6)
        Me.TxtALName.Name = "TxtALName"
        Me.TxtALName.ReadOnly = True
        Me.TxtALName.Size = New System.Drawing.Size(197, 20)
        Me.TxtALName.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ChkBSP)
        Me.GroupBox1.Controls.Add(Me.OptNonGSA)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.OptGSA)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TxtMS)
        Me.GroupBox1.Controls.Add(Me.TxtKH)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(317, 36)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'ChkBSP
        '
        Me.ChkBSP.AutoSize = True
        Me.ChkBSP.Location = New System.Drawing.Point(267, 13)
        Me.ChkBSP.Name = "ChkBSP"
        Me.ChkBSP.Size = New System.Drawing.Size(47, 17)
        Me.ChkBSP.TabIndex = 6
        Me.ChkBSP.Text = "BSP"
        Me.ChkBSP.UseVisualStyleBackColor = True
        '
        'OptNonGSA
        '
        Me.OptNonGSA.AutoSize = True
        Me.OptNonGSA.Checked = True
        Me.OptNonGSA.Location = New System.Drawing.Point(194, 12)
        Me.OptNonGSA.Name = "OptNonGSA"
        Me.OptNonGSA.Size = New System.Drawing.Size(67, 17)
        Me.OptNonGSA.TabIndex = 5
        Me.OptNonGSA.TabStop = True
        Me.OptNonGSA.Text = "NonGSA"
        Me.OptNonGSA.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(123, 14)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(22, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "KH"
        '
        'OptGSA
        '
        Me.OptGSA.AutoSize = True
        Me.OptGSA.Location = New System.Drawing.Point(6, 12)
        Me.OptGSA.Name = "OptGSA"
        Me.OptGSA.Size = New System.Drawing.Size(47, 17)
        Me.OptGSA.TabIndex = 2
        Me.OptGSA.Text = "GSA"
        Me.OptGSA.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(61, 14)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(23, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "MS"
        '
        'TxtMS
        '
        Me.TxtMS.Location = New System.Drawing.Point(87, 11)
        Me.TxtMS.Name = "TxtMS"
        Me.TxtMS.Size = New System.Drawing.Size(35, 20)
        Me.TxtMS.TabIndex = 3
        '
        'TxtKH
        '
        Me.TxtKH.Location = New System.Drawing.Point(146, 11)
        Me.TxtKH.Name = "TxtKH"
        Me.TxtKH.Size = New System.Drawing.Size(34, 20)
        Me.TxtKH.TabIndex = 4
        '
        'CmbCurr
        '
        Me.CmbCurr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCurr.FormattingEnabled = True
        Me.CmbCurr.Items.AddRange(New Object() {"USD", "VND"})
        Me.CmbCurr.Location = New System.Drawing.Point(418, 6)
        Me.CmbCurr.Name = "CmbCurr"
        Me.CmbCurr.Size = New System.Drawing.Size(58, 21)
        Me.CmbCurr.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Airline"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(268, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "DocCode"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(367, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Currency"
        '
        'LblAdd
        '
        Me.LblAdd.AutoSize = True
        Me.LblAdd.Enabled = False
        Me.LblAdd.Location = New System.Drawing.Point(434, 88)
        Me.LblAdd.Name = "LblAdd"
        Me.LblAdd.Size = New System.Drawing.Size(26, 13)
        Me.LblAdd.TabIndex = 8
        Me.LblAdd.TabStop = True
        Me.LblAdd.Text = "Add"
        '
        'CmbTVCompany
        '
        Me.CmbTVCompany.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbTVCompany.FormattingEnabled = True
        Me.CmbTVCompany.Items.AddRange(New Object() {"GDS", "TVP", "TVT", "VLM"})
        Me.CmbTVCompany.Location = New System.Drawing.Point(418, 43)
        Me.CmbTVCompany.Name = "CmbTVCompany"
        Me.CmbTVCompany.Size = New System.Drawing.Size(58, 21)
        Me.CmbTVCompany.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(348, 46)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(68, 13)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "TV Company"
        '
        'LblSave
        '
        Me.LblSave.AutoSize = True
        Me.LblSave.Enabled = False
        Me.LblSave.Location = New System.Drawing.Point(348, 88)
        Me.LblSave.Name = "LblSave"
        Me.LblSave.Size = New System.Drawing.Size(32, 13)
        Me.LblSave.TabIndex = 8
        Me.LblSave.TabStop = True
        Me.LblSave.Text = "Save"
        '
        'NewALSetUp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(480, 110)
        Me.Controls.Add(Me.CmbTVCompany)
        Me.Controls.Add(Me.LblSave)
        Me.Controls.Add(Me.LblAdd)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmbCurr)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TxtALName)
        Me.Controls.Add(Me.TxtDocCode)
        Me.Controls.Add(Me.TxtAL)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "NewALSetUp"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. New AL Setup"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtAL As System.Windows.Forms.TextBox
    Friend WithEvents TxtDocCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtALName As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents OptGSA As System.Windows.Forms.RadioButton
    Friend WithEvents CmbCurr As System.Windows.Forms.ComboBox
    Friend WithEvents TxtMS As System.Windows.Forms.TextBox
    Friend WithEvents TxtKH As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OptNonGSA As System.Windows.Forms.RadioButton
    Friend WithEvents ChkBSP As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents LblAdd As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbTVCompany As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LblSave As System.Windows.Forms.LinkLabel
End Class
