<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BSP_Agent
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
        Me.GridCode = New System.Windows.Forms.DataGridView()
        Me.ChkNewOnly = New System.Windows.Forms.CheckBox()
        Me.LblUpdate = New System.Windows.Forms.LinkLabel()
        Me.LblHAN = New System.Windows.Forms.LinkLabel()
        Me.CmbAgt = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblSGN = New System.Windows.Forms.LinkLabel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TxtFullName = New System.Windows.Forms.TextBox()
        Me.TxtAddress = New System.Windows.Forms.TextBox()
        Me.TxtTaxCode = New System.Windows.Forms.TextBox()
        Me.LblUpdateTax = New System.Windows.Forms.LinkLabel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        CType(Me.GridCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridCode
        '
        Me.GridCode.AllowUserToAddRows = False
        Me.GridCode.AllowUserToDeleteRows = False
        Me.GridCode.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridCode.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCode.Location = New System.Drawing.Point(1, 1)
        Me.GridCode.Name = "GridCode"
        Me.GridCode.RowHeadersVisible = False
        Me.GridCode.Size = New System.Drawing.Size(389, 467)
        Me.GridCode.TabIndex = 0
        '
        'ChkNewOnly
        '
        Me.ChkNewOnly.AutoSize = True
        Me.ChkNewOnly.Location = New System.Drawing.Point(1, 474)
        Me.ChkNewOnly.Name = "ChkNewOnly"
        Me.ChkNewOnly.Size = New System.Drawing.Size(69, 17)
        Me.ChkNewOnly.TabIndex = 1
        Me.ChkNewOnly.Text = "NewOnly"
        Me.ChkNewOnly.UseVisualStyleBackColor = True
        '
        'LblUpdate
        '
        Me.LblUpdate.AutoSize = True
        Me.LblUpdate.Location = New System.Drawing.Point(188, 9)
        Me.LblUpdate.Name = "LblUpdate"
        Me.LblUpdate.Size = New System.Drawing.Size(90, 13)
        Me.LblUpdate.TabIndex = 2
        Me.LblUpdate.TabStop = True
        Me.LblUpdate.Tag = "SGN"
        Me.LblUpdate.Text = "Update TV Agent"
        '
        'LblHAN
        '
        Me.LblHAN.AutoSize = True
        Me.LblHAN.Location = New System.Drawing.Point(121, 424)
        Me.LblHAN.Name = "LblHAN"
        Me.LblHAN.Size = New System.Drawing.Size(61, 13)
        Me.LblHAN.TabIndex = 2
        Me.LblHAN.TabStop = True
        Me.LblHAN.Tag = "HAN"
        Me.LblHAN.Text = "HAN Agent"
        '
        'CmbAgt
        '
        Me.CmbAgt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAgt.FormattingEnabled = True
        Me.CmbAgt.Location = New System.Drawing.Point(46, 6)
        Me.CmbAgt.Name = "CmbAgt"
        Me.CmbAgt.Size = New System.Drawing.Size(136, 21)
        Me.CmbAgt.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Agent"
        '
        'LblSGN
        '
        Me.LblSGN.AutoSize = True
        Me.LblSGN.Location = New System.Drawing.Point(6, 424)
        Me.LblSGN.Name = "LblSGN"
        Me.LblSGN.Size = New System.Drawing.Size(98, 13)
        Me.LblSGN.TabIndex = 5
        Me.LblSGN.TabStop = True
        Me.LblSGN.Text = "SGN NonTV Agent"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(392, 1)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(389, 466)
        Me.TabControl1.TabIndex = 6
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.LblSGN)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.LblHAN)
        Me.TabPage1.Controls.Add(Me.CmbAgt)
        Me.TabPage1.Controls.Add(Me.LblUpdate)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(381, 440)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Sales"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.LblUpdateTax)
        Me.TabPage2.Controls.Add(Me.TxtTaxCode)
        Me.TabPage2.Controls.Add(Me.TxtAddress)
        Me.TabPage2.Controls.Add(Me.TxtFullName)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(381, 440)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Accounting"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TxtFullName
        '
        Me.TxtFullName.Location = New System.Drawing.Point(3, 30)
        Me.TxtFullName.Name = "TxtFullName"
        Me.TxtFullName.Size = New System.Drawing.Size(375, 20)
        Me.TxtFullName.TabIndex = 0
        '
        'TxtAddress
        '
        Me.TxtAddress.Location = New System.Drawing.Point(3, 69)
        Me.TxtAddress.Name = "TxtAddress"
        Me.TxtAddress.Size = New System.Drawing.Size(375, 20)
        Me.TxtAddress.TabIndex = 0
        '
        'TxtTaxCode
        '
        Me.TxtTaxCode.Location = New System.Drawing.Point(3, 111)
        Me.TxtTaxCode.Name = "TxtTaxCode"
        Me.TxtTaxCode.Size = New System.Drawing.Size(137, 20)
        Me.TxtTaxCode.TabIndex = 0
        '
        'LblUpdateTax
        '
        Me.LblUpdateTax.AutoSize = True
        Me.LblUpdateTax.Location = New System.Drawing.Point(146, 114)
        Me.LblUpdateTax.Name = "LblUpdateTax"
        Me.LblUpdateTax.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdateTax.TabIndex = 1
        Me.LblUpdateTax.TabStop = True
        Me.LblUpdateTax.Text = "Update"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(0, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Full Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(0, 53)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Address"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(0, 92)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 13)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "TaxCode"
        '
        'BSP_Agent
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.ChkNewOnly)
        Me.Controls.Add(Me.GridCode)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BSP_Agent"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS 12 :. BSP Agent"
        CType(Me.GridCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridCode As System.Windows.Forms.DataGridView
    Friend WithEvents ChkNewOnly As System.Windows.Forms.CheckBox
    Friend WithEvents LblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents LblHAN As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbAgt As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblSGN As System.Windows.Forms.LinkLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblUpdateTax As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtTaxCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtAddress As System.Windows.Forms.TextBox
    Friend WithEvents TxtFullName As System.Windows.Forms.TextBox
End Class
