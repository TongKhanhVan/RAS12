<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class unReportedTickets
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
        Me.GridUnRptTix = New System.Windows.Forms.DataGridView()
        Me.Opt1S = New System.Windows.Forms.RadioButton()
        Me.Opt1A = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CmbLstNdays = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.GridUnRptTix, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridUnRptTix
        '
        Me.GridUnRptTix.AllowUserToAddRows = False
        Me.GridUnRptTix.AllowUserToDeleteRows = False
        Me.GridUnRptTix.BackgroundColor = System.Drawing.Color.LightCyan
        Me.GridUnRptTix.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridUnRptTix.Location = New System.Drawing.Point(4, 26)
        Me.GridUnRptTix.Name = "GridUnRptTix"
        Me.GridUnRptTix.RowHeadersVisible = False
        Me.GridUnRptTix.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridUnRptTix.Size = New System.Drawing.Size(773, 470)
        Me.GridUnRptTix.TabIndex = 0
        '
        'Opt1S
        '
        Me.Opt1S.AutoSize = True
        Me.Opt1S.Location = New System.Drawing.Point(12, 3)
        Me.Opt1S.Name = "Opt1S"
        Me.Opt1S.Size = New System.Drawing.Size(38, 17)
        Me.Opt1S.TabIndex = 1
        Me.Opt1S.TabStop = True
        Me.Opt1S.Text = "1S"
        Me.Opt1S.UseVisualStyleBackColor = True
        '
        'Opt1A
        '
        Me.Opt1A.AutoSize = True
        Me.Opt1A.Location = New System.Drawing.Point(56, 3)
        Me.Opt1A.Name = "Opt1A"
        Me.Opt1A.Size = New System.Drawing.Size(38, 17)
        Me.Opt1A.TabIndex = 1
        Me.Opt1A.TabStop = True
        Me.Opt1A.Text = "1A"
        Me.Opt1A.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(176, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "days"
        '
        'CmbLstNdays
        '
        Me.CmbLstNdays.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbLstNdays.FormattingEnabled = True
        Me.CmbLstNdays.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21"})
        Me.CmbLstNdays.Location = New System.Drawing.Point(144, 2)
        Me.CmbLstNdays.Name = "CmbLstNdays"
        Me.CmbLstNdays.Size = New System.Drawing.Size(33, 21)
        Me.CmbLstNdays.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(111, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(27, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Last"
        '
        'unReportedTickets
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.CmbLstNdays)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Opt1A)
        Me.Controls.Add(Me.Opt1S)
        Me.Controls.Add(Me.GridUnRptTix)
        Me.Controls.Add(Me.Label1)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "unReportedTickets"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS 12 :. unReported Tickets"
        CType(Me.GridUnRptTix, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridUnRptTix As System.Windows.Forms.DataGridView
    Friend WithEvents Opt1S As System.Windows.Forms.RadioButton
    Friend WithEvents Opt1A As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CmbLstNdays As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
