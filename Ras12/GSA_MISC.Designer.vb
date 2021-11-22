<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GSA_MISC
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabFrontEnd = New System.Windows.Forms.TabPage()
        Me.OptAL = New System.Windows.Forms.RadioButton()
        Me.OptCust = New System.Windows.Forms.RadioButton()
        Me.TxtValidThru = New System.Windows.Forms.DateTimePicker()
        Me.GridFE = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtMaxLoss = New System.Windows.Forms.TextBox()
        Me.LblDeleteFrontEnd = New System.Windows.Forms.LinkLabel()
        Me.LblUpdateFrontEnd = New System.Windows.Forms.LinkLabel()
        Me.TxtAL = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CmbCustomer = New System.Windows.Forms.ComboBox()
        Me.TabControl1.SuspendLayout()
        Me.TabFrontEnd.SuspendLayout()
        CType(Me.GridFE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabFrontEnd)
        Me.TabControl1.Location = New System.Drawing.Point(0, 20)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(780, 477)
        Me.TabControl1.TabIndex = 0
        '
        'TabFrontEnd
        '
        Me.TabFrontEnd.Controls.Add(Me.OptAL)
        Me.TabFrontEnd.Controls.Add(Me.OptCust)
        Me.TabFrontEnd.Controls.Add(Me.TxtValidThru)
        Me.TabFrontEnd.Controls.Add(Me.GridFE)
        Me.TabFrontEnd.Controls.Add(Me.Label3)
        Me.TabFrontEnd.Controls.Add(Me.Label2)
        Me.TabFrontEnd.Controls.Add(Me.TxtMaxLoss)
        Me.TabFrontEnd.Controls.Add(Me.LblDeleteFrontEnd)
        Me.TabFrontEnd.Controls.Add(Me.LblUpdateFrontEnd)
        Me.TabFrontEnd.Controls.Add(Me.TxtAL)
        Me.TabFrontEnd.Controls.Add(Me.Label1)
        Me.TabFrontEnd.Location = New System.Drawing.Point(4, 22)
        Me.TabFrontEnd.Name = "TabFrontEnd"
        Me.TabFrontEnd.Padding = New System.Windows.Forms.Padding(3)
        Me.TabFrontEnd.Size = New System.Drawing.Size(772, 451)
        Me.TabFrontEnd.TabIndex = 0
        Me.TabFrontEnd.Text = "FrontEnd"
        Me.TabFrontEnd.UseVisualStyleBackColor = True
        '
        'OptAL
        '
        Me.OptAL.AutoSize = True
        Me.OptAL.Location = New System.Drawing.Point(66, 431)
        Me.OptAL.Name = "OptAL"
        Me.OptAL.Size = New System.Drawing.Size(50, 17)
        Me.OptAL.TabIndex = 7
        Me.OptAL.TabStop = True
        Me.OptAL.Text = "ByAL"
        Me.OptAL.UseVisualStyleBackColor = True
        '
        'OptCust
        '
        Me.OptCust.AutoSize = True
        Me.OptCust.Location = New System.Drawing.Point(2, 431)
        Me.OptCust.Name = "OptCust"
        Me.OptCust.Size = New System.Drawing.Size(58, 17)
        Me.OptCust.TabIndex = 7
        Me.OptCust.TabStop = True
        Me.OptCust.Text = "ByCust"
        Me.OptCust.UseVisualStyleBackColor = True
        '
        'TxtValidThru
        '
        Me.TxtValidThru.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtValidThru.Location = New System.Drawing.Point(247, 9)
        Me.TxtValidThru.Name = "TxtValidThru"
        Me.TxtValidThru.Size = New System.Drawing.Size(102, 20)
        Me.TxtValidThru.TabIndex = 6
        '
        'GridFE
        '
        Me.GridFE.AllowUserToAddRows = False
        Me.GridFE.AllowUserToDeleteRows = False
        Me.GridFE.BackgroundColor = System.Drawing.Color.Ivory
        Me.GridFE.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridFE.Location = New System.Drawing.Point(6, 30)
        Me.GridFE.Name = "GridFE"
        Me.GridFE.ReadOnly = True
        Me.GridFE.RowHeadersVisible = False
        Me.GridFE.Size = New System.Drawing.Size(391, 400)
        Me.GridFE.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(189, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(52, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "ValidThru"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(88, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "MaxLoss"
        '
        'TxtMaxLoss
        '
        Me.TxtMaxLoss.Location = New System.Drawing.Point(137, 9)
        Me.TxtMaxLoss.Name = "TxtMaxLoss"
        Me.TxtMaxLoss.Size = New System.Drawing.Size(46, 20)
        Me.TxtMaxLoss.TabIndex = 3
        Me.TxtMaxLoss.Text = "0.01"
        Me.TxtMaxLoss.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblDeleteFrontEnd
        '
        Me.LblDeleteFrontEnd.AutoSize = True
        Me.LblDeleteFrontEnd.Location = New System.Drawing.Point(359, 433)
        Me.LblDeleteFrontEnd.Name = "LblDeleteFrontEnd"
        Me.LblDeleteFrontEnd.Size = New System.Drawing.Size(38, 13)
        Me.LblDeleteFrontEnd.TabIndex = 2
        Me.LblDeleteFrontEnd.TabStop = True
        Me.LblDeleteFrontEnd.Text = "Delete"
        '
        'LblUpdateFrontEnd
        '
        Me.LblUpdateFrontEnd.AutoSize = True
        Me.LblUpdateFrontEnd.Location = New System.Drawing.Point(355, 12)
        Me.LblUpdateFrontEnd.Name = "LblUpdateFrontEnd"
        Me.LblUpdateFrontEnd.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdateFrontEnd.TabIndex = 2
        Me.LblUpdateFrontEnd.TabStop = True
        Me.LblUpdateFrontEnd.Text = "Update"
        '
        'TxtAL
        '
        Me.TxtAL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAL.Location = New System.Drawing.Point(47, 9)
        Me.TxtAL.MaxLength = 2
        Me.TxtAL.Name = "TxtAL"
        Me.TxtAL.Size = New System.Drawing.Size(38, 20)
        Me.TxtAL.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Airline"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(554, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(107, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Customer ShortName"
        '
        'CmbCustomer
        '
        Me.CmbCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCustomer.FormattingEnabled = True
        Me.CmbCustomer.Location = New System.Drawing.Point(668, 1)
        Me.CmbCustomer.Name = "CmbCustomer"
        Me.CmbCustomer.Size = New System.Drawing.Size(112, 21)
        Me.CmbCustomer.TabIndex = 4
        '
        'GSA_MISC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 497)
        Me.Controls.Add(Me.CmbCustomer)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "GSA_MISC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS 12 :. GSA MISC Setting"
        Me.TabControl1.ResumeLayout(False)
        Me.TabFrontEnd.ResumeLayout(False)
        Me.TabFrontEnd.PerformLayout()
        CType(Me.GridFE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabFrontEnd As System.Windows.Forms.TabPage
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtMaxLoss As System.Windows.Forms.TextBox
    Friend WithEvents LblUpdateFrontEnd As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtAL As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtValidThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GridFE As System.Windows.Forms.DataGridView
    Friend WithEvents OptAL As System.Windows.Forms.RadioButton
    Friend WithEvents OptCust As System.Windows.Forms.RadioButton
    Friend WithEvents LblDeleteFrontEnd As System.Windows.Forms.LinkLabel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CmbCustomer As System.Windows.Forms.ComboBox
End Class
