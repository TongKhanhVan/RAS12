<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UpdateForEx
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GridROE = New System.Windows.Forms.DataGridView()
        Me.CmbCurr = New System.Windows.Forms.ComboBox()
        Me.txtBSR = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CmdAdd = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtBBR = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtSurcharge = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtApplySCto = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtApplyROEto = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtEffectDate = New System.Windows.Forms.DateTimePicker()
        Me.txtFilterValue = New System.Windows.Forms.TextBox()
        Me.CmdShow = New System.Windows.Forms.Button()
        Me.CmdFilter = New System.Windows.Forms.Button()
        Me.CmdMark = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TxtApplyROE_COS = New System.Windows.Forms.TextBox()
        CType(Me.GridROE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridROE
        '
        Me.GridROE.AllowUserToAddRows = False
        Me.GridROE.AllowUserToDeleteRows = False
        Me.GridROE.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridROE.BackgroundColor = System.Drawing.Color.LemonChiffon
        Me.GridROE.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridROE.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.GridROE.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.GridROE.DefaultCellStyle = DataGridViewCellStyle6
        Me.GridROE.Location = New System.Drawing.Point(2, 79)
        Me.GridROE.Name = "GridROE"
        Me.GridROE.Size = New System.Drawing.Size(1012, 394)
        Me.GridROE.TabIndex = 4
        '
        'CmbCurr
        '
        Me.CmbCurr.FormattingEnabled = True
        Me.CmbCurr.Location = New System.Drawing.Point(54, 6)
        Me.CmbCurr.Name = "CmbCurr"
        Me.CmbCurr.Size = New System.Drawing.Size(52, 21)
        Me.CmbCurr.TabIndex = 0
        '
        'txtBSR
        '
        Me.txtBSR.Location = New System.Drawing.Point(141, 6)
        Me.txtBSR.Name = "txtBSR"
        Me.txtBSR.Size = New System.Drawing.Size(76, 20)
        Me.txtBSR.TabIndex = 1
        Me.txtBSR.Text = "0"
        Me.txtBSR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(-1, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Currency"
        '
        'CmdAdd
        '
        Me.CmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdAdd.Location = New System.Drawing.Point(943, 53)
        Me.CmdAdd.Name = "CmdAdd"
        Me.CmdAdd.Size = New System.Drawing.Size(71, 22)
        Me.CmdAdd.TabIndex = 7
        Me.CmdAdd.Text = "Add"
        Me.CmdAdd.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(109, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "BSR"
        '
        'txtBBR
        '
        Me.txtBBR.Location = New System.Drawing.Point(256, 6)
        Me.txtBBR.Name = "txtBBR"
        Me.txtBBR.Size = New System.Drawing.Size(76, 20)
        Me.txtBBR.TabIndex = 2
        Me.txtBBR.Text = "0"
        Me.txtBBR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(222, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "BBR"
        '
        'txtSurcharge
        '
        Me.txtSurcharge.Location = New System.Drawing.Point(54, 53)
        Me.txtSurcharge.Name = "txtSurcharge"
        Me.txtSurcharge.Size = New System.Drawing.Size(52, 20)
        Me.txtSurcharge.TabIndex = 4
        Me.txtSurcharge.Text = "0"
        Me.txtSurcharge.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(-1, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "SurCharge"
        '
        'txtApplySCto
        '
        Me.txtApplySCto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtApplySCto.Location = New System.Drawing.Point(256, 53)
        Me.txtApplySCto.Name = "txtApplySCto"
        Me.txtApplySCto.Size = New System.Drawing.Size(76, 20)
        Me.txtApplySCto.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(185, 55)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(66, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Apply SC To"
        '
        'txtApplyROEto
        '
        Me.txtApplyROEto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtApplyROEto.Location = New System.Drawing.Point(370, 7)
        Me.txtApplyROEto.Name = "txtApplyROEto"
        Me.txtApplyROEto.Size = New System.Drawing.Size(644, 20)
        Me.txtApplyROEto.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(333, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Apply2"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(-1, 34)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 13)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Effective"
        '
        'txtEffectDate
        '
        Me.txtEffectDate.CustomFormat = "dd-MMM-yy HH:mm"
        Me.txtEffectDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtEffectDate.Location = New System.Drawing.Point(54, 30)
        Me.txtEffectDate.Name = "txtEffectDate"
        Me.txtEffectDate.Size = New System.Drawing.Size(163, 20)
        Me.txtEffectDate.TabIndex = 6
        '
        'txtFilterValue
        '
        Me.txtFilterValue.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtFilterValue.Location = New System.Drawing.Point(2, 479)
        Me.txtFilterValue.Name = "txtFilterValue"
        Me.txtFilterValue.Size = New System.Drawing.Size(54, 20)
        Me.txtFilterValue.TabIndex = 10
        '
        'CmdShow
        '
        Me.CmdShow.Location = New System.Drawing.Point(134, 479)
        Me.CmdShow.Name = "CmdShow"
        Me.CmdShow.Size = New System.Drawing.Size(83, 22)
        Me.CmdShow.TabIndex = 8
        Me.CmdShow.Text = "ShowAll"
        Me.CmdShow.UseVisualStyleBackColor = True
        '
        'CmdFilter
        '
        Me.CmdFilter.Location = New System.Drawing.Point(70, 479)
        Me.CmdFilter.Name = "CmdFilter"
        Me.CmdFilter.Size = New System.Drawing.Size(58, 22)
        Me.CmdFilter.TabIndex = 8
        Me.CmdFilter.Text = "Filter"
        Me.CmdFilter.UseVisualStyleBackColor = True
        '
        'CmdMark
        '
        Me.CmdMark.Location = New System.Drawing.Point(716, 477)
        Me.CmdMark.Name = "CmdMark"
        Me.CmdMark.Size = New System.Drawing.Size(71, 22)
        Me.CmdMark.TabIndex = 8
        Me.CmdMark.Text = "Delete"
        Me.CmdMark.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(300, 34)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(70, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Apply2 (COS)"
        '
        'TxtApplyROE_COS
        '
        Me.TxtApplyROE_COS.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtApplyROE_COS.Location = New System.Drawing.Point(370, 31)
        Me.TxtApplyROE_COS.Name = "TxtApplyROE_COS"
        Me.TxtApplyROE_COS.Size = New System.Drawing.Size(644, 20)
        Me.TxtApplyROE_COS.TabIndex = 3
        '
        'UpdateForEx
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 501)
        Me.Controls.Add(Me.txtFilterValue)
        Me.Controls.Add(Me.txtEffectDate)
        Me.Controls.Add(Me.CmdMark)
        Me.Controls.Add(Me.CmdFilter)
        Me.Controls.Add(Me.CmdShow)
        Me.Controls.Add(Me.CmdAdd)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtApplyROE_COS)
        Me.Controls.Add(Me.txtApplyROEto)
        Me.Controls.Add(Me.txtApplySCto)
        Me.Controls.Add(Me.txtSurcharge)
        Me.Controls.Add(Me.txtBBR)
        Me.Controls.Add(Me.txtBSR)
        Me.Controls.Add(Me.CmbCurr)
        Me.Controls.Add(Me.GridROE)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "UpdateForEx"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. UpdateForEx"
        CType(Me.GridROE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridROE As System.Windows.Forms.DataGridView
    Friend WithEvents CmbCurr As System.Windows.Forms.ComboBox
    Friend WithEvents txtBSR As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CmdAdd As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtBBR As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSurcharge As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtApplySCto As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtApplyROEto As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtEffectDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtFilterValue As System.Windows.Forms.TextBox
    Friend WithEvents CmdShow As System.Windows.Forms.Button
    Friend WithEvents CmdFilter As System.Windows.Forms.Button
    Friend WithEvents CmdMark As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtApplyROE_COS As System.Windows.Forms.TextBox
End Class
