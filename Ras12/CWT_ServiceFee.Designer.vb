<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CWT_ServiceFee
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
        Me.CmbCust = New System.Windows.Forms.ComboBox()
        Me.CmbService = New System.Windows.Forms.ComboBox()
        Me.CmbCurr = New System.Windows.Forms.ComboBox()
        Me.CmbTRX = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblUpdate = New System.Windows.Forms.LinkLabel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GridSF = New System.Windows.Forms.DataGridView()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.LstReg = New System.Windows.Forms.CheckedListBox()
        Me.LblCheckAll = New System.Windows.Forms.LinkLabel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TxtAmt = New System.Windows.Forms.TextBox()
        Me.ChkVAT = New System.Windows.Forms.CheckBox()
        Me.chkWzAIR = New System.Windows.Forms.CheckBox()
        Me.TxtValidFrom = New System.Windows.Forms.DateTimePicker()
        Me.TxtValidThru = New System.Windows.Forms.DateTimePicker()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.label15 = New System.Windows.Forms.Label()
        Me.CmbBase = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.CmbTKT = New System.Windows.Forms.ComboBox()
        Me.CmbAL = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.CmbRtg = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CmbCabin = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.TxtFltTime = New System.Windows.Forms.TextBox()
        Me.cmbSegTKT = New System.Windows.Forms.ComboBox()
        Me.TxtPersonal = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.CmbOWRT = New System.Windows.Forms.ComboBox()
        CType(Me.GridSF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CmbCust
        '
        Me.CmbCust.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCust.FormattingEnabled = True
        Me.CmbCust.Location = New System.Drawing.Point(60, 12)
        Me.CmbCust.Name = "CmbCust"
        Me.CmbCust.Size = New System.Drawing.Size(115, 21)
        Me.CmbCust.TabIndex = 0
        '
        'CmbService
        '
        Me.CmbService.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbService.FormattingEnabled = True
        Me.CmbService.Items.AddRange(New Object() {"AIR", "HTL", "CAR", "INS", "VSA", "AHC"})
        Me.CmbService.Location = New System.Drawing.Point(223, 12)
        Me.CmbService.Name = "CmbService"
        Me.CmbService.Size = New System.Drawing.Size(70, 21)
        Me.CmbService.TabIndex = 1
        '
        'CmbCurr
        '
        Me.CmbCurr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCurr.FormattingEnabled = True
        Me.CmbCurr.Items.AddRange(New Object() {"USD", "VND", "PCT"})
        Me.CmbCurr.Location = New System.Drawing.Point(416, 12)
        Me.CmbCurr.Name = "CmbCurr"
        Me.CmbCurr.Size = New System.Drawing.Size(51, 21)
        Me.CmbCurr.TabIndex = 4
        '
        'CmbTRX
        '
        Me.CmbTRX.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbTRX.FormattingEnabled = True
        Me.CmbTRX.Items.AddRange(New Object() {"S", "R", "V"})
        Me.CmbTRX.Location = New System.Drawing.Point(322, 12)
        Me.CmbTRX.Name = "CmbTRX"
        Me.CmbTRX.Size = New System.Drawing.Size(32, 21)
        Me.CmbTRX.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Customer"
        '
        'LblUpdate
        '
        Me.LblUpdate.AutoSize = True
        Me.LblUpdate.Location = New System.Drawing.Point(735, 92)
        Me.LblUpdate.Name = "LblUpdate"
        Me.LblUpdate.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdate.TabIndex = 13
        Me.LblUpdate.TabStop = True
        Me.LblUpdate.Text = "Update"
        Me.LblUpdate.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(178, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Service"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(293, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(29, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "TRX"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(390, 15)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(26, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Curr"
        '
        'GridSF
        '
        Me.GridSF.AllowUserToAddRows = False
        Me.GridSF.AllowUserToDeleteRows = False
        Me.GridSF.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridSF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridSF.Location = New System.Drawing.Point(2, 112)
        Me.GridSF.Name = "GridSF"
        Me.GridSF.RowHeadersVisible = False
        Me.GridSF.Size = New System.Drawing.Size(775, 363)
        Me.GridSF.TabIndex = 7
        '
        'LblDelete
        '
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Location = New System.Drawing.Point(-1, 478)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 8
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        Me.LblDelete.Visible = False
        '
        'LstReg
        '
        Me.LstReg.CheckOnClick = True
        Me.LstReg.ColumnWidth = 40
        Me.LstReg.FormattingEnabled = True
        Me.LstReg.Location = New System.Drawing.Point(60, 57)
        Me.LstReg.MultiColumn = True
        Me.LstReg.Name = "LstReg"
        Me.LstReg.Size = New System.Drawing.Size(649, 49)
        Me.LstReg.TabIndex = 7
        '
        'LblCheckAll
        '
        Me.LblCheckAll.AutoSize = True
        Me.LblCheckAll.Location = New System.Drawing.Point(0, 92)
        Me.LblCheckAll.Name = "LblCheckAll"
        Me.LblCheckAll.Size = New System.Drawing.Size(50, 13)
        Me.LblCheckAll.TabIndex = 10
        Me.LblCheckAll.TabStop = True
        Me.LblCheckAll.Text = "Tick ALL"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 57)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 13)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Countries"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(2, 73)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 13)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "(Regional)"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(549, 15)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 13)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "Amount"
        '
        'TxtAmt
        '
        Me.TxtAmt.Location = New System.Drawing.Point(592, 12)
        Me.TxtAmt.Name = "TxtAmt"
        Me.TxtAmt.Size = New System.Drawing.Size(53, 20)
        Me.TxtAmt.TabIndex = 8
        Me.TxtAmt.Text = "0"
        Me.TxtAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChkVAT
        '
        Me.ChkVAT.AutoSize = True
        Me.ChkVAT.Checked = True
        Me.ChkVAT.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkVAT.Location = New System.Drawing.Point(711, 75)
        Me.ChkVAT.Name = "ChkVAT"
        Me.ChkVAT.Size = New System.Drawing.Size(70, 17)
        Me.ChkVAT.TabIndex = 9
        Me.ChkVAT.Text = "VAT Incl."
        Me.ChkVAT.UseVisualStyleBackColor = True
        '
        'chkWzAIR
        '
        Me.chkWzAIR.AutoSize = True
        Me.chkWzAIR.Location = New System.Drawing.Point(711, 57)
        Me.chkWzAIR.Name = "chkWzAIR"
        Me.chkWzAIR.Size = New System.Drawing.Size(60, 17)
        Me.chkWzAIR.TabIndex = 12
        Me.chkWzAIR.Text = "WzAIR"
        Me.chkWzAIR.UseVisualStyleBackColor = True
        '
        'TxtValidFrom
        '
        Me.TxtValidFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtValidFrom.Location = New System.Drawing.Point(60, 35)
        Me.TxtValidFrom.Name = "TxtValidFrom"
        Me.TxtValidFrom.Size = New System.Drawing.Size(115, 20)
        Me.TxtValidFrom.TabIndex = 10
        '
        'TxtValidThru
        '
        Me.TxtValidThru.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.TxtValidThru.Location = New System.Drawing.Point(223, 34)
        Me.TxtValidThru.Name = "TxtValidThru"
        Me.TxtValidThru.Size = New System.Drawing.Size(70, 20)
        Me.TxtValidThru.TabIndex = 11
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(2, 39)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 13)
        Me.Label10.TabIndex = 4
        Me.Label10.Text = "Valid From"
        '
        'label15
        '
        Me.label15.AutoSize = True
        Me.label15.Location = New System.Drawing.Point(178, 39)
        Me.label15.Name = "label15"
        Me.label15.Size = New System.Drawing.Size(29, 13)
        Me.label15.TabIndex = 4
        Me.label15.Text = "Thru"
        '
        'CmbBase
        '
        Me.CmbBase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbBase.FormattingEnabled = True
        Me.CmbBase.Items.AddRange(New Object() {"Fare", "F+T"})
        Me.CmbBase.Location = New System.Drawing.Point(416, 34)
        Me.CmbBase.Name = "CmbBase"
        Me.CmbBase.Size = New System.Drawing.Size(51, 21)
        Me.CmbBase.TabIndex = 5
        Me.CmbBase.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(386, 39)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(31, 13)
        Me.Label12.TabIndex = 4
        Me.Label12.Text = "Base"
        Me.Label12.Visible = False
        '
        'CmbTKT
        '
        Me.CmbTKT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbTKT.FormattingEnabled = True
        Me.CmbTKT.Items.AddRange(New Object() {"A", "E", "I"})
        Me.CmbTKT.Location = New System.Drawing.Point(354, 12)
        Me.CmbTKT.Name = "CmbTKT"
        Me.CmbTKT.Size = New System.Drawing.Size(34, 21)
        Me.CmbTKT.TabIndex = 3
        '
        'CmbAL
        '
        Me.CmbAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Items.AddRange(New Object() {"LCC", "FSC"})
        Me.CmbAL.Location = New System.Drawing.Point(592, 33)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(53, 21)
        Me.CmbAL.TabIndex = 14
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(560, 38)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(20, 13)
        Me.Label11.TabIndex = 4
        Me.Label11.Text = "AL"
        '
        'CmbRtg
        '
        Me.CmbRtg.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbRtg.FormattingEnabled = True
        Me.CmbRtg.Items.AddRange(New Object() {"DOM", "REG", "REG2", "INTL", "INT2"})
        Me.CmbRtg.Location = New System.Drawing.Point(502, 12)
        Me.CmbRtg.Name = "CmbRtg"
        Me.CmbRtg.Size = New System.Drawing.Size(48, 21)
        Me.CmbRtg.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(470, 15)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "RTG"
        '
        'CmbCabin
        '
        Me.CmbCabin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCabin.FormattingEnabled = True
        Me.CmbCabin.Items.AddRange(New Object() {"*", "F", "C", "Y"})
        Me.CmbCabin.Location = New System.Drawing.Point(354, 35)
        Me.CmbCabin.Name = "CmbCabin"
        Me.CmbCabin.Size = New System.Drawing.Size(34, 21)
        Me.CmbCabin.TabIndex = 16
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(293, 39)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(34, 13)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Cabin"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(645, 36)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(41, 13)
        Me.Label13.TabIndex = 4
        Me.Label13.Text = "FltTime"
        '
        'TxtFltTime
        '
        Me.TxtFltTime.Location = New System.Drawing.Point(686, 33)
        Me.TxtFltTime.Name = "TxtFltTime"
        Me.TxtFltTime.Size = New System.Drawing.Size(37, 20)
        Me.TxtFltTime.TabIndex = 8
        Me.TxtFltTime.Text = "0"
        Me.TxtFltTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmbSegTKT
        '
        Me.cmbSegTKT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSegTKT.FormattingEnabled = True
        Me.cmbSegTKT.Items.AddRange(New Object() {"TKT", "SEG"})
        Me.cmbSegTKT.Location = New System.Drawing.Point(469, 34)
        Me.cmbSegTKT.Name = "cmbSegTKT"
        Me.cmbSegTKT.Size = New System.Drawing.Size(81, 21)
        Me.cmbSegTKT.TabIndex = 17
        '
        'TxtPersonal
        '
        Me.TxtPersonal.Location = New System.Drawing.Point(701, 12)
        Me.TxtPersonal.Name = "TxtPersonal"
        Me.TxtPersonal.Size = New System.Drawing.Size(76, 20)
        Me.TxtPersonal.TabIndex = 8
        Me.TxtPersonal.Text = "0"
        Me.TxtPersonal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(645, 15)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(48, 13)
        Me.Label14.TabIndex = 4
        Me.Label14.Text = "Personal"
        '
        'CmbOWRT
        '
        Me.CmbOWRT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbOWRT.FormattingEnabled = True
        Me.CmbOWRT.Items.AddRange(New Object() {"--", "OW", "RT"})
        Me.CmbOWRT.Location = New System.Drawing.Point(729, 33)
        Me.CmbOWRT.Name = "CmbOWRT"
        Me.CmbOWRT.Size = New System.Drawing.Size(48, 21)
        Me.CmbOWRT.TabIndex = 18
        '
        'CWT_ServiceFee
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.CmbOWRT)
        Me.Controls.Add(Me.cmbSegTKT)
        Me.Controls.Add(Me.CmbCabin)
        Me.Controls.Add(Me.CmbRtg)
        Me.Controls.Add(Me.CmbAL)
        Me.Controls.Add(Me.CmbTKT)
        Me.Controls.Add(Me.CmbBase)
        Me.Controls.Add(Me.TxtValidThru)
        Me.Controls.Add(Me.TxtValidFrom)
        Me.Controls.Add(Me.chkWzAIR)
        Me.Controls.Add(Me.ChkVAT)
        Me.Controls.Add(Me.LblCheckAll)
        Me.Controls.Add(Me.LstReg)
        Me.Controls.Add(Me.LblDelete)
        Me.Controls.Add(Me.GridSF)
        Me.Controls.Add(Me.LblUpdate)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.label15)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CmbTRX)
        Me.Controls.Add(Me.CmbCurr)
        Me.Controls.Add(Me.TxtFltTime)
        Me.Controls.Add(Me.TxtPersonal)
        Me.Controls.Add(Me.TxtAmt)
        Me.Controls.Add(Me.CmbService)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.CmbCust)
        Me.Controls.Add(Me.Label9)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CWT_ServiceFee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Corporate Travel Service :: RAS :. Service Fee"
        CType(Me.GridSF, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CmbCust As System.Windows.Forms.ComboBox
    Friend WithEvents CmbService As System.Windows.Forms.ComboBox
    Friend WithEvents CmbCurr As System.Windows.Forms.ComboBox
    Friend WithEvents CmbTRX As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GridSF As System.Windows.Forms.DataGridView
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LstReg As System.Windows.Forms.CheckedListBox
    Friend WithEvents LblCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtAmt As System.Windows.Forms.TextBox
    Friend WithEvents ChkVAT As System.Windows.Forms.CheckBox
    Friend WithEvents chkWzAIR As System.Windows.Forms.CheckBox
    Friend WithEvents TxtValidFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtValidThru As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents label15 As System.Windows.Forms.Label
    Friend WithEvents CmbBase As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents CmbTKT As System.Windows.Forms.ComboBox
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents CmbRtg As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CmbCabin As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtFltTime As System.Windows.Forms.TextBox
    Friend WithEvents cmbSegTKT As System.Windows.Forms.ComboBox
    Friend WithEvents TxtPersonal As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents CmbOWRT As System.Windows.Forms.ComboBox
End Class
