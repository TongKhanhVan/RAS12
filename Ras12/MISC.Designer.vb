<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMISC
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
        Me.PnlCC_Customer = New System.Windows.Forms.Panel()
        Me.LblChangeMinBLC = New System.Windows.Forms.LinkLabel()
        Me.txtMinBLC = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ChkAL = New System.Windows.Forms.CheckedListBox()
        Me.LblDelete = New System.Windows.Forms.LinkLabel()
        Me.CmdCC_Cust_Save = New System.Windows.Forms.Button()
        Me.txtPPCoef = New System.Windows.Forms.TextBox()
        Me.txtCRCoef = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCC_Cust_CustName = New System.Windows.Forms.TextBox()
        Me.CmbCC_Cust_Customer = New System.Windows.Forms.ComboBox()
        Me.GridCC_Cust = New System.Windows.Forms.DataGridView()
        Me.ChkAll = New System.Windows.Forms.CheckBox()
        Me.PnlCRExt = New System.Windows.Forms.Panel()
        Me.CmdCRReject = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.CmdCRApprove = New System.Windows.Forms.Button()
        Me.GridQ4Approval = New System.Windows.Forms.DataGridView()
        Me.txtCRExtNote = New System.Windows.Forms.TextBox()
        Me.CmbVAT = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PnlCC_Customer.SuspendLayout()
        CType(Me.GridCC_Cust, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PnlCRExt.SuspendLayout()
        CType(Me.GridQ4Approval, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PnlCC_Customer
        '
        Me.PnlCC_Customer.Controls.Add(Me.Label2)
        Me.PnlCC_Customer.Controls.Add(Me.CmbVAT)
        Me.PnlCC_Customer.Controls.Add(Me.LblChangeMinBLC)
        Me.PnlCC_Customer.Controls.Add(Me.txtMinBLC)
        Me.PnlCC_Customer.Controls.Add(Me.Label1)
        Me.PnlCC_Customer.Controls.Add(Me.ChkAL)
        Me.PnlCC_Customer.Controls.Add(Me.LblDelete)
        Me.PnlCC_Customer.Controls.Add(Me.CmdCC_Cust_Save)
        Me.PnlCC_Customer.Controls.Add(Me.txtPPCoef)
        Me.PnlCC_Customer.Controls.Add(Me.txtCRCoef)
        Me.PnlCC_Customer.Controls.Add(Me.Label4)
        Me.PnlCC_Customer.Controls.Add(Me.Label6)
        Me.PnlCC_Customer.Controls.Add(Me.Label5)
        Me.PnlCC_Customer.Controls.Add(Me.Label8)
        Me.PnlCC_Customer.Controls.Add(Me.Label3)
        Me.PnlCC_Customer.Controls.Add(Me.txtCC_Cust_CustName)
        Me.PnlCC_Customer.Controls.Add(Me.CmbCC_Cust_Customer)
        Me.PnlCC_Customer.Controls.Add(Me.GridCC_Cust)
        Me.PnlCC_Customer.Controls.Add(Me.ChkAll)
        Me.PnlCC_Customer.Location = New System.Drawing.Point(2, 2)
        Me.PnlCC_Customer.Name = "PnlCC_Customer"
        Me.PnlCC_Customer.Size = New System.Drawing.Size(803, 143)
        Me.PnlCC_Customer.TabIndex = 1
        Me.PnlCC_Customer.Visible = False
        '
        'LblChangeMinBLC
        '
        Me.LblChangeMinBLC.AutoSize = True
        Me.LblChangeMinBLC.Location = New System.Drawing.Point(713, 62)
        Me.LblChangeMinBLC.Name = "LblChangeMinBLC"
        Me.LblChangeMinBLC.Size = New System.Drawing.Size(87, 13)
        Me.LblChangeMinBLC.TabIndex = 12
        Me.LblChangeMinBLC.TabStop = True
        Me.LblChangeMinBLC.Text = "Change Min BLC"
        '
        'txtMinBLC
        '
        Me.txtMinBLC.Location = New System.Drawing.Point(684, 23)
        Me.txtMinBLC.Name = "txtMinBLC"
        Me.txtMinBLC.Size = New System.Drawing.Size(116, 20)
        Me.txtMinBLC.TabIndex = 11
        Me.txtMinBLC.Text = "0"
        Me.txtMinBLC.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(681, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "MinBLC"
        '
        'ChkAL
        '
        Me.ChkAL.CheckOnClick = True
        Me.ChkAL.ColumnWidth = 40
        Me.ChkAL.FormattingEnabled = True
        Me.ChkAL.Location = New System.Drawing.Point(326, 29)
        Me.ChkAL.MultiColumn = True
        Me.ChkAL.Name = "ChkAL"
        Me.ChkAL.Size = New System.Drawing.Size(239, 49)
        Me.ChkAL.TabIndex = 8
        '
        'LblDelete
        '
        Me.LblDelete.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LblDelete.AutoSize = True
        Me.LblDelete.Enabled = False
        Me.LblDelete.Location = New System.Drawing.Point(572, 8)
        Me.LblDelete.Name = "LblDelete"
        Me.LblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LblDelete.TabIndex = 7
        Me.LblDelete.TabStop = True
        Me.LblDelete.Text = "Delete"
        '
        'CmdCC_Cust_Save
        '
        Me.CmdCC_Cust_Save.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdCC_Cust_Save.Enabled = False
        Me.CmdCC_Cust_Save.Location = New System.Drawing.Point(571, 58)
        Me.CmdCC_Cust_Save.Name = "CmdCC_Cust_Save"
        Me.CmdCC_Cust_Save.Size = New System.Drawing.Size(53, 20)
        Me.CmdCC_Cust_Save.TabIndex = 5
        Me.CmdCC_Cust_Save.Text = "Save"
        Me.CmdCC_Cust_Save.UseVisualStyleBackColor = True
        '
        'txtPPCoef
        '
        Me.txtPPCoef.Location = New System.Drawing.Point(196, 27)
        Me.txtPPCoef.Name = "txtPPCoef"
        Me.txtPPCoef.Size = New System.Drawing.Size(33, 20)
        Me.txtPPCoef.TabIndex = 2
        Me.txtPPCoef.Text = "1.0"
        Me.txtPPCoef.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCRCoef
        '
        Me.txtCRCoef.Location = New System.Drawing.Point(117, 27)
        Me.txtCRCoef.Name = "txtCRCoef"
        Me.txtCRCoef.Size = New System.Drawing.Size(34, 20)
        Me.txtCRCoef.TabIndex = 1
        Me.txtCRCoef.Text = "0.0"
        Me.txtCRCoef.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(88, 30)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(31, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "PSP"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(167, 30)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "PPD"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 30)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 13)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Safe Coefficient: "
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(235, 30)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 13)
        Me.Label8.TabIndex = 3
        Me.Label8.Text = "Airlines"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(107, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Customer ShortName"
        '
        'txtCC_Cust_CustName
        '
        Me.txtCC_Cust_CustName.Location = New System.Drawing.Point(234, 5)
        Me.txtCC_Cust_CustName.Name = "txtCC_Cust_CustName"
        Me.txtCC_Cust_CustName.ReadOnly = True
        Me.txtCC_Cust_CustName.Size = New System.Drawing.Size(332, 20)
        Me.txtCC_Cust_CustName.TabIndex = 2
        '
        'CmbCC_Cust_Customer
        '
        Me.CmbCC_Cust_Customer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCC_Cust_Customer.FormattingEnabled = True
        Me.CmbCC_Cust_Customer.Location = New System.Drawing.Point(117, 5)
        Me.CmbCC_Cust_Customer.Name = "CmbCC_Cust_Customer"
        Me.CmbCC_Cust_Customer.Size = New System.Drawing.Size(112, 21)
        Me.CmbCC_Cust_Customer.TabIndex = 1
        '
        'GridCC_Cust
        '
        Me.GridCC_Cust.AllowUserToAddRows = False
        Me.GridCC_Cust.AllowUserToDeleteRows = False
        Me.GridCC_Cust.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridCC_Cust.BackgroundColor = System.Drawing.Color.Snow
        Me.GridCC_Cust.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridCC_Cust.Location = New System.Drawing.Point(1, 84)
        Me.GridCC_Cust.Name = "GridCC_Cust"
        Me.GridCC_Cust.ReadOnly = True
        Me.GridCC_Cust.RowHeadersVisible = False
        Me.GridCC_Cust.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridCC_Cust.Size = New System.Drawing.Size(801, 56)
        Me.GridCC_Cust.TabIndex = 0
        '
        'ChkAll
        '
        Me.ChkAll.AutoSize = True
        Me.ChkAll.Location = New System.Drawing.Point(281, 31)
        Me.ChkAll.Name = "ChkAll"
        Me.ChkAll.Size = New System.Drawing.Size(45, 17)
        Me.ChkAll.TabIndex = 9
        Me.ChkAll.Text = "ALL"
        Me.ChkAll.UseVisualStyleBackColor = True
        '
        'PnlCRExt
        '
        Me.PnlCRExt.Controls.Add(Me.CmdCRReject)
        Me.PnlCRExt.Controls.Add(Me.Label7)
        Me.PnlCRExt.Controls.Add(Me.CmdCRApprove)
        Me.PnlCRExt.Controls.Add(Me.GridQ4Approval)
        Me.PnlCRExt.Controls.Add(Me.txtCRExtNote)
        Me.PnlCRExt.Location = New System.Drawing.Point(2, 151)
        Me.PnlCRExt.Name = "PnlCRExt"
        Me.PnlCRExt.Size = New System.Drawing.Size(803, 61)
        Me.PnlCRExt.TabIndex = 2
        Me.PnlCRExt.Visible = False
        '
        'CmdCRReject
        '
        Me.CmdCRReject.Enabled = False
        Me.CmdCRReject.Location = New System.Drawing.Point(636, 4)
        Me.CmdCRReject.Name = "CmdCRReject"
        Me.CmdCRReject.Size = New System.Drawing.Size(84, 26)
        Me.CmdCRReject.TabIndex = 4
        Me.CmdCRReject.Text = "Reject"
        Me.CmdCRReject.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(118, 10)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(30, 13)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Note"
        '
        'CmdCRApprove
        '
        Me.CmdCRApprove.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdCRApprove.Enabled = False
        Me.CmdCRApprove.Location = New System.Drawing.Point(726, 3)
        Me.CmdCRApprove.Name = "CmdCRApprove"
        Me.CmdCRApprove.Size = New System.Drawing.Size(74, 26)
        Me.CmdCRApprove.TabIndex = 1
        Me.CmdCRApprove.Text = "Approve"
        Me.CmdCRApprove.UseVisualStyleBackColor = True
        '
        'GridQ4Approval
        '
        Me.GridQ4Approval.AllowUserToAddRows = False
        Me.GridQ4Approval.AllowUserToDeleteRows = False
        Me.GridQ4Approval.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridQ4Approval.BackgroundColor = System.Drawing.Color.Ivory
        Me.GridQ4Approval.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridQ4Approval.Location = New System.Drawing.Point(2, 33)
        Me.GridQ4Approval.Name = "GridQ4Approval"
        Me.GridQ4Approval.RowHeadersVisible = False
        Me.GridQ4Approval.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridQ4Approval.Size = New System.Drawing.Size(800, 25)
        Me.GridQ4Approval.TabIndex = 0
        '
        'txtCRExtNote
        '
        Me.txtCRExtNote.Location = New System.Drawing.Point(154, 7)
        Me.txtCRExtNote.MaxLength = 200
        Me.txtCRExtNote.Name = "txtCRExtNote"
        Me.txtCRExtNote.Size = New System.Drawing.Size(476, 20)
        Me.txtCRExtNote.TabIndex = 2
        '
        'CmbVAT
        '
        Me.CmbVAT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbVAT.FormattingEnabled = True
        Me.CmbVAT.Items.AddRange(New Object() {"TVTR", "TVPM"})
        Me.CmbVAT.Location = New System.Drawing.Point(117, 48)
        Me.CmbVAT.Name = "CmbVAT"
        Me.CmbVAT.Size = New System.Drawing.Size(112, 21)
        Me.CmbVAT.TabIndex = 13
        Me.CmbVAT.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "VAT Invoice Of"
        Me.Label2.Visible = False
        '
        'frmMISC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(807, 516)
        Me.Controls.Add(Me.PnlCRExt)
        Me.Controls.Add(Me.PnlCC_Customer)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMISC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. MISC"
        Me.PnlCC_Customer.ResumeLayout(False)
        Me.PnlCC_Customer.PerformLayout()
        CType(Me.GridCC_Cust, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PnlCRExt.ResumeLayout(False)
        Me.PnlCRExt.PerformLayout()
        CType(Me.GridQ4Approval, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PnlCC_Customer As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtCC_Cust_CustName As System.Windows.Forms.TextBox
    Friend WithEvents CmbCC_Cust_Customer As System.Windows.Forms.ComboBox
    Friend WithEvents GridCC_Cust As System.Windows.Forms.DataGridView
    Friend WithEvents txtPPCoef As System.Windows.Forms.TextBox
    Friend WithEvents txtCRCoef As System.Windows.Forms.TextBox
    Friend WithEvents CmdCC_Cust_Save As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents PnlCRExt As System.Windows.Forms.Panel
    Friend WithEvents GridQ4Approval As System.Windows.Forms.DataGridView
    Friend WithEvents CmdCRApprove As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtCRExtNote As System.Windows.Forms.TextBox
    Friend WithEvents CmdCRReject As System.Windows.Forms.Button
    Friend WithEvents LblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkAL As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ChkAll As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LblChangeMinBLC As System.Windows.Forms.LinkLabel
    Friend WithEvents txtMinBLC As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CmbVAT As System.Windows.Forms.ComboBox
End Class
