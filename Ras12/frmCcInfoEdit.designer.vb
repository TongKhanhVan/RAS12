<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCcInfoEdit
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtCardHolder = New System.Windows.Forms.TextBox()
        Me.txtCardNbr = New System.Windows.Forms.TextBox()
        Me.txtExpDate = New System.Windows.Forms.TextBox()
        Me.txtRemark = New System.Windows.Forms.TextBox()
        Me.cboCardType = New System.Windows.Forms.ComboBox()
        Me.lbkSave = New System.Windows.Forms.LinkLabel()
        Me.cboBiz = New System.Windows.Forms.ComboBox()
        Me.txtCustShortName = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "CustShortName"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(0, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "CardHolder"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(0, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "CardType/Nbr"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(0, 89)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(91, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "ExpDate(MM/YY)"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(0, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(67, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Biz/Per/Unk"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(0, 136)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(44, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Remark"
        '
        'txtCardHolder
        '
        Me.txtCardHolder.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCardHolder.Location = New System.Drawing.Point(107, 33)
        Me.txtCardHolder.Name = "txtCardHolder"
        Me.txtCardHolder.Size = New System.Drawing.Size(284, 20)
        Me.txtCardHolder.TabIndex = 1
        '
        'txtCardNbr
        '
        Me.txtCardNbr.Location = New System.Drawing.Point(160, 59)
        Me.txtCardNbr.Name = "txtCardNbr"
        Me.txtCardNbr.Size = New System.Drawing.Size(231, 20)
        Me.txtCardNbr.TabIndex = 3
        '
        'txtExpDate
        '
        Me.txtExpDate.Location = New System.Drawing.Point(107, 82)
        Me.txtExpDate.Name = "txtExpDate"
        Me.txtExpDate.Size = New System.Drawing.Size(54, 20)
        Me.txtExpDate.TabIndex = 4
        '
        'txtRemark
        '
        Me.txtRemark.Location = New System.Drawing.Point(107, 135)
        Me.txtRemark.Name = "txtRemark"
        Me.txtRemark.Size = New System.Drawing.Size(284, 20)
        Me.txtRemark.TabIndex = 6
        '
        'cboCardType
        '
        Me.cboCardType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCardType.FormattingEnabled = True
        Me.cboCardType.Items.AddRange(New Object() {"VI", "CA", "AX", "DC"})
        Me.cboCardType.Location = New System.Drawing.Point(107, 58)
        Me.cboCardType.Name = "cboCardType"
        Me.cboCardType.Size = New System.Drawing.Size(54, 21)
        Me.cboCardType.TabIndex = 2
        '
        'lbkSave
        '
        Me.lbkSave.AutoSize = True
        Me.lbkSave.Location = New System.Drawing.Point(359, 171)
        Me.lbkSave.Name = "lbkSave"
        Me.lbkSave.Size = New System.Drawing.Size(32, 13)
        Me.lbkSave.TabIndex = 7
        Me.lbkSave.TabStop = True
        Me.lbkSave.Text = "Save"
        '
        'cboBiz
        '
        Me.cboBiz.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBiz.FormattingEnabled = True
        Me.cboBiz.Items.AddRange(New Object() {"UNK", "BIZ", "PER"})
        Me.cboBiz.Location = New System.Drawing.Point(107, 108)
        Me.cboBiz.Name = "cboBiz"
        Me.cboBiz.Size = New System.Drawing.Size(54, 21)
        Me.cboBiz.TabIndex = 8
        '
        'txtCustShortName
        '
        Me.txtCustShortName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustShortName.Location = New System.Drawing.Point(107, 6)
        Me.txtCustShortName.Name = "txtCustShortName"
        Me.txtCustShortName.ReadOnly = True
        Me.txtCustShortName.Size = New System.Drawing.Size(284, 20)
        Me.txtCustShortName.TabIndex = 9
        '
        'frmCcInfoEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(392, 273)
        Me.Controls.Add(Me.txtCustShortName)
        Me.Controls.Add(Me.cboBiz)
        Me.Controls.Add(Me.lbkSave)
        Me.Controls.Add(Me.cboCardType)
        Me.Controls.Add(Me.txtRemark)
        Me.Controls.Add(Me.txtExpDate)
        Me.Controls.Add(Me.txtCardNbr)
        Me.Controls.Add(Me.txtCardHolder)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmCcInfoEdit"
        Me.Text = "CcInfoEdit"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtCardHolder As System.Windows.Forms.TextBox
    Friend WithEvents txtCardNbr As System.Windows.Forms.TextBox
    Friend WithEvents txtExpDate As System.Windows.Forms.TextBox
    Friend WithEvents txtRemark As System.Windows.Forms.TextBox
    Friend WithEvents cboCardType As System.Windows.Forms.ComboBox
    Friend WithEvents lbkSave As System.Windows.Forms.LinkLabel
    Friend WithEvents cboBiz As System.Windows.Forms.ComboBox
    Friend WithEvents txtCustShortName As System.Windows.Forms.TextBox
End Class
