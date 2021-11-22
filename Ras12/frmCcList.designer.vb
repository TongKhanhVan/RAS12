<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCcList
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
        Me.dgCcList = New System.Windows.Forms.DataGridView()
        Me.cboCustShortName = New System.Windows.Forms.ComboBox()
        Me.lbkNew = New System.Windows.Forms.LinkLabel()
        Me.lbkExpire = New System.Windows.Forms.LinkLabel()
        Me.cboStatus = New System.Windows.Forms.ComboBox()
        Me.chkHideCcNbr = New System.Windows.Forms.CheckBox()
        CType(Me.dgCcList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgCcList
        '
        Me.dgCcList.AllowUserToAddRows = False
        Me.dgCcList.AllowUserToDeleteRows = False
        Me.dgCcList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgCcList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCcList.Location = New System.Drawing.Point(1, 42)
        Me.dgCcList.Name = "dgCcList"
        Me.dgCcList.ReadOnly = True
        Me.dgCcList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgCcList.Size = New System.Drawing.Size(1012, 499)
        Me.dgCcList.TabIndex = 57
        '
        'cboCustShortName
        '
        Me.cboCustShortName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustShortName.FormattingEnabled = True
        Me.cboCustShortName.Location = New System.Drawing.Point(1, 12)
        Me.cboCustShortName.Name = "cboCustShortName"
        Me.cboCustShortName.Size = New System.Drawing.Size(145, 21)
        Me.cboCustShortName.TabIndex = 58
        '
        'lbkNew
        '
        Me.lbkNew.AutoSize = True
        Me.lbkNew.Location = New System.Drawing.Point(-2, 551)
        Me.lbkNew.Name = "lbkNew"
        Me.lbkNew.Size = New System.Drawing.Size(29, 13)
        Me.lbkNew.TabIndex = 59
        Me.lbkNew.TabStop = True
        Me.lbkNew.Text = "New"
        '
        'lbkExpire
        '
        Me.lbkExpire.AutoSize = True
        Me.lbkExpire.Location = New System.Drawing.Point(33, 551)
        Me.lbkExpire.Name = "lbkExpire"
        Me.lbkExpire.Size = New System.Drawing.Size(36, 13)
        Me.lbkExpire.TabIndex = 61
        Me.lbkExpire.TabStop = True
        Me.lbkExpire.Text = "Expire"
        '
        'cboStatus
        '
        Me.cboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatus.FormattingEnabled = True
        Me.cboStatus.Items.AddRange(New Object() {"OK", "EX"})
        Me.cboStatus.Location = New System.Drawing.Point(169, 12)
        Me.cboStatus.Name = "cboStatus"
        Me.cboStatus.Size = New System.Drawing.Size(88, 21)
        Me.cboStatus.TabIndex = 62
        '
        'chkHideCcNbr
        '
        Me.chkHideCcNbr.AutoSize = True
        Me.chkHideCcNbr.Checked = True
        Me.chkHideCcNbr.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHideCcNbr.Location = New System.Drawing.Point(275, 12)
        Me.chkHideCcNbr.Name = "chkHideCcNbr"
        Me.chkHideCcNbr.Size = New System.Drawing.Size(78, 17)
        Me.chkHideCcNbr.TabIndex = 63
        Me.chkHideCcNbr.Text = "HideCcNbr"
        Me.chkHideCcNbr.UseVisualStyleBackColor = True
        '
        'frmCcList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 573)
        Me.Controls.Add(Me.chkHideCcNbr)
        Me.Controls.Add(Me.cboStatus)
        Me.Controls.Add(Me.lbkExpire)
        Me.Controls.Add(Me.lbkNew)
        Me.Controls.Add(Me.cboCustShortName)
        Me.Controls.Add(Me.dgCcList)
        Me.Name = "frmCcList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CcList"
        CType(Me.dgCcList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgCcList As System.Windows.Forms.DataGridView
    Friend WithEvents cboCustShortName As System.Windows.Forms.ComboBox
    Friend WithEvents lbkNew As System.Windows.Forms.LinkLabel
    Friend WithEvents lbkExpire As System.Windows.Forms.LinkLabel
    Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
    Friend WithEvents chkHideCcNbr As System.Windows.Forms.CheckBox
End Class
