<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectCcNbr
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
        Me.chkHideCcNbr = New System.Windows.Forms.CheckBox()
        Me.chkShowAll = New System.Windows.Forms.CheckBox()
        Me.lbkSelect = New System.Windows.Forms.LinkLabel()
        CType(Me.dgCcList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgCcList
        '
        Me.dgCcList.AllowUserToAddRows = False
        Me.dgCcList.AllowUserToDeleteRows = False
        Me.dgCcList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgCcList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgCcList.Location = New System.Drawing.Point(2, 12)
        Me.dgCcList.Name = "dgCcList"
        Me.dgCcList.ReadOnly = True
        Me.dgCcList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgCcList.Size = New System.Drawing.Size(1012, 524)
        Me.dgCcList.TabIndex = 58
        '
        'chkHideCcNbr
        '
        Me.chkHideCcNbr.AutoSize = True
        Me.chkHideCcNbr.Checked = True
        Me.chkHideCcNbr.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHideCcNbr.Location = New System.Drawing.Point(733, 544)
        Me.chkHideCcNbr.Name = "chkHideCcNbr"
        Me.chkHideCcNbr.Size = New System.Drawing.Size(78, 17)
        Me.chkHideCcNbr.TabIndex = 64
        Me.chkHideCcNbr.Text = "HideCcNbr"
        Me.chkHideCcNbr.UseVisualStyleBackColor = True
        '
        'chkShowAll
        '
        Me.chkShowAll.AutoSize = True
        Me.chkShowAll.Location = New System.Drawing.Point(817, 542)
        Me.chkShowAll.Name = "chkShowAll"
        Me.chkShowAll.Size = New System.Drawing.Size(64, 17)
        Me.chkShowAll.TabIndex = 65
        Me.chkShowAll.Text = "ShowAll"
        Me.chkShowAll.UseVisualStyleBackColor = True
        '
        'lbkSelect
        '
        Me.lbkSelect.AutoSize = True
        Me.lbkSelect.Location = New System.Drawing.Point(12, 548)
        Me.lbkSelect.Name = "lbkSelect"
        Me.lbkSelect.Size = New System.Drawing.Size(37, 13)
        Me.lbkSelect.TabIndex = 66
        Me.lbkSelect.TabStop = True
        Me.lbkSelect.Text = "Select"
        '
        'frmSelectCcNbr
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 573)
        Me.Controls.Add(Me.lbkSelect)
        Me.Controls.Add(Me.chkShowAll)
        Me.Controls.Add(Me.chkHideCcNbr)
        Me.Controls.Add(Me.dgCcList)
        Me.Name = "frmSelectCcNbr"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SelectCcNbr"
        CType(Me.dgCcList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgCcList As System.Windows.Forms.DataGridView
    Friend WithEvents chkHideCcNbr As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowAll As System.Windows.Forms.CheckBox
    Friend WithEvents lbkSelect As System.Windows.Forms.LinkLabel
End Class
