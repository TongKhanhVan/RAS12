<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LockTheUnLocked
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
        Me.GridPendingTRX = New System.Windows.Forms.DataGridView()
        Me.LblLock = New System.Windows.Forms.LinkLabel()
        Me.TxtTRXNo = New System.Windows.Forms.TextBox()
        Me.LblUnLockThis = New System.Windows.Forms.LinkLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabDaily = New System.Windows.Forms.TabPage()
        Me.TabAL = New System.Windows.Forms.TabPage()
        Me.LblUnlockALRPT = New System.Windows.Forms.LinkLabel()
        Me.GridALRPT = New System.Windows.Forms.DataGridView()
        CType(Me.GridPendingTRX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabDaily.SuspendLayout()
        Me.TabAL.SuspendLayout()
        CType(Me.GridALRPT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridPendingTRX
        '
        Me.GridPendingTRX.AllowUserToAddRows = False
        Me.GridPendingTRX.AllowUserToDeleteRows = False
        Me.GridPendingTRX.BackgroundColor = System.Drawing.Color.Honeydew
        Me.GridPendingTRX.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPendingTRX.Location = New System.Drawing.Point(1, 31)
        Me.GridPendingTRX.Name = "GridPendingTRX"
        Me.GridPendingTRX.RowHeadersVisible = False
        Me.GridPendingTRX.Size = New System.Drawing.Size(467, 243)
        Me.GridPendingTRX.TabIndex = 0
        '
        'LblLock
        '
        Me.LblLock.AutoSize = True
        Me.LblLock.Location = New System.Drawing.Point(404, 277)
        Me.LblLock.Name = "LblLock"
        Me.LblLock.Size = New System.Drawing.Size(61, 13)
        Me.LblLock.TabIndex = 1
        Me.LblLock.TabStop = True
        Me.LblLock.Text = "Lock Again"
        '
        'TxtTRXNo
        '
        Me.TxtTRXNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTRXNo.Location = New System.Drawing.Point(55, 6)
        Me.TxtTRXNo.Name = "TxtTRXNo"
        Me.TxtTRXNo.Size = New System.Drawing.Size(121, 20)
        Me.TxtTRXNo.TabIndex = 2
        '
        'LblUnLockThis
        '
        Me.LblUnLockThis.AutoSize = True
        Me.LblUnLockThis.Enabled = False
        Me.LblUnLockThis.Location = New System.Drawing.Point(187, 9)
        Me.LblUnLockThis.Name = "LblUnLockThis"
        Me.LblUnLockThis.Size = New System.Drawing.Size(87, 13)
        Me.LblUnLockThis.TabIndex = 3
        Me.LblUnLockThis.TabStop = True
        Me.LblUnLockThis.Text = "UnLockThisTRX"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(0, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "TRX No."
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabDaily)
        Me.TabControl1.Controls.Add(Me.TabAL)
        Me.TabControl1.Location = New System.Drawing.Point(3, 4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(479, 322)
        Me.TabControl1.TabIndex = 5
        '
        'TabDaily
        '
        Me.TabDaily.Controls.Add(Me.TxtTRXNo)
        Me.TabDaily.Controls.Add(Me.LblLock)
        Me.TabDaily.Controls.Add(Me.Label1)
        Me.TabDaily.Controls.Add(Me.GridPendingTRX)
        Me.TabDaily.Controls.Add(Me.LblUnLockThis)
        Me.TabDaily.Location = New System.Drawing.Point(4, 22)
        Me.TabDaily.Name = "TabDaily"
        Me.TabDaily.Padding = New System.Windows.Forms.Padding(3)
        Me.TabDaily.Size = New System.Drawing.Size(471, 296)
        Me.TabDaily.TabIndex = 0
        Me.TabDaily.Text = "DailyRPT"
        Me.TabDaily.UseVisualStyleBackColor = True
        '
        'TabAL
        '
        Me.TabAL.Controls.Add(Me.LblUnlockALRPT)
        Me.TabAL.Controls.Add(Me.GridALRPT)
        Me.TabAL.Location = New System.Drawing.Point(4, 22)
        Me.TabAL.Name = "TabAL"
        Me.TabAL.Padding = New System.Windows.Forms.Padding(3)
        Me.TabAL.Size = New System.Drawing.Size(471, 296)
        Me.TabAL.TabIndex = 1
        Me.TabAL.Text = "AL RPT"
        Me.TabAL.UseVisualStyleBackColor = True
        '
        'LblUnlockALRPT
        '
        Me.LblUnlockALRPT.AutoSize = True
        Me.LblUnlockALRPT.Location = New System.Drawing.Point(423, 278)
        Me.LblUnlockALRPT.Name = "LblUnlockALRPT"
        Me.LblUnlockALRPT.Size = New System.Drawing.Size(45, 13)
        Me.LblUnlockALRPT.TabIndex = 1
        Me.LblUnlockALRPT.TabStop = True
        Me.LblUnlockALRPT.Text = "UnLock"
        Me.LblUnlockALRPT.Visible = False
        '
        'GridALRPT
        '
        Me.GridALRPT.AllowUserToAddRows = False
        Me.GridALRPT.AllowUserToDeleteRows = False
        Me.GridALRPT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridALRPT.Location = New System.Drawing.Point(5, 3)
        Me.GridALRPT.Name = "GridALRPT"
        Me.GridALRPT.RowHeadersVisible = False
        Me.GridALRPT.Size = New System.Drawing.Size(463, 266)
        Me.GridALRPT.TabIndex = 0
        '
        'LockTheUnLocked
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(482, 326)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "LockTheUnLocked"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TranViet Airlines :: RAS :. Lock The UnLocked TRX"
        CType(Me.GridPendingTRX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabDaily.ResumeLayout(False)
        Me.TabDaily.PerformLayout()
        Me.TabAL.ResumeLayout(False)
        Me.TabAL.PerformLayout()
        CType(Me.GridALRPT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GridPendingTRX As System.Windows.Forms.DataGridView
    Friend WithEvents LblLock As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtTRXNo As System.Windows.Forms.TextBox
    Friend WithEvents LblUnLockThis As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabDaily As System.Windows.Forms.TabPage
    Friend WithEvents TabAL As System.Windows.Forms.TabPage
    Friend WithEvents LblUnlockALRPT As System.Windows.Forms.LinkLabel
    Friend WithEvents GridALRPT As System.Windows.Forms.DataGridView
End Class
