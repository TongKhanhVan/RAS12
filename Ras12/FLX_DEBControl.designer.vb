<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FLX_DEBControl
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
        Me.GridRQ = New System.Windows.Forms.DataGridView()
        Me.TxtRMK = New System.Windows.Forms.TextBox()
        Me.LblReject = New System.Windows.Forms.LinkLabel()
        Me.LblApprove = New System.Windows.Forms.LinkLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        CType(Me.GridRQ, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridRQ
        '
        Me.GridRQ.AllowUserToAddRows = False
        Me.GridRQ.AllowUserToDeleteRows = False
        Me.GridRQ.BackgroundColor = System.Drawing.Color.WhiteSmoke
        Me.GridRQ.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridRQ.Location = New System.Drawing.Point(0, 30)
        Me.GridRQ.Name = "GridRQ"
        Me.GridRQ.RowHeadersVisible = False
        Me.GridRQ.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridRQ.Size = New System.Drawing.Size(768, 433)
        Me.GridRQ.TabIndex = 0
        '
        'TxtRMK
        '
        Me.TxtRMK.Enabled = False
        Me.TxtRMK.Location = New System.Drawing.Point(42, 4)
        Me.TxtRMK.Name = "TxtRMK"
        Me.TxtRMK.Size = New System.Drawing.Size(294, 20)
        Me.TxtRMK.TabIndex = 1
        '
        'LblReject
        '
        Me.LblReject.AutoSize = True
        Me.LblReject.Location = New System.Drawing.Point(342, 7)
        Me.LblReject.Name = "LblReject"
        Me.LblReject.Size = New System.Drawing.Size(38, 13)
        Me.LblReject.TabIndex = 2
        Me.LblReject.TabStop = True
        Me.LblReject.Text = "Reject"
        Me.LblReject.Visible = False
        '
        'LblApprove
        '
        Me.LblApprove.AutoSize = True
        Me.LblApprove.Location = New System.Drawing.Point(386, 7)
        Me.LblApprove.Name = "LblApprove"
        Me.LblApprove.Size = New System.Drawing.Size(47, 13)
        Me.LblApprove.TabIndex = 3
        Me.LblApprove.TabStop = True
        Me.LblApprove.Text = "Approve"
        Me.LblApprove.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Note"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(3, 4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(779, 492)
        Me.TabControl1.TabIndex = 6
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GridRQ)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.TxtRMK)
        Me.TabPage1.Controls.Add(Me.LblReject)
        Me.TabPage1.Controls.Add(Me.LblApprove)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(771, 466)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "DEB Approval"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'FLX_DEBControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FLX_DEBControl"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: Flex Tour 10 :. DEB Approval"
        CType(Me.GridRQ, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GridRQ As System.Windows.Forms.DataGridView
    Friend WithEvents TxtRMK As System.Windows.Forms.TextBox
    Friend WithEvents LblReject As System.Windows.Forms.LinkLabel
    Friend WithEvents LblApprove As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
End Class
