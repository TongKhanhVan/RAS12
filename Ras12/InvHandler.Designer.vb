<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InvHandler
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.LckLblDelete = New System.Windows.Forms.LinkLabel
        Me.LblRePrint = New System.Windows.Forms.LinkLabel
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.GridInvHandlerInv = New System.Windows.Forms.DataGridView
        Me.LblPreviewInv = New System.Windows.Forms.LinkLabel
        Me.LblSearchInv = New System.Windows.Forms.LinkLabel
        Me.TxtInvHandlerInv2Search = New System.Windows.Forms.TextBox
        Me.CmbAL = New System.Windows.Forms.ComboBox
        CType(Me.GridInvHandlerInv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LckLblDelete
        '
        Me.LckLblDelete.AutoSize = True
        Me.LckLblDelete.Location = New System.Drawing.Point(740, 358)
        Me.LckLblDelete.Name = "LckLblDelete"
        Me.LckLblDelete.Size = New System.Drawing.Size(38, 13)
        Me.LckLblDelete.TabIndex = 6
        Me.LckLblDelete.TabStop = True
        Me.LckLblDelete.Text = "Delete"
        Me.LckLblDelete.Visible = False
        '
        'LblRePrint
        '
        Me.LblRePrint.AutoSize = True
        Me.LblRePrint.Location = New System.Drawing.Point(331, 9)
        Me.LblRePrint.Name = "LblRePrint"
        Me.LblRePrint.Size = New System.Drawing.Size(42, 13)
        Me.LblRePrint.TabIndex = 6
        Me.LblRePrint.TabStop = True
        Me.LblRePrint.Text = "RePrint"
        Me.LblRePrint.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(89, 9)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 13)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "InvNo."
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(35, 13)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "Airline"
        '
        'GridInvHandlerInv
        '
        Me.GridInvHandlerInv.AllowUserToAddRows = False
        Me.GridInvHandlerInv.AllowUserToDeleteRows = False
        Me.GridInvHandlerInv.BackgroundColor = System.Drawing.Color.Beige
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridInvHandlerInv.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.GridInvHandlerInv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridInvHandlerInv.Location = New System.Drawing.Point(1, 32)
        Me.GridInvHandlerInv.Name = "GridInvHandlerInv"
        Me.GridInvHandlerInv.ReadOnly = True
        Me.GridInvHandlerInv.RowHeadersVisible = False
        Me.GridInvHandlerInv.Size = New System.Drawing.Size(777, 323)
        Me.GridInvHandlerInv.TabIndex = 3
        '
        'LblPreviewInv
        '
        Me.LblPreviewInv.AutoSize = True
        Me.LblPreviewInv.Location = New System.Drawing.Point(280, 9)
        Me.LblPreviewInv.Name = "LblPreviewInv"
        Me.LblPreviewInv.Size = New System.Drawing.Size(45, 13)
        Me.LblPreviewInv.TabIndex = 2
        Me.LblPreviewInv.TabStop = True
        Me.LblPreviewInv.Text = "Preview"
        Me.LblPreviewInv.Visible = False
        '
        'LblSearchInv
        '
        Me.LblSearchInv.AutoSize = True
        Me.LblSearchInv.Location = New System.Drawing.Point(231, 9)
        Me.LblSearchInv.Name = "LblSearchInv"
        Me.LblSearchInv.Size = New System.Drawing.Size(41, 13)
        Me.LblSearchInv.TabIndex = 2
        Me.LblSearchInv.TabStop = True
        Me.LblSearchInv.Text = "Search"
        '
        'TxtInvHandlerInv2Search
        '
        Me.TxtInvHandlerInv2Search.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtInvHandlerInv2Search.Location = New System.Drawing.Point(134, 6)
        Me.TxtInvHandlerInv2Search.Name = "TxtInvHandlerInv2Search"
        Me.TxtInvHandlerInv2Search.Size = New System.Drawing.Size(91, 20)
        Me.TxtInvHandlerInv2Search.TabIndex = 1
        '
        'CmbAL
        '
        Me.CmbAL.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbAL.FormattingEnabled = True
        Me.CmbAL.Location = New System.Drawing.Point(41, 5)
        Me.CmbAL.Name = "CmbAL"
        Me.CmbAL.Size = New System.Drawing.Size(43, 21)
        Me.CmbAL.TabIndex = 0
        '
        'InvHandler
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(779, 380)
        Me.Controls.Add(Me.LckLblDelete)
        Me.Controls.Add(Me.LblRePrint)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.GridInvHandlerInv)
        Me.Controls.Add(Me.LblPreviewInv)
        Me.Controls.Add(Me.LblSearchInv)
        Me.Controls.Add(Me.TxtInvHandlerInv2Search)
        Me.Controls.Add(Me.CmbAL)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "InvHandler"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :.  Invoice Handler"
        CType(Me.GridInvHandlerInv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents GridInvHandlerInv As System.Windows.Forms.DataGridView
    Friend WithEvents LblPreviewInv As System.Windows.Forms.LinkLabel
    Friend WithEvents LblSearchInv As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtInvHandlerInv2Search As System.Windows.Forms.TextBox
    Friend WithEvents CmbAL As System.Windows.Forms.ComboBox
    Friend WithEvents LckLblDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents LblRePrint As System.Windows.Forms.LinkLabel
End Class
