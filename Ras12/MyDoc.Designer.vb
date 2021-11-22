<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MyDoc
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
        Me.txtSearch = New System.Windows.Forms.RichTextBox()
        Me.LblSearch = New System.Windows.Forms.LinkLabel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.CmbCat = New System.Windows.Forms.ComboBox()
        Me.TreeRight = New System.Windows.Forms.TreeView()
        Me.WebDisplay = New System.Windows.Forms.WebBrowser()
        Me.LblEmail = New System.Windows.Forms.LinkLabel()
        Me.TxtName = New System.Windows.Forms.TextBox()
        Me.txtMail = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSearch
        '
        Me.txtSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtSearch.Location = New System.Drawing.Point(3, 469)
        Me.txtSearch.Multiline = False
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(312, 25)
        Me.txtSearch.TabIndex = 2
        Me.txtSearch.Text = ""
        '
        'LblSearch
        '
        Me.LblSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblSearch.AutoSize = True
        Me.LblSearch.Location = New System.Drawing.Point(321, 474)
        Me.LblSearch.Name = "LblSearch"
        Me.LblSearch.Size = New System.Drawing.Size(41, 13)
        Me.LblSearch.TabIndex = 3
        Me.LblSearch.TabStop = True
        Me.LblSearch.Text = "Search"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.CmbCat)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TreeRight)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.WebDisplay)
        Me.SplitContainer1.Size = New System.Drawing.Size(776, 466)
        Me.SplitContainer1.SplitterDistance = 140
        Me.SplitContainer1.TabIndex = 3
        '
        'CmbCat
        '
        Me.CmbCat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmbCat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCat.FormattingEnabled = True
        Me.CmbCat.Location = New System.Drawing.Point(1, 5)
        Me.CmbCat.Name = "CmbCat"
        Me.CmbCat.Size = New System.Drawing.Size(136, 21)
        Me.CmbCat.TabIndex = 0
        '
        'TreeRight
        '
        Me.TreeRight.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeRight.CheckBoxes = True
        Me.TreeRight.Location = New System.Drawing.Point(0, 28)
        Me.TreeRight.Name = "TreeRight"
        Me.TreeRight.Size = New System.Drawing.Size(137, 435)
        Me.TreeRight.TabIndex = 1
        '
        'WebDisplay
        '
        Me.WebDisplay.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WebDisplay.Location = New System.Drawing.Point(3, 5)
        Me.WebDisplay.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebDisplay.Name = "WebDisplay"
        Me.WebDisplay.Size = New System.Drawing.Size(619, 456)
        Me.WebDisplay.TabIndex = 7
        '
        'LblEmail
        '
        Me.LblEmail.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblEmail.AutoSize = True
        Me.LblEmail.Location = New System.Drawing.Point(747, 474)
        Me.LblEmail.Name = "LblEmail"
        Me.LblEmail.Size = New System.Drawing.Size(32, 13)
        Me.LblEmail.TabIndex = 6
        Me.LblEmail.TabStop = True
        Me.LblEmail.Text = "Email"
        '
        'TxtName
        '
        Me.TxtName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtName.Location = New System.Drawing.Point(394, 471)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.Size = New System.Drawing.Size(118, 20)
        Me.TxtName.TabIndex = 4
        '
        'txtMail
        '
        Me.txtMail.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMail.ForeColor = System.Drawing.Color.Gray
        Me.txtMail.Location = New System.Drawing.Point(514, 471)
        Me.txtMail.Name = "txtMail"
        Me.txtMail.Size = New System.Drawing.Size(231, 20)
        Me.txtMail.TabIndex = 5
        Me.txtMail.Text = "eMail"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(368, 474)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(20, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "To"
        '
        'MyDoc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtMail)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.LblEmail)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.LblSearch)
        Me.Controls.Add(Me.txtSearch)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MyDoc"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS 12 :. Quick Reference"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtSearch As System.Windows.Forms.RichTextBox
    Friend WithEvents LblSearch As System.Windows.Forms.LinkLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TreeRight As System.Windows.Forms.TreeView
    Friend WithEvents CmbCat As System.Windows.Forms.ComboBox
    Friend WithEvents LblEmail As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtName As System.Windows.Forms.TextBox
    Friend WithEvents txtMail As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents WebDisplay As System.Windows.Forms.WebBrowser
End Class
