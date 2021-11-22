<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NewsUpdater
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
        Me.GridNew = New System.Windows.Forms.DataGridView()
        Me.TxtPic1 = New System.Windows.Forms.TextBox()
        Me.CmbPos = New System.Windows.Forms.ComboBox()
        Me.TxtFull = New System.Windows.Forms.RichTextBox()
        Me.TxtSubj = New System.Windows.Forms.TextBox()
        Me.TxtPic2 = New System.Windows.Forms.TextBox()
        Me.txtPic0 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblUpdate = New System.Windows.Forms.LinkLabel()
        Me.LblRemove = New System.Windows.Forms.LinkLabel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TxtPic3 = New System.Windows.Forms.TextBox()
        Me.cmdPict0 = New System.Windows.Forms.Button()
        Me.cmdPict3 = New System.Windows.Forms.Button()
        Me.cmdPict2 = New System.Windows.Forms.Button()
        Me.cmdPict1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ChkALLACO = New System.Windows.Forms.CheckBox()
        Me.txtHTML = New MSDN.Html.Editor.HtmlEditorControl()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.LblUploadFile = New System.Windows.Forms.LinkLabel()
        Me.LblUploadPict = New System.Windows.Forms.LinkLabel()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        CType(Me.GridNew, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridNew
        '
        Me.GridNew.AllowUserToAddRows = False
        Me.GridNew.AllowUserToDeleteRows = False
        Me.GridNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GridNew.BackgroundColor = System.Drawing.Color.SkyBlue
        Me.GridNew.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridNew.Location = New System.Drawing.Point(4, 5)
        Me.GridNew.Name = "GridNew"
        Me.GridNew.RowHeadersVisible = False
        Me.GridNew.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridNew.Size = New System.Drawing.Size(987, 510)
        Me.GridNew.TabIndex = 0
        '
        'TxtPic1
        '
        Me.TxtPic1.Location = New System.Drawing.Point(685, 4)
        Me.TxtPic1.Name = "TxtPic1"
        Me.TxtPic1.ReadOnly = True
        Me.TxtPic1.Size = New System.Drawing.Size(146, 20)
        Me.TxtPic1.TabIndex = 1
        '
        'CmbPos
        '
        Me.CmbPos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbPos.FormattingEnabled = True
        Me.CmbPos.Items.AddRange(New Object() {"News", "Promo", "FareSheet", "UserGuide"})
        Me.CmbPos.Location = New System.Drawing.Point(350, 4)
        Me.CmbPos.Name = "CmbPos"
        Me.CmbPos.Size = New System.Drawing.Size(67, 21)
        Me.CmbPos.TabIndex = 3
        '
        'TxtFull
        '
        Me.TxtFull.Location = New System.Drawing.Point(48, 26)
        Me.TxtFull.MaxLength = 256
        Me.TxtFull.Name = "TxtFull"
        Me.TxtFull.Size = New System.Drawing.Size(608, 51)
        Me.TxtFull.TabIndex = 2
        Me.TxtFull.Text = ""
        '
        'TxtSubj
        '
        Me.TxtSubj.Location = New System.Drawing.Point(48, 4)
        Me.TxtSubj.MaxLength = 100
        Me.TxtSubj.Name = "TxtSubj"
        Me.TxtSubj.Size = New System.Drawing.Size(254, 20)
        Me.TxtSubj.TabIndex = 1
        '
        'TxtPic2
        '
        Me.TxtPic2.Location = New System.Drawing.Point(685, 26)
        Me.TxtPic2.Name = "TxtPic2"
        Me.TxtPic2.ReadOnly = True
        Me.TxtPic2.Size = New System.Drawing.Size(146, 20)
        Me.TxtPic2.TabIndex = 1
        '
        'txtPic0
        '
        Me.txtPic0.Enabled = False
        Me.txtPic0.Location = New System.Drawing.Point(481, 4)
        Me.txtPic0.Name = "txtPic0"
        Me.txtPic0.Size = New System.Drawing.Size(145, 20)
        Me.txtPic0.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Subject"
        '
        'LblUpdate
        '
        Me.LblUpdate.AutoSize = True
        Me.LblUpdate.Location = New System.Drawing.Point(950, 29)
        Me.LblUpdate.Name = "LblUpdate"
        Me.LblUpdate.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdate.TabIndex = 5
        Me.LblUpdate.TabStop = True
        Me.LblUpdate.Text = "Update"
        '
        'LblRemove
        '
        Me.LblRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblRemove.AutoSize = True
        Me.LblRemove.Location = New System.Drawing.Point(3, 520)
        Me.LblRemove.Name = "LblRemove"
        Me.LblRemove.Size = New System.Drawing.Size(47, 13)
        Me.LblRemove.TabIndex = 5
        Me.LblRemove.TabStop = True
        Me.LblRemove.Text = "Remove"
        Me.LblRemove.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(28, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Brief"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(304, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Position"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(419, 7)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "PicGlance"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(656, 7)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(28, 13)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Pic1"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(656, 50)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(28, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Pic3"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(656, 29)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(28, 13)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "Pic2"
        '
        'TxtPic3
        '
        Me.TxtPic3.Location = New System.Drawing.Point(685, 47)
        Me.TxtPic3.Name = "TxtPic3"
        Me.TxtPic3.ReadOnly = True
        Me.TxtPic3.Size = New System.Drawing.Size(146, 20)
        Me.TxtPic3.TabIndex = 1
        '
        'cmdPict0
        '
        Me.cmdPict0.Location = New System.Drawing.Point(632, 3)
        Me.cmdPict0.Name = "cmdPict0"
        Me.cmdPict0.Size = New System.Drawing.Size(24, 21)
        Me.cmdPict0.TabIndex = 7
        Me.cmdPict0.Text = "..."
        Me.cmdPict0.UseVisualStyleBackColor = True
        '
        'cmdPict3
        '
        Me.cmdPict3.Location = New System.Drawing.Point(834, 46)
        Me.cmdPict3.Name = "cmdPict3"
        Me.cmdPict3.Size = New System.Drawing.Size(22, 21)
        Me.cmdPict3.TabIndex = 7
        Me.cmdPict3.Text = "..."
        Me.cmdPict3.UseVisualStyleBackColor = True
        '
        'cmdPict2
        '
        Me.cmdPict2.Location = New System.Drawing.Point(834, 25)
        Me.cmdPict2.Name = "cmdPict2"
        Me.cmdPict2.Size = New System.Drawing.Size(22, 21)
        Me.cmdPict2.TabIndex = 7
        Me.cmdPict2.Text = "..."
        Me.cmdPict2.UseVisualStyleBackColor = True
        '
        'cmdPict1
        '
        Me.cmdPict1.Location = New System.Drawing.Point(834, 3)
        Me.cmdPict1.Name = "cmdPict1"
        Me.cmdPict1.Size = New System.Drawing.Size(22, 21)
        Me.cmdPict1.TabIndex = 7
        Me.cmdPict1.Text = "..."
        Me.cmdPict1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ChkALLACO
        '
        Me.ChkALLACO.AutoSize = True
        Me.ChkALLACO.Location = New System.Drawing.Point(862, 6)
        Me.ChkALLACO.Name = "ChkALLACO"
        Me.ChkALLACO.Size = New System.Drawing.Size(87, 17)
        Me.ChkALLACO.TabIndex = 8
        Me.ChkALLACO.Text = "CountryWide"
        Me.ChkALLACO.UseVisualStyleBackColor = True
        '
        'txtHTML
        '
        Me.txtHTML.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtHTML.InnerText = Nothing
        Me.txtHTML.Location = New System.Drawing.Point(-4, 79)
        Me.txtHTML.Name = "txtHTML"
        Me.txtHTML.Size = New System.Drawing.Size(993, 451)
        Me.txtHTML.TabIndex = 9
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(-1, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1000, 562)
        Me.TabControl1.TabIndex = 11
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.LblUploadFile)
        Me.TabPage1.Controls.Add(Me.LblUploadPict)
        Me.TabPage1.Controls.Add(Me.txtHTML)
        Me.TabPage1.Controls.Add(Me.cmdPict2)
        Me.TabPage1.Controls.Add(Me.TxtSubj)
        Me.TabPage1.Controls.Add(Me.cmdPict1)
        Me.TabPage1.Controls.Add(Me.TxtPic1)
        Me.TabPage1.Controls.Add(Me.TxtPic2)
        Me.TabPage1.Controls.Add(Me.cmdPict3)
        Me.TabPage1.Controls.Add(Me.txtPic0)
        Me.TabPage1.Controls.Add(Me.cmdPict0)
        Me.TabPage1.Controls.Add(Me.TxtPic3)
        Me.TabPage1.Controls.Add(Me.Label7)
        Me.TabPage1.Controls.Add(Me.TxtFull)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.CmbPos)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.LblUpdate)
        Me.TabPage1.Controls.Add(Me.ChkALLACO)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(992, 536)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Update"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'LblUploadFile
        '
        Me.LblUploadFile.AutoSize = True
        Me.LblUploadFile.Location = New System.Drawing.Point(931, 64)
        Me.LblUploadFile.Name = "LblUploadFile"
        Me.LblUploadFile.Size = New System.Drawing.Size(57, 13)
        Me.LblUploadFile.TabIndex = 11
        Me.LblUploadFile.TabStop = True
        Me.LblUploadFile.Text = "Upload file"
        '
        'LblUploadPict
        '
        Me.LblUploadPict.AutoSize = True
        Me.LblUploadPict.Location = New System.Drawing.Point(862, 50)
        Me.LblUploadPict.Name = "LblUploadPict"
        Me.LblUploadPict.Size = New System.Drawing.Size(62, 13)
        Me.LblUploadPict.TabIndex = 10
        Me.LblUploadPict.TabStop = True
        Me.LblUploadPict.Text = "Upload Pict"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GridNew)
        Me.TabPage2.Controls.Add(Me.LblRemove)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(992, 536)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "View"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'NewsUpdater
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(999, 569)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "NewsUpdater"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: News Updater"
        CType(Me.GridNew, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GridNew As System.Windows.Forms.DataGridView
    Friend WithEvents TxtPic1 As System.Windows.Forms.TextBox
    Friend WithEvents CmbPos As System.Windows.Forms.ComboBox
    Friend WithEvents TxtFull As System.Windows.Forms.RichTextBox
    Friend WithEvents TxtSubj As System.Windows.Forms.TextBox
    Friend WithEvents TxtPic2 As System.Windows.Forms.TextBox
    Friend WithEvents txtPic0 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents LblRemove As System.Windows.Forms.LinkLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtPic3 As System.Windows.Forms.TextBox
    Friend WithEvents cmdPict0 As System.Windows.Forms.Button
    Friend WithEvents cmdPict3 As System.Windows.Forms.Button
    Friend WithEvents cmdPict2 As System.Windows.Forms.Button
    Friend WithEvents cmdPict1 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ChkALLACO As System.Windows.Forms.CheckBox
    Friend WithEvents txtHTML As MSDN.Html.Editor.HtmlEditorControl
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents LblUploadPict As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUploadFile As System.Windows.Forms.LinkLabel
End Class
