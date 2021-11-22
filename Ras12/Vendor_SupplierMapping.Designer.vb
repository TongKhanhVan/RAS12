<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Vendor_SupplierMapping
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
        Me.GridSupplier = New System.Windows.Forms.DataGridView()
        Me.GridVendor = New System.Windows.Forms.DataGridView()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.LblContains = New System.Windows.Forms.Label()
        Me.LblRef = New System.Windows.Forms.Label()
        Me.LblCreateVendor = New System.Windows.Forms.LinkLabel()
        Me.LblMap = New System.Windows.Forms.LinkLabel()
        Me.LblSearch = New System.Windows.Forms.LinkLabel()
        Me.OptSupplierAddrress = New System.Windows.Forms.RadioButton()
        Me.OptAccountNo = New System.Windows.Forms.RadioButton()
        Me.OptVendorName = New System.Windows.Forms.RadioButton()
        Me.OptSupplierName = New System.Windows.Forms.RadioButton()
        Me.TxtSearch = New System.Windows.Forms.TextBox()
        Me.txtSupplierID = New System.Windows.Forms.TextBox()
        CType(Me.GridSupplier, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridVendor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridSupplier
        '
        Me.GridSupplier.AllowUserToAddRows = False
        Me.GridSupplier.AllowUserToDeleteRows = False
        Me.GridSupplier.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridSupplier.Location = New System.Drawing.Point(0, 3)
        Me.GridSupplier.Name = "GridSupplier"
        Me.GridSupplier.RowHeadersVisible = False
        Me.GridSupplier.Size = New System.Drawing.Size(772, 464)
        Me.GridSupplier.TabIndex = 0
        '
        'GridVendor
        '
        Me.GridVendor.AllowUserToAddRows = False
        Me.GridVendor.AllowUserToDeleteRows = False
        Me.GridVendor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridVendor.Location = New System.Drawing.Point(2, 36)
        Me.GridVendor.Name = "GridVendor"
        Me.GridVendor.RowHeadersVisible = False
        Me.GridVendor.Size = New System.Drawing.Size(769, 371)
        Me.GridVendor.TabIndex = 1
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(1, 1)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(780, 496)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GridSupplier)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(772, 470)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "New Supplier"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.txtSupplierID)
        Me.TabPage2.Controls.Add(Me.LblContains)
        Me.TabPage2.Controls.Add(Me.LblRef)
        Me.TabPage2.Controls.Add(Me.LblCreateVendor)
        Me.TabPage2.Controls.Add(Me.LblMap)
        Me.TabPage2.Controls.Add(Me.LblSearch)
        Me.TabPage2.Controls.Add(Me.OptSupplierAddrress)
        Me.TabPage2.Controls.Add(Me.OptAccountNo)
        Me.TabPage2.Controls.Add(Me.OptVendorName)
        Me.TabPage2.Controls.Add(Me.OptSupplierName)
        Me.TabPage2.Controls.Add(Me.TxtSearch)
        Me.TabPage2.Controls.Add(Me.GridVendor)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(772, 470)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Existing Vendor"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'LblContains
        '
        Me.LblContains.AutoSize = True
        Me.LblContains.Location = New System.Drawing.Point(3, 452)
        Me.LblContains.Name = "LblContains"
        Me.LblContains.Size = New System.Drawing.Size(48, 13)
        Me.LblContains.TabIndex = 8
        Me.LblContains.Text = "Contains"
        '
        'LblRef
        '
        Me.LblRef.AutoSize = True
        Me.LblRef.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblRef.ForeColor = System.Drawing.Color.MediumBlue
        Me.LblRef.Location = New System.Drawing.Point(4, 9)
        Me.LblRef.Name = "LblRef"
        Me.LblRef.Size = New System.Drawing.Size(10, 15)
        Me.LblRef.TabIndex = 7
        Me.LblRef.Text = "."
        '
        'LblCreateVendor
        '
        Me.LblCreateVendor.AutoSize = True
        Me.LblCreateVendor.Location = New System.Drawing.Point(571, 452)
        Me.LblCreateVendor.Name = "LblCreateVendor"
        Me.LblCreateVendor.Size = New System.Drawing.Size(158, 13)
        Me.LblCreateVendor.TabIndex = 6
        Me.LblCreateVendor.TabStop = True
        Me.LblCreateVendor.Text = "CreateVendor4SelectedSupplier"
        '
        'LblMap
        '
        Me.LblMap.AutoSize = True
        Me.LblMap.Location = New System.Drawing.Point(571, 415)
        Me.LblMap.Name = "LblMap"
        Me.LblMap.Size = New System.Drawing.Size(188, 13)
        Me.LblMap.TabIndex = 5
        Me.LblMap.TabStop = True
        Me.LblMap.Text = "MapSelectedSuppler2SelectedVendor"
        Me.LblMap.Visible = False
        '
        'LblSearch
        '
        Me.LblSearch.AutoSize = True
        Me.LblSearch.Location = New System.Drawing.Point(229, 452)
        Me.LblSearch.Name = "LblSearch"
        Me.LblSearch.Size = New System.Drawing.Size(41, 13)
        Me.LblSearch.TabIndex = 4
        Me.LblSearch.TabStop = True
        Me.LblSearch.Text = "Search"
        '
        'OptSupplierAddrress
        '
        Me.OptSupplierAddrress.AutoSize = True
        Me.OptSupplierAddrress.Location = New System.Drawing.Point(115, 429)
        Me.OptSupplierAddrress.Name = "OptSupplierAddrress"
        Me.OptSupplierAddrress.Size = New System.Drawing.Size(116, 17)
        Me.OptSupplierAddrress.TabIndex = 3
        Me.OptSupplierAddrress.TabStop = True
        Me.OptSupplierAddrress.Text = "By SupplierAddress"
        Me.OptSupplierAddrress.UseVisualStyleBackColor = True
        '
        'OptAccountNo
        '
        Me.OptAccountNo.AutoSize = True
        Me.OptAccountNo.Location = New System.Drawing.Point(3, 429)
        Me.OptAccountNo.Name = "OptAccountNo"
        Me.OptAccountNo.Size = New System.Drawing.Size(94, 17)
        Me.OptAccountNo.TabIndex = 3
        Me.OptAccountNo.TabStop = True
        Me.OptAccountNo.Text = "By AccountNo"
        Me.OptAccountNo.UseVisualStyleBackColor = True
        '
        'OptVendorName
        '
        Me.OptVendorName.AutoSize = True
        Me.OptVendorName.Location = New System.Drawing.Point(115, 411)
        Me.OptVendorName.Name = "OptVendorName"
        Me.OptVendorName.Size = New System.Drawing.Size(102, 17)
        Me.OptVendorName.TabIndex = 3
        Me.OptVendorName.TabStop = True
        Me.OptVendorName.Text = "By VendorName"
        Me.OptVendorName.UseVisualStyleBackColor = True
        '
        'OptSupplierName
        '
        Me.OptSupplierName.AutoSize = True
        Me.OptSupplierName.Location = New System.Drawing.Point(3, 411)
        Me.OptSupplierName.Name = "OptSupplierName"
        Me.OptSupplierName.Size = New System.Drawing.Size(106, 17)
        Me.OptSupplierName.TabIndex = 3
        Me.OptSupplierName.TabStop = True
        Me.OptSupplierName.Text = "By SupplierName"
        Me.OptSupplierName.UseVisualStyleBackColor = True
        '
        'TxtSearch
        '
        Me.TxtSearch.Location = New System.Drawing.Point(51, 447)
        Me.TxtSearch.Name = "TxtSearch"
        Me.TxtSearch.Size = New System.Drawing.Size(172, 20)
        Me.TxtSearch.TabIndex = 2
        '
        'txtSupplierID
        '
        Me.txtSupplierID.Location = New System.Drawing.Point(276, 449)
        Me.txtSupplierID.Name = "txtSupplierID"
        Me.txtSupplierID.Size = New System.Drawing.Size(61, 20)
        Me.txtSupplierID.TabIndex = 9
        Me.txtSupplierID.Text = "0"
        Me.txtSupplierID.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSupplierID.Visible = False
        '
        'Vendor_SupplierMapping
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Vendor_SupplierMapping"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS12 :. Vendor Supplier Mapping"
        CType(Me.GridSupplier, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridVendor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GridSupplier As System.Windows.Forms.DataGridView
    Friend WithEvents GridVendor As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents LblMap As System.Windows.Forms.LinkLabel
    Friend WithEvents LblSearch As System.Windows.Forms.LinkLabel
    Friend WithEvents OptSupplierAddrress As System.Windows.Forms.RadioButton
    Friend WithEvents OptAccountNo As System.Windows.Forms.RadioButton
    Friend WithEvents OptSupplierName As System.Windows.Forms.RadioButton
    Friend WithEvents TxtSearch As System.Windows.Forms.TextBox
    Friend WithEvents LblCreateVendor As System.Windows.Forms.LinkLabel
    Friend WithEvents LblRef As System.Windows.Forms.Label
    Friend WithEvents OptVendorName As System.Windows.Forms.RadioButton
    Friend WithEvents LblContains As System.Windows.Forms.Label
    Friend WithEvents txtSupplierID As System.Windows.Forms.TextBox
End Class
