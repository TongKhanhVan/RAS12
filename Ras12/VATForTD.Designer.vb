<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VATForTD
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
        Me.GridOrder = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtVATNo = New System.Windows.Forms.TextBox()
        Me.LblUpdate = New System.Windows.Forms.LinkLabel()
        Me.ChkWzVATOnly = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LblSearch = New System.Windows.Forms.LinkLabel()
        Me.txtOrderID = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        CType(Me.GridOrder, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridOrder
        '
        Me.GridOrder.AllowUserToAddRows = False
        Me.GridOrder.AllowUserToDeleteRows = False
        Me.GridOrder.BackgroundColor = System.Drawing.Color.MintCream
        Me.GridOrder.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridOrder.Location = New System.Drawing.Point(1, 27)
        Me.GridOrder.Name = "GridOrder"
        Me.GridOrder.RowHeadersVisible = False
        Me.GridOrder.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridOrder.Size = New System.Drawing.Size(779, 440)
        Me.GridOrder.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(540, 474)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "VAT No."
        '
        'txtVATNo
        '
        Me.txtVATNo.Location = New System.Drawing.Point(587, 470)
        Me.txtVATNo.Name = "txtVATNo"
        Me.txtVATNo.Size = New System.Drawing.Size(145, 20)
        Me.txtVATNo.TabIndex = 2
        '
        'LblUpdate
        '
        Me.LblUpdate.AutoSize = True
        Me.LblUpdate.Location = New System.Drawing.Point(738, 474)
        Me.LblUpdate.Name = "LblUpdate"
        Me.LblUpdate.Size = New System.Drawing.Size(42, 13)
        Me.LblUpdate.TabIndex = 3
        Me.LblUpdate.TabStop = True
        Me.LblUpdate.Text = "Update"
        '
        'ChkWzVATOnly
        '
        Me.ChkWzVATOnly.AutoSize = True
        Me.ChkWzVATOnly.Location = New System.Drawing.Point(679, 5)
        Me.ChkWzVATOnly.Name = "ChkWzVATOnly"
        Me.ChkWzVATOnly.Size = New System.Drawing.Size(101, 17)
        Me.ChkWzVATOnly.TabIndex = 4
        Me.ChkWzVATOnly.Text = "WzVATNo Only"
        Me.ChkWzVATOnly.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(-2, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "OrderID"
        '
        'LblSearch
        '
        Me.LblSearch.AutoSize = True
        Me.LblSearch.Location = New System.Drawing.Point(111, 7)
        Me.LblSearch.Name = "LblSearch"
        Me.LblSearch.Size = New System.Drawing.Size(41, 13)
        Me.LblSearch.TabIndex = 6
        Me.LblSearch.TabStop = True
        Me.LblSearch.Text = "Search"
        '
        'txtOrderID
        '
        Me.txtOrderID.Location = New System.Drawing.Point(48, 4)
        Me.txtOrderID.Name = "txtOrderID"
        Me.txtOrderID.Size = New System.Drawing.Size(57, 20)
        Me.txtOrderID.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.SeaGreen
        Me.Label3.Location = New System.Drawing.Point(0, 473)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Label3"
        '
        'VATForTD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtOrderID)
        Me.Controls.Add(Me.LblSearch)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ChkWzVATOnly)
        Me.Controls.Add(Me.LblUpdate)
        Me.Controls.Add(Me.txtVATNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GridOrder)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "VATForTD"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: Flex Tour 10 :. VAT For TD"
        CType(Me.GridOrder, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridOrder As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtVATNo As System.Windows.Forms.TextBox
    Friend WithEvents LblUpdate As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkWzVATOnly As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblSearch As System.Windows.Forms.LinkLabel
    Friend WithEvents txtOrderID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
