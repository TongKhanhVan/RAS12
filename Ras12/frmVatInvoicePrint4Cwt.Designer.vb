<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVatInvoicePrint4Cwt
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
        Me.cboCustomer = New System.Windows.Forms.ComboBox()
        Me.lbkLoadFile = New System.Windows.Forms.LinkLabel()
        Me.dgrTktListing = New System.Windows.Forms.DataGridView()
        Me.lbkPrint = New System.Windows.Forms.LinkLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtpInvoiceDate = New System.Windows.Forms.DateTimePicker()
        Me.cboSheetName = New System.Windows.Forms.ComboBox()
        CType(Me.dgrTktListing, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboCustomer
        '
        Me.cboCustomer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomer.FormattingEnabled = True
        Me.cboCustomer.Items.AddRange(New Object() {"EY HAN", "EY SGN", "ORACLE"})
        Me.cboCustomer.Location = New System.Drawing.Point(12, 12)
        Me.cboCustomer.Name = "cboCustomer"
        Me.cboCustomer.Size = New System.Drawing.Size(121, 21)
        Me.cboCustomer.TabIndex = 0
        '
        'lbkLoadFile
        '
        Me.lbkLoadFile.AutoSize = True
        Me.lbkLoadFile.Location = New System.Drawing.Point(156, 20)
        Me.lbkLoadFile.Name = "lbkLoadFile"
        Me.lbkLoadFile.Size = New System.Drawing.Size(47, 13)
        Me.lbkLoadFile.TabIndex = 1
        Me.lbkLoadFile.TabStop = True
        Me.lbkLoadFile.Text = "LoadFile"
        '
        'dgrTktListing
        '
        Me.dgrTktListing.AllowUserToAddRows = False
        Me.dgrTktListing.AllowUserToDeleteRows = False
        Me.dgrTktListing.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgrTktListing.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgrTktListing.Location = New System.Drawing.Point(12, 39)
        Me.dgrTktListing.Name = "dgrTktListing"
        Me.dgrTktListing.ReadOnly = True
        Me.dgrTktListing.Size = New System.Drawing.Size(992, 551)
        Me.dgrTktListing.TabIndex = 2
        '
        'lbkPrint
        '
        Me.lbkPrint.AutoSize = True
        Me.lbkPrint.Location = New System.Drawing.Point(189, 601)
        Me.lbkPrint.Name = "lbkPrint"
        Me.lbkPrint.Size = New System.Drawing.Size(28, 13)
        Me.lbkPrint.TabIndex = 3
        Me.lbkPrint.TabStop = True
        Me.lbkPrint.Text = "Print"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 601)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "InvoiceDate"
        '
        'dtpInvoiceDate
        '
        Me.dtpInvoiceDate.CustomFormat = "dd MMM yy"
        Me.dtpInvoiceDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpInvoiceDate.Location = New System.Drawing.Point(83, 594)
        Me.dtpInvoiceDate.Name = "dtpInvoiceDate"
        Me.dtpInvoiceDate.Size = New System.Drawing.Size(91, 20)
        Me.dtpInvoiceDate.TabIndex = 6
        '
        'cboSheetName
        '
        Me.cboSheetName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSheetName.FormattingEnabled = True
        Me.cboSheetName.Location = New System.Drawing.Point(221, 12)
        Me.cboSheetName.Name = "cboSheetName"
        Me.cboSheetName.Size = New System.Drawing.Size(202, 21)
        Me.cboSheetName.TabIndex = 8
        '
        'frmVatInvoicePrint4Cwt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1016, 623)
        Me.Controls.Add(Me.cboSheetName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dtpInvoiceDate)
        Me.Controls.Add(Me.lbkPrint)
        Me.Controls.Add(Me.dgrTktListing)
        Me.Controls.Add(Me.lbkLoadFile)
        Me.Controls.Add(Me.cboCustomer)
        Me.Name = "frmVatInvoicePrint4Cwt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VatInvoicePrint4Cwt"
        CType(Me.dgrTktListing, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboCustomer As System.Windows.Forms.ComboBox
    Friend WithEvents lbkLoadFile As System.Windows.Forms.LinkLabel
    Friend WithEvents dgrTktListing As System.Windows.Forms.DataGridView
    Friend WithEvents lbkPrint As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpInvoiceDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboSheetName As System.Windows.Forms.ComboBox
End Class
