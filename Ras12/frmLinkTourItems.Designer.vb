<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLinkTourItems
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
        Me.grdItems = New System.Windows.Forms.DataGridView()
        Me.lbkOK = New System.Windows.Forms.LinkLabel()
        Me.lbkCancel = New System.Windows.Forms.LinkLabel()
        CType(Me.grdItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grdItems
        '
        Me.grdItems.AllowUserToAddRows = False
        Me.grdItems.AllowUserToDeleteRows = False
        Me.grdItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdItems.Location = New System.Drawing.Point(0, 0)
        Me.grdItems.Name = "grdItems"
        Me.grdItems.ReadOnly = True
        Me.grdItems.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdItems.Size = New System.Drawing.Size(496, 325)
        Me.grdItems.TabIndex = 0
        '
        'lbkOK
        '
        Me.lbkOK.AutoSize = True
        Me.lbkOK.Location = New System.Drawing.Point(157, 340)
        Me.lbkOK.Name = "lbkOK"
        Me.lbkOK.Size = New System.Drawing.Size(22, 13)
        Me.lbkOK.TabIndex = 1
        Me.lbkOK.TabStop = True
        Me.lbkOK.Text = "OK"
        '
        'lbkCancel
        '
        Me.lbkCancel.AutoSize = True
        Me.lbkCancel.Location = New System.Drawing.Point(238, 340)
        Me.lbkCancel.Name = "lbkCancel"
        Me.lbkCancel.Size = New System.Drawing.Size(40, 13)
        Me.lbkCancel.TabIndex = 2
        Me.lbkCancel.TabStop = True
        Me.lbkCancel.Text = "Cancel"
        '
        'frmLinkTourItems
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 473)
        Me.Controls.Add(Me.lbkCancel)
        Me.Controls.Add(Me.lbkOK)
        Me.Controls.Add(Me.grdItems)
        Me.Name = "frmLinkTourItems"
        Me.Text = "LinkTourItems"
        CType(Me.grdItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grdItems As System.Windows.Forms.DataGridView
    Friend WithEvents lbkOK As System.Windows.Forms.LinkLabel
    Friend WithEvents lbkCancel As System.Windows.Forms.LinkLabel
End Class
