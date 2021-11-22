<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FoxRefund
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
        Me.GridRef = New System.Windows.Forms.DataGridView
        CType(Me.GridRef, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridRef
        '
        Me.GridRef.AllowUserToAddRows = False
        Me.GridRef.AllowUserToDeleteRows = False
        Me.GridRef.BackgroundColor = System.Drawing.Color.PaleTurquoise
        Me.GridRef.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridRef.Location = New System.Drawing.Point(1, 34)
        Me.GridRef.Name = "GridRef"
        Me.GridRef.RowHeadersVisible = False
        Me.GridRef.Size = New System.Drawing.Size(620, 306)
        Me.GridRef.TabIndex = 0
        '
        'FoxRefund
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(621, 341)
        Me.Controls.Add(Me.GridRef)
        Me.Location = New System.Drawing.Point(0, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FoxRefund"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Airlines :: RAS :. Fox RV"
        CType(Me.GridRef, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GridRef As System.Windows.Forms.DataGridView
End Class
