Public Class VATForTD

    Private Sub VATForTD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadGridOrder("")
    End Sub
    Private Sub LoadGridOrder(pDK As String)
        Dim strDK As String = IIf(Me.ChkWzVATOnly.Checked, " VATNo<>''", " VATNo=''")
        If pDK <> "" Then strDK = pDK
        Me.GridOrder.DataSource = GetDataTable("select RecID as OrderID, CustID, CustName, Phone, PmtStatus, TTLDue, VATNo " & _
                                               " from [42.117.5.86].flx.dbo.Ordermanager" & _
                                                " where " & strDK & " and status <>'XX' and FstUpdate>'1-Jul-15' and city='" & myStaff.City & "'")
        Me.GridOrder.Columns(0).Width = 50
        Me.GridOrder.Columns(1).Width = 50
        Me.GridOrder.Columns(2).Width = 200
        Me.GridOrder.Columns(4).Width = 50
        Me.GridOrder.Columns(5).Width = 75

        Me.GridOrder.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridOrder.Columns(5).DefaultCellStyle.Format = "#,##0"
        Me.LblUpdate.Visible = False
        Me.Label3.Text = ""
    End Sub

    Private Sub GridOrder_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridOrder.CellClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblUpdate.Visible = True
        Dim DeptID As Int16 = ScalarToInt("[42.117.5.86].flx.dbo.Booking", "top 1 DeptID", "status='OK' and OrderID=" & Me.GridOrder.CurrentRow.Cells("OrderID").Value)
        Me.Label3.Text = ScalarToString("[42.117.5.86].flx.dbo.Departure", "Departure", "RecID=" & DeptID)
    End Sub

    Private Sub LblUpdate_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update [42.117.5.86].flx.dbo.Ordermanager set VATNo=@VATNo where recID=" & Me.GridOrder.CurrentRow.Cells("OrderID").Value
        cmd.Parameters.Add("@VATNo", SqlDbType.VarChar).Value = Me.txtVATNo.Text
        cmd.ExecuteNonQuery()
        LoadGridOrder("")
    End Sub
    Private Sub ChkAllOrder_Click(sender As Object, e As EventArgs) Handles ChkWzVATOnly.Click
        Me.txtOrderID.Text = ""
        LoadGridOrder("")
    End Sub
    Private Sub LblSearch_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSearch.LinkClicked
        LoadGridOrder("RecID=" & Me.txtOrderID.Text)
    End Sub
End Class