Public Class FLX_DEBControl
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Sub GridRQ_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridRQ.CellContentClick
        Me.TxtRMK.Enabled = True
    End Sub

    Private Sub TxtRMK_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtRMK.TextChanged
        Me.LblReject.Visible = True
        Me.LblApprove.Visible = True
    End Sub
    Private Sub DEBApproval_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub DEBApproval_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        LoadGridRQ("QQ")
    End Sub
    Private Sub LoadGridRQ(ByVal pDK As String)
        Me.GridRQ.DataSource = GetDataTable("select RecID, F1 as OrderID, F2 as Amount, F3 as CustID from [42.117.5.86].flx.dbo.Actionlog where doWhat='" & pDK & "' and F5='ITP'")
        Me.GridRQ.Columns(0).Visible = False
        Me.GridRQ.Columns(1).Width = 50
        Me.GridRQ.Columns(2).Width = 50
        Me.GridRQ.Columns(3).Width = 50
        Me.GridRQ.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridRQ.Columns(2).DefaultCellStyle.Format = "#,##0.0"
        Me.LblReject.Visible = False
        Me.LblApprove.Visible = False
        Me.TxtRMK.Enabled = False
    End Sub
    Private Sub LblReject_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblReject.LinkClicked
        Reject_Approve("XX")
    End Sub
    Private Sub LblApprove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblApprove.LinkClicked
        Reject_Approve("OK")
    End Sub
    Private Sub Reject_Approve(ByVal pStatus As String)
        cmd.CommandText = "Update [117.42.5.86].flx.dbo.Actionlog set Dowhat='" & pStatus & "', F4='" & myStaff.SICode & "', F6='" & Now.ToString & _
            ", F7='" & Me.TxtRMK.Text.Replace("--", "") & "' where RecID=" & Me.GridRQ.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridRQ("QQ")
    End Sub
End Class