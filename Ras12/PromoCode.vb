Public Class PromoCode
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Sub PromoCode_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Conn.Close()
        frmMain.Close()
        frmMain.Dispose()
        End
    End Sub
    Private Sub PromoCode_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        LoadGridPromoCode()
    End Sub
    Private Sub LoadGridPromoCode()
        Dim strSQL As String
        strSQl = "select RecID, PromoCode, Description, ValidFrom, ValidThru, Msg from PromoCode where PromoCode <>'...'"
        If Me.ChkActiveOnly.Checked Then
            strSQL = strSQL & " and Status='OK' and '" & Now.Date & "' between ValidFrom and ValidThru"
        End If
        Me.GridPromoCode.DataSource = GetDataTable(strSQL)
        Me.GridPromoCode.Columns(0).Visible = False
        Me.GridPromoCode.Columns(1).Width = 100
        Me.GridPromoCode.Columns(2).Width = 250
        Me.GridPromoCode.Columns(3).Width = 72
        Me.GridPromoCode.Columns(4).Width = 72
        Me.GridPromoCode.Columns(5).Width = 256
        Me.LblDeactivate.Visible = False
    End Sub
    Private Sub LblViewFull_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblViewFull.LinkClicked
        Me.GridPromoCode.DataSource = GetDataTable("select * from PromoCode where PromoCode <>'...'")
        Me.LblDeactivate.Visible = False
    End Sub

    Private Sub LblAdd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAdd.LinkClicked
        Dim L As Int16, ALList As String = "TV", strSQL As String
        Dim dTbl As DataTable = GetDataTable(myStaff.TVA)
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            ALList = ALList & "_" & dTbl.Rows(i)("VAL")
        Next
        For i As Int16 = 0 To Me.GridPromoCode.RowCount - 1
            If IsDBNull(Me.GridPromoCode.Item(0, i).Value) And Me.GridPromoCode.Item(1, i).Value <> "" Then
                strSQL = "insert into PromoCode(PromoCode, Description, ValidFrom, ValidThru, MSG, fstUser) values ('"
                L = Me.GridPromoCode.Item("PromoCode", i).Value.ToString.Length
                If L > 25 Then L = 25
                strSQL = strSQL & Me.GridPromoCode.Item("PromoCode", i).Value.ToString.Substring(0, L).ToUpper & "','"
                L = Me.GridPromoCode.Item("Description", i).Value.ToString.Length
                If L > 150 Then L = 150
                strSQL = strSQL & Me.GridPromoCode.Item("Description", i).Value.ToString.Substring(0, L) & "','"
                strSQL = strSQL & Me.GridPromoCode.Item("ValidFrom", i).Value & "','"
                strSQL = strSQL & Me.GridPromoCode.Item("ValidThru", i).Value & "','"
                strSQL = strSQL & Me.GridPromoCode.Item("MSG", i).Value & "','"
                strSQL = strSQL & myStaff.SICode & "')"
                cmd.CommandText = strSQL
                If InStr(ALList, Me.GridPromoCode.Item("PromoCode", i).Value.ToString.Substring(0, 2).ToUpper) > 0 Then
                    cmd.ExecuteNonQuery()
                Else
                    MsgBox("PromoCode Should Start With an Airline Code Or [TV]. Action Aborted", MsgBoxStyle.Critical, msgTitle)
                End If
            End If
        Next
        LoadGridPromoCode()
    End Sub

    Private Sub LblDeactivate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDeactivate.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("PromoCode", "XX", Me.GridPromoCode.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        LoadGridPromoCode()
    End Sub
    Private Sub GridPromoCode_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridPromoCode.CellContentClick
        Me.LblDeactivate.Visible = True
    End Sub
    Private Sub LblRefresh_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblRefresh.LinkClicked
        LoadGridPromoCode()
    End Sub
End Class