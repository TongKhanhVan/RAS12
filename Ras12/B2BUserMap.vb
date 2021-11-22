Public Class B2BUserMap
    Dim cmd_tvvn As SqlClient.SqlCommand = Conn_TVVN.CreateCommand
    Private MyCust As New objCustomer
    Private Sub GridUser_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridUser.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If Me.ChkXXOnly.Checked Then
            Me.LblReinstate.Visible = True
        Else
            If Me.GridUser.CurrentRow.Cells("Status").Value = "QQ" Then
                Me.LblMap.Visible = True
            Else
                Me.LblMap.Visible = False
            End If
            Me.LblDelete.Visible = True
        End If
    End Sub

    Private Sub B2BUserMap_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Conn_TVVN.Close()
        Me.Dispose()
    End Sub

    Private Sub B2BUserMap_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Conn_TVVN.ConnectionString = CnStr_TVVN
        Conn_TVVN.Open()
        LoadGridUser(myStaff.City)
        MyCust.GenCustList()
        LoadCmb_VAL(Me.CmbCust, MyCust.List_TA)
        Me.CmbCust.DisplayMember = "DIS"
    End Sub
    Private Sub LoadGridUser(pCity As String)
        Dim strSQL = "Select RecID, Email, AgentShortName, Status from tblUser where city='" & pCity & "'"
        If Me.ChkXXOnly.Checked Then
            strSQL = strSQL & " and status ='XX'"
        Else
            strSQL = strSQL & " and status <>'XX'"
        End If
        Me.GridUser.DataSource = GetDataTable(strSQL, Conn_TVVN)

        Me.GridUser.Columns(0).Visible = False
        Me.GridUser.Columns(1).Width = 256
        Me.GridUser.Columns("Status").Visible = False
        Me.LblDelete.Visible = False
        Me.LblMap.Visible = False
        Me.LblReinstate.Visible = False
    End Sub

    Private Sub LblMap_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblMap.LinkClicked
        If Me.CmbCust.Text = "" Then Exit Sub
        cmd_tvvn.CommandText = "update tblUser set status='OK', AgentShortName='" & Me.CmbCust.Text & "' where recid=" & _
            Me.GridUser.CurrentRow.Cells(0).Value
        cmd_tvvn.ExecuteNonQuery()
        LoadGridUser(myStaff.City)
    End Sub

    Private Sub LblDelete_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        cmd_tvvn.CommandText = "update tblUser set status='XX', lstUser='" & myStaff.SICode & "', lstUpdate=getdate() where recid=" & _
            Me.GridUser.CurrentRow.Cells(0).Value
        cmd_tvvn.ExecuteNonQuery()
        LoadGridUser(myStaff.City)
    End Sub
    Private Sub ChkXXOnly_Click(sender As Object, e As EventArgs) Handles ChkXXOnly.Click
        LoadGridUser(myStaff.City)
    End Sub
    Private Sub LblReinstate_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblReinstate.LinkClicked
        cmd_tvvn.CommandText = "update tblUser set status='OK' where recid=" & Me.GridUser.CurrentRow.Cells(0).Value
        cmd_tvvn.ExecuteNonQuery()
        LoadGridUser(myStaff.City)
    End Sub
End Class