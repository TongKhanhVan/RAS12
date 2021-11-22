Public Class GSA_MISC
    Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Const DKLostSales As String = " cat='LOSTSALES' and Status='OK' "
    Private MyCust As New objCustomer
    Private Sub GSA_MISC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        MyCust.GenCustList()
        LoadCmb_VAL(Me.CmbCustomer, MyCust.List_TA)
    End Sub
    Private Sub OptCust_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptCust.Click, OptAL.Click
        LoadGridFE()
    End Sub

    Private Sub CmbCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCustomer.SelectedIndexChanged
        LoadGridFE()
    End Sub
    Private Sub LoadGridFE()
        Dim strSQL As String
        If Me.TxtAL.Text = "" Then Exit Sub
        Try
            strSQL = "select RecID, val as CustID, val2 as ShortName, val1 as AL, details as MaxLoss, LstUpdate as ValidThru from MISC " & _
                " where " & DKLostSales
            If Me.OptAL.Checked Then
                strSQL = strSQL & " and val1='" & Me.TxtAL.Text & "'"
            Else
                strSQL = strSQL & " and val=" & Me.CmbCustomer.SelectedValue
            End If
            Me.GridFE.DataSource = GetDataTable(strSQL)
            Me.GridFE.Columns(0).Visible = False
            Me.GridFE.Columns(1).Width = 56
            Me.GridFE.Columns(3).Width = 32
            Me.GridFE.Columns(4).Width = 56
            Me.LblDeleteFrontEnd.Visible = False
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LblUpdateFrontEnd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateFrontEnd.LinkClicked
        Dim tmpRecID As Integer
        tmpRecID = ScalarToInt("MISC", "RecID", DKLostSales & " and val='" & Me.CmbCustomer.SelectedValue & "' and val1='" & Me.TxtAL.Text & "'")
        If tmpRecID > 0 Then
            MsgBox("Setting for this Customer/AL Already Exists", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        cmd.CommandText = "insert MISC (cat, val,val1, details, lstUpdate, FstUser, val2) values ('LOSTSALES'," & Me.CmbCustomer.SelectedValue & ",'" & _
            Me.TxtAL.Text & "','" & Me.TxtMaxLoss.Text & "','" & Format(Me.TxtValidThru.Value, "dd-MMM-yy") & " 23:59','" & myStaff.SICode & "','" & _
            Me.CmbCustomer.Text & "')"
        cmd.ExecuteNonQuery()
        LoadGridFE()
    End Sub

    Private Sub LblDeleteFrontEnd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDeleteFrontEnd.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("MISC", "XX", Me.GridFE.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        LoadGridFE()
    End Sub

    Private Sub GridFE_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridFE.CellContentClick
        Me.LblDeleteFrontEnd.Visible = True
    End Sub
    Private Sub TxtAL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtAL.TextChanged
        LoadGridFE()
    End Sub
End Class