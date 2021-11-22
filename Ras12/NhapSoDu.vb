Imports SharedFunctions.Crd_Ctrl
Imports SharedFunctions.MySharedFunctions
Public Class NhapSoDu
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private CharEntered As Boolean = False
    Private MyCust As New objCustomer
    Private Sub NhapSoDu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MyCust.CustID = 0
        Me.Dispose()
    End Sub
    Private Sub NhapSoDu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        MyCust.CustID = 0
        GenComboValue()
    End Sub
    Private Sub GenComboValue()
        Dim strSQL As String
        LoadCmb_MSC(Me.CmbCNCurr, "Curr")
        Me.CmbCNCurr.Text = "VND"
        strSQL = MyCust.List_CC
        If MySession.Domain = "EDU" Then
            strSQL = strSQL & " and custid <0"
        End If
        LoadCmb_VAL(Me.CmbCustomer, strSQL)
    End Sub

    Private Sub CmbCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCustomer.SelectedIndexChanged
        Try
            MyCust.CustID = Me.CmbCustomer.SelectedValue
            Me.txtCustFullName.Text = MyCust.FullName
            LoadGridSoDu()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub LoadGridSoDu()
        Dim strSQL As String
        Try
            strSQL = String.Format("select * from ChotCongNo where status='OK' and Custid={0}", Me.CmbCustomer.SelectedValue)
            Me.GridSoDu.DataSource = GetDataTable(strSQL)
            Me.GridSoDu.Columns(0).Visible = False
            Me.GridSoDu.Columns(1).Visible = False
            Me.GridSoDu.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.GridSoDu.Columns(4).DefaultCellStyle.Format = "#,##0.00"
            Me.GridSoDu.Columns("FstUSer").Width = 32
            Me.GridSoDu.Columns("LstUSer").Width = 32
        Catch ex As Exception

        End Try
        Me.LblDelete.Visible = False
    End Sub

    Private Sub txtCNAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCNAmount.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub

    Private Sub txtCNAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCNAmount.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtCNAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCNAmount.LostFocus
        Dim aa As Double
        Try
            aa = CDbl(txtCNAmount.Text)
            Me.txtCNAmount.Text = Format(aa, "#,##0.00")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdUpdate.Click
        Dim MyAns As Int16, tmpAmt As Decimal = CDec(Me.txtCNAmount.Text)
        Dim tmpROE As Decimal
        MyAns = MsgBox("Plz Check Your Input Carefully Before Update. Need Further Thought?", MsgBoxStyle.Information Or MsgBoxStyle.YesNo, msgTitle)
        If MyAns = vbYes Then Exit Sub
        If Me.txtCNDescription.Text = "" Or Me.CmbDebType.Text = "" Then
            MsgBox("invalid Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        tmpROE = ForEX_12(Now.Date, Me.CmbCNCurr.Text, "BBR", "YY").Amount
        cmd.CommandText = "Insert ChotCongNo (CustID, CustShortName, AsOf, VND_AVail, FstUser) Values ('" & _
            Me.CmbCustomer.SelectedValue & "','" & Me.CmbCustomer.Text & "','" & _
            Format(Me.TxtAsOf.Value.Date, "dd-MMM-yy") & " 23:59" & "'," & tmpAmt * tmpROE & ",'" & myStaff.SICode & "')"
        cmd.ExecuteNonQuery()

        RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "Aft NhapSoDu " & Me.txtCNAmount.Text, Conn, myStaff.SICode, CnStr)

        MsgBox("Updated", MsgBoxStyle.Information, msgTitle)
        Me.txtCNAmount.Text = 0
        Me.txtCNDescription.Text = ""
    End Sub
    Private Sub CmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdRefresh.Click
        RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "ViewFrm NhapSoDu", Conn, myStaff.SICode, CnStr)
    End Sub

    Private Sub GridSoDu_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridSoDu.CellContentClick
        Me.LblDelete.Visible = True
    End Sub
    Private Sub LblDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("ChotCongNo", "XX", Me.GridSoDu.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        LoadGridSoDu()
    End Sub
End Class
