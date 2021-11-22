Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.Crd_Ctrl
Public Class BGUpdate
    Private CharEntered As Boolean = False
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private MyCust As New objCustomer

    Private Sub BGUpdate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
        MyCust.CustID = 0
    End Sub

    Private Sub BGUpdate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Me.BackColor = pubVarBackColor
        MyCust.GenCustList()
        genComboValue()
        LoadGridBG()
        CheckRightForALLForm(Me)

    End Sub

    Private Sub genComboValue()
        LoadCmb_VAL(Me.CmbCustomer, MyCust.List_CR)
        LoadCmb_MSC(Me.CmbBank, "Bank")
        LoadCmb_MSC(Me.cmbTVC, "select Description as VAL from misc where cat='TVCompany'")
    End Sub
    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblUpdate.LinkClicked
        Me.LckLblUpdate.Visible = False
        Dim tmpContractValid As String = Format(Me.txtContractValidUntil.Value, "dd-MMM-yy")
        If Not Me.txtContractValidUntil.Checked Then tmpContractValid = "01-Jan-2000"
        cmd.CommandText = "insert into BG (CustId, CustShortName, ContractNo, ContractValidity, ContractOrgCopy, " & _
                "BGAMount, BGValidFrm, BGExpireDate, Bank, BGOrgCopy, Remark, TVCompany, FstUser) values (" & _
                "@CustId, @CustShortName, @ContractNo, @ContractValidity, @ContractOrgCopy, " & _
                "@BGAMount, @BGValidFrm, @BGExpireDate, @Bank, @BGOrgCopy, @Remark, @TVCompany, @FstUser)"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@CustId", SqlDbType.Int).Value = Me.CmbCustomer.SelectedValue
        cmd.Parameters.Add("@CustShortName", SqlDbType.VarChar).Value = Me.CmbCustomer.Text
        cmd.Parameters.Add("@ContractNo", SqlDbType.VarChar).Value = Me.TxtContractNo.Text
        cmd.Parameters.Add("@ContractValidity", SqlDbType.VarChar).Value = tmpContractValid
        cmd.Parameters.Add("@ContractOrgCopy", SqlDbType.VarChar).Value = Me.CmbContractOrgCopy.Text
        cmd.Parameters.Add("@BGAMount", SqlDbType.Decimal).Value = CDec(Me.TxtAmount.Text)
        cmd.Parameters.Add("@BGValidFrm", SqlDbType.DateTime).Value = Me.TxtValidFrm.Value.Date
        cmd.Parameters.Add("@BGExpireDate", SqlDbType.DateTime).Value = Me.TxtValidThru.Value.Date
        cmd.Parameters.Add("@Bank", SqlDbType.VarChar).Value = Me.CmbBank.Text
        cmd.Parameters.Add("@BGOrgCopy", SqlDbType.VarChar).Value = "Acct"
        cmd.Parameters.Add("@Remark", SqlDbType.VarChar).Value = Me.txtNote.Text
        cmd.Parameters.Add("@TVCompany", SqlDbType.VarChar).Value = Me.cmbTVC.Text
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.ExecuteNonQuery()
        Me.TxtAmount.Text = 0
        Me.TxtContractNo.Text = ""
        Me.LckLblUpdate.Visible = False
        MyCust.CustID = Me.CmbCustomer.SelectedValue
        RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "Update BG_CRD", Conn, myStaff.SICode, CnStr)
        LoadGridBG()
    End Sub
    Private Sub LoadGridBG()
        Dim strSQL As String
        Me.LblRenewInProc.Visible = False
        Me.LblLock.Visible = False
        Me.LckLblApprove.Visible = False

        strSQL = "Select RecID, CustID, CustShortName, BGAmount as Amount, BGValidFrm as ValidFrm, BGExpireDate as ExpireDate, status, "
        If Me.OptBG.Checked Then
            strSQL = strSQL & " Bank, "
        End If
        strSQL = strSQL & "ContractNo, ContractValidity, ContractOrgCopy, TVCompany, FstUser, FstUpdate "
        If Me.ChkActiveOnly.Checked Then
            strSQL = strSQL & " from BG where status='OK' and '" & Now.Date & "' between bGValidFrm and BGExpireDate"
        Else
            strSQL = strSQL & ", status, LstUser, LstUpdate from BG  where (status <>'OK' or not ('" & Now.Date & "' between bGValidFrm and BGExpireDate))"
        End If
        If Me.ChkSelectCustOnly.Checked Then
            strSQL = strSQL & " and  custID=" & Me.CmbCustomer.SelectedValue
        End If
        If Me.OptBG.Checked Then
            strSQL = strSQL & " and bank not in ('CRD','CSH')"
        ElseIf Me.OptCredit.Checked Then
            strSQL = strSQL & " and bank ='CRD' "
        Else
            strSQL = strSQL & " and bank ='CSH' "
        End If
        Me.GridBG.DataSource = GetDataTable(strSQL)
        Me.GridBG.Columns("RecID").Visible = False
        Me.GridBG.Columns("CustID").Width = 40
        Me.GridBG.Columns("FstUser").Width = 40
        Me.GridBG.Columns("Amount").Width = 70
        Me.GridBG.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridBG.Columns("Amount").DefaultCellStyle.Format = "#,##0"
        Me.LblDelete.Visible = False
        Me.txtContractValidUntil.Checked = False

        On Error Resume Next
        Me.GridBG.Columns("LstUser").Width = 40
        Me.GridBG.Columns("Bank").Width = 40
        Me.GridBG.Columns("Status").Width = 40
        On Error GoTo 0

    End Sub
    Private Sub TxtAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAmount.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub

    Private Sub TxtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAmount.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub TxtAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAmount.LostFocus
        Dim aa As Decimal = Me.TxtAmount.Text
        Me.TxtAmount.Text = Format(aa, "#,##0")
    End Sub

    Private Sub TxtAmount_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAmount.TextChanged
        Me.LckLblUpdate.Visible = True
    End Sub
    Private Sub CmbCustomer_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCustomer.LostFocus
        LoadGridBG()
        Me.LckLblUpdate.Visible = True
    End Sub
    Private Sub LblExportCredit_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        InHoaDon(Application.StartupPath, "PSP_Credit_Listing.xlt", "F", "", Now.Date, Now.Date, 0, "", "")
    End Sub
    Private Sub LblExportBG_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblExportBG.LinkClicked
        InHoaDon(Application.StartupPath, "BG_Listing.xlt", "F", "", Now.Date, Now.Date, 0, "", "")
    End Sub

    Private Sub LblLock_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblLock.LinkClicked
        Lock_unlock("LK")
    End Sub
    Private Sub Lock_unlock(ByVal pStatus As String)
        Dim tmpRecNo As Integer = Me.GridBG.CurrentRow.Cells("recID").Value
        cmd.CommandText = ChangeStatus_ByID("CC_Setting", pStatus, tmpRecNo) & "; " & _
            UpdateLogFile("CC_Setting", "(un)Lock", tmpRecNo, pStatus, "", "", "", "", "", "")
        cmd.ExecuteNonQuery()
        MyCust.CustID = Me.GridBG.CurrentRow.Cells("CustID").Value
        RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "Del BG_CRD", Conn, myStaff.SICode, CnStr)
    End Sub
    Private Sub LblUnlock_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblUnlock.LinkClicked
        Lock_unlock("OK")
    End Sub

    Private Sub ChkActiveOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkSelectCustOnly.Click
        LoadGridBG()
    End Sub

    Private Sub GridBG_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridBG.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If Me.GridBG.CurrentRow.Cells("Status").Value = "QQ" Then
            Me.LckLblApprove.Visible = True
        Else
            Me.LckLblApprove.Visible = False
        End If
        If Me.ChkActiveOnly.Checked Then
            Me.LblRenewInProc.Visible = Me.OptBG.Checked
            Me.LblLock.Visible = Not Me.OptBG.Checked
            Me.LblDelete.Visible = True
        Else
            Me.LblRenewInProc.Visible = False
            Me.LblLock.Visible = False
            Me.LckLblUnlock.Visible = Not OptBG.Checked
        End If
    End Sub
    Private Sub DeleteBG_CR(ByVal pRecID As Integer)
        cmd.CommandText = ChangeStatus_ByID("BG", "XX", pRecID)
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub LblDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        DeleteBG_CR(Me.GridBG.CurrentRow.Cells("RecID").Value)
        MyCust.CustID = Me.GridBG.CurrentRow.Cells("CustID").Value
        RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "Del BG_CRD", Conn, myStaff.SICode, CnStr)
        LoadGridBG()
    End Sub
    Private Sub LblRenewInProc_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblRenewInProc.LinkClicked
        If Me.GridBG.CurrentRow.Cells("ExpireDate").Value < Now.Date.AddDays(14) Then
            cmd.CommandText = String.Format("Update BG set RenewinProc=1 where recid={0}", Me.GridBG.CurrentRow.Cells("RecID").Value)
            cmd.ExecuteNonQuery()
        End If
        Me.LblRenewInProc.Visible = False
    End Sub

    Private Sub ChkActiveOnly_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkActiveOnly.Click
        LoadGridBG()
    End Sub

    Private Sub OptBG_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptBG.Click, OptDeposit.Click, OptCredit.Click
        Me.CmbBank.Items.Clear()
        If Me.OptCredit.Checked Then
            Me.CmbBank.Items.Add("CRD")
            Me.LblLock.Visible = Me.ChkActiveOnly.Checked
        ElseIf Me.OptDeposit.Checked Then
            Me.CmbBank.Items.Add("CSH")
            Me.LblLock.Visible = Me.ChkActiveOnly.Checked
        Else
            LoadCmb_MSC(Me.CmbBank, "Bank")
        End If
        Me.CmbBank.Text = Me.CmbBank.Items(0).ToString
    End Sub

    Private Sub LckLblApprove_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LckLblApprove.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("BG", "OK", Me.GridBG.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        LoadGridBG()
    End Sub
End Class