Public Class SCB
    Dim myPath As String = Application.StartupPath
    Private mstrBankName As String
    Public Sub New(strBankName As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mstrBankName = strBankName
        Me.Text = "TransViet Travel :: RAS 12 :." & strBankName & " Straight2Bank"
    End Sub
    Private Sub SCB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadGridBatch()
        LoadCmb_VAL(Me.CmbAcctToExp, "select a.recID as VAL, shortname + ': ' + LEFT(accountnumber,3)  as DIS " & _
                     " from unc_Company c inner join UNC_Accounts a on c.RecID =a.CompanyID and a.LstUser ='" _
                     & mstrBankName & "'")
        LoadGridUNC("OK")
    End Sub
    Private Sub LoadGridBatch()
        Dim strQuerry As String = "select Distinct SCBNo from UNC_Payments" _
                                  & " where ((status='OK' and SCBNo <>'') " _
                                  & " or (status='E0' and SCBNo='" & mstrBankName & "'))"

        Select Case mstrBankName
            Case "SCB"
                strQuerry = strQuerry & " and substring(SCBNo,1,3)='C00'"
            Case Else
                strQuerry = strQuerry & " and substring(SCBNo,1,3)='" & mstrBankName & "'"
        End Select
        strQuerry = strQuerry & "order by SCBNo"

        Me.GridBatch.DataSource = GetDataTable(strQuerry)
        Me.LblDelete.Visible = False
        Me.GridUNCinBatch.DataSource = Nothing
    End Sub
    Private Sub LoadGridUNC(pStatus As String)
        Try
            Dim StrDKDate As String = "17-Nov-14"
            If Me.CmbAcctToExp.SelectedValue <> 1143 Then StrDKDate = "29-Feb-16"

            If mstrBankName = "VCB" Then
                StrDKDate = "08 Dec 16"
            End If
            Me.GridUNC.DataSource = GetDataTable("select RecID, Amount, ShortName, AccountName, AccountNumber, BankName, BankAddress, Description " & _
                                                 " from UNC_payments where payerAccountID =" & Me.CmbAcctToExp.SelectedValue & _
                                                 " and (scbNo='' or SCBNo='" & mstrBankName _
                                                 & "') and hasPrinted=0 and fstupdate >'" & StrDKDate & _
                                                 "' and status='" & pStatus & "'")
            Me.GridUNC.Columns("recID").Visible = False
            Me.GridUNC.Columns("Amount").Width = 75
            Me.GridUNC.Columns("ShortName").Width = 128
            Me.GridUNC.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.GridUNC.Columns("Amount").DefaultCellStyle.Format = "#,##0.0"
        Catch ex As Exception

        End Try
    End Sub
    Private Sub lblUpdateBatchNo_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdateBatchNo.LinkClicked
        UpdateBatchNo(txtLotNo.Text.ToUpper.Trim)
    End Sub
    Private Function UpdateBatchNo(strBatchNbr As String) As Boolean
        Dim chckLotIDExist As Integer = ScalarToInt("UNC_payments", "RecID" _
                                                    , "payerAccountID=" & Me.CmbAcctToExp.SelectedValue _
                                                    & " and SCBNo='" & strBatchNbr & "'")
        If chckLotIDExist > 0 Then
            MsgBox("Batch No. You Entered Already Exists.", MsgBoxStyle.Critical, msgTitle)
            Exit Function
        End If
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update UNC_payments set status='OK', SCBNo='" & strBatchNbr _
            & "' where payerAccountID=" & Me.CmbAcctToExp.SelectedValue _
            & " and status='E0' and SCBNo='" & mstrBankName & "'"
        cmd.ExecuteNonQuery()

        If mstrBankName.Contains("SCB") Then
            CreateSms4UNC(strBatchNbr, mstrBankName, True)
        End If
        Return True
    End Function
    Private Sub LblExport_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblExport.LinkClicked
        'Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim lstQuerries As New List(Of String)
        Dim strExportFileName As String = String.Empty

        Dim tmpRecID As Integer = ScalarToInt("UNC_Payments", "Top 1 RECID", "status='E0' and SCBNo='' and " & _
                                              "PayerAccountID=" & Me.CmbAcctToExp.SelectedValue)
        If tmpRecID > 0 Then
            MsgBox("Lot No Already Exists. Plz Check Your Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        For i As Int16 = 0 To Me.GridUNC.RowCount - 1
            If Me.GridUNC.Item("S", i).Value = True Then
                lstQuerries.Add("update UNC_payments set status='E0', ScbNo='" _
                                & mstrBankName & "' where recid=" & Me.GridUNC.Item("RecID", i).Value)
            End If
        Next
        If Not (lstQuerries.Count > 0 AndAlso UpdateListOfQuerries(lstQuerries, Conn)) Then
            Exit Sub
        End If
        Select Case mstrBankName
            Case "SCB"
                If MySession.City = "SGN" Then
                    strExportFileName = "R12_2SCB_SGN.xlt"
                Else
                    strExportFileName = "R12_2SCB.xlt"
                End If
                InHoaDon(myPath, strExportFileName, "V", "E0", Now.Date, Now.Date _
                         , Me.CmbAcctToExp.SelectedValue, 0, mstrBankName, "")
                InHoaDon(myPath, strExportFileName, "V", "E0", Now.Date, Now.Date _
                         , Me.CmbAcctToExp.SelectedValue, 0, mstrBankName, "")
            Case "VCB"
                If MySession.City = "SGN" Then
                    strExportFileName = "R12_2VCB_SGN.xlt"
                End If
                InHoaDon(myPath, strExportFileName, "V", "E0", Now.Date, Now.Date _
                         , Me.CmbAcctToExp.SelectedValue, 0, mstrBankName, "")
                UpdateBatchNo(GenerateBankPaymentBatchNbr(Conn, "VCB"))
        End Select
        
    End Sub
    Private Sub ChkPending_Click(sender As Object, e As EventArgs) Handles ChkPending.Click
        If Me.ChkPending.Checked Then
            LoadGridUNC("E0")
        Else
            LoadGridUNC("OK")
        End If
        Me.LblUpdateBatchNo.Enabled = Me.ChkPending.Checked
        Me.TxtLotNo.Enabled = Me.ChkPending.Checked
        Me.LblExport.Enabled = Not Me.ChkPending.Checked
    End Sub

    Private Sub GridBatch_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridBatch.CellClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblDelete.Visible = True
        Dim strDK As String = "status='E0'"
        If Me.GridBatch.CurrentRow.Cells("SCBNo").Value <> "" Then
            strDK = "status='OK' and SCBNo='" & Me.GridBatch.CurrentRow.Cells("SCBNo").Value & "'"
        End If
        Me.GridUNCinBatch.DataSource = GetDataTable("Select RecID, TRX_TC, Curr, Amount, InvNo, RMKNoibo, RefNo, ShortName as Beneficiary," & _
                                                     "AccountName, AccountNumber, BankName, BankAddress, PayerACcountID, Charge, Description," & _
                                                     "Swift, FstUpdate from UNC_payments where " & strDK )
        Me.GridUNCinBatch.Columns("Curr").Width = 32
        Me.GridUNCinBatch.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridUNCinBatch.Columns("Amount").DefaultCellStyle.Format = "#,##0.00"
    End Sub

    Private Sub LblDelete_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        If Me.GridBatch.CurrentRow.Cells("SCBNo").Value = "" Then
            cmd.CommandText = ChangeStatus_ByDK("UNC_Payments", "OK", "status='E0'")
        Else
            cmd.CommandText = "update Unc_payments set SCBNo='' where scbNo='" & Me.GridBatch.CurrentRow.Cells("SCBNo").Value & "'"
            If myStaff.City = "SGN" And InStr("HXT_NMH_SYS", myStaff.SICode) = 0 Then Exit Sub
            If myStaff.City = "HAN" And InStr("NTH_SYS", myStaff.SICode) = 0 Then Exit Sub
        End If
        cmd.ExecuteNonQuery()
        LoadGridBatch()
    End Sub

    Private Sub GridUNC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridUNC.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Dim KQ As Decimal = 0
        Me.txtAmount.Focus()
        If e.ColumnIndex = 0 Then
            For r As Int16 = 0 To Me.GridUNC.RowCount - 1
                If GridUNC.Item(0, r).Value Then
                    KQ = KQ + GridUNC.Item("Amount", r).Value
                End If
            Next
        End If
        Me.txtAmount.Text = Format(KQ, "#,##0.00")
    End Sub

    Private Sub CmbAcctToExp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbAcctToExp.SelectedIndexChanged
        LoadGridUNC("OK")
    End Sub

    
End Class
