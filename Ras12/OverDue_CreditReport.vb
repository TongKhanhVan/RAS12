Public Class OverDue_CreditReport
    Private Sub OverDue_CreditReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.txtFrm.Value = Now.Date.AddDays(-7)
        LoadGridOverDueCredit()
    End Sub
    Private Sub LoadGridOverDueCredit()
        Dim strDKCustType As String = " and custid in (select custId from cust_detail where status='OK' and cat='Channel' and val in ('TA','TO'))"
        Me.GridOverDue.DataSource = GetDataTable("select CustShortName as CustName, InvAmt as Amount, DueDate from GhiNoKhach " & _
            " where status <>'XX' and conno>0 and duedate <getdate() " & strDKCustType)
        Me.GridOverDue.Columns(2).Width = 108
        Me.GridOverDue.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridOverDue.Columns(1).DefaultCellStyle.Format = "#,##0"
        Me.GridOverLimit.DataSource = GetDataTable("select custID, CustShortName as CustName," & _
                                                   "case PPCoef when 0 then VND_PSP_Avail else VND_PPD_Avail end as CurrBLC," & _
                                                    " getdate() as LstPmt from cc_BLC  " & _
                                                    " where RecID in (Select Max(RecID) From CC_BLC Group by CustID) " & _
                                                    strDKCustType & _
                                                    " and PPCoef * VND_PPD_Avail + CRCoef * VND_PSP_Avail <-1 order by CurrBLC ")
        Me.GridOverLimit.Columns(0).Visible = False
        Me.GridOverLimit.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridOverLimit.Columns(2).DefaultCellStyle.Format = "#,##0"
        For i As Int16 = 0 To Me.GridOverLimit.RowCount - 1
            Me.GridOverLimit.Item("LstPmt", i).Value = ScalarToDate("KhachTra", "top 1 FstUpdate", "custid=" & Me.GridOverLimit.Item(0, i).Value & _
                " and status<>'XX' order by recid desc")
        Next
    End Sub

    Private Sub LblGo_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblGo.LinkClicked
        If Me.CmbUser.Text = "" Then Exit Sub
        Dim Qty As Decimal
        Me.GridUser.Rows.Clear()
        Qty = ScalarToDec("ActionLog", "count(*)", " tableName='CREXT' and DoWhat='OK' and ActionBy='" & Me.CmbUser.Text & _
            "' and ActionDate between '" & Me.txtFrm.Value.Date & "' and '" & Me.txtThru.Value & "'")
        Me.GridUser.Rows.Add()
        Me.GridUser.Item(0, 0).Value = "No Of TRX"
        Me.GridUser.Item(1, 0).Value = Qty

        Qty = ScalarToDec("ActionLog", "count(distinct f3)", " tableName='CREXT' and DoWhat='OK' and ActionBy='" & Me.CmbUser.Text & _
            "' and ActionDate between '" & Me.txtFrm.Value.Date & "' and '" & Me.txtThru.Value & "' ")
        Me.GridUser.Rows.Add()
        Me.GridUser.Item(0, 1).Value = "No Of Agents"
        Me.GridUser.Item(1, 1).Value = Qty

        Qty = ScalarToDec("ActionLog", "isnull(Sum(cast(F5 as dec(12,2))),0)", " tableName='CREXT' and DoWhat='OK' and ActionBy='" & Me.CmbUser.Text & _
            "' and ActionDate between '" & Me.txtFrm.Value.Date & "' and '" & Me.txtThru.Value & "' ")
        Me.GridUser.Rows.Add()

        Me.GridUser.Item(0, 2).Value = "TTL VND Approved"
        Me.GridUser.Item(1, 2).Value = Format(Qty, "#,##0")
    End Sub
End Class