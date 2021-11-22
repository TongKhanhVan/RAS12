Public Class Payment4VCR
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private tmpDaTra As Decimal, Curr As String
    Private StrSQL As String
    Private Sub Payment4VCR_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub LoadGridDEB()
        Me.GridDEB.DataSource = GetDataTable("Select RecID, FOP, Curr, ROE, Amount, Payer, FstUser from flx_FOP " & _
                                             " where status='QQ' and FOP='DEB' and left(payer,5)='[VCR]'")
        Me.GridDEB.Columns(0).Visible = False
        Me.GridDEB.Columns("FOP").Width = 32
        Me.GridDEB.Columns("Curr").Width = 32
        Me.GridDEB.Columns("ROE").Width = 32
        Me.GridDEB.Columns("FstUser").Width = 32
        Me.GridDEB.Columns("Amount").Width = 75
        Me.GridDEB.Columns("Payer").Width = 150
        Me.GridDEB.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridDEB.Columns("Amount").DefaultCellStyle.Format = "#,##0"
        Me.LblClearDEB.Visible = False
    End Sub
    Private Sub Payment4VCR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        GenComboValue()
        LoadGridDEB()

    End Sub
    Private Sub GenComboValue()
        Dim strDK As String = " as VAL from discountVCR where VCRNO<>'' and status ='QQ'"
        Me.FOP.DataSource = GetDataTable("Select VAL from misc where status='OK' and cat='FOP' and val not in ('VCR','PSP','MCO','PTA','EXC','UCF','C/N','MCE','3RD','PPD')")
        Me.FOP.ValueMember = "VAL"
        Me.Currency.DataSource = GetDataTable("Select VAL from misc where status='OK' and cat='Curr' order by vAL")
        Me.Currency.ValueMember = "VAL"
        Me.GridFOP.Rows.Clear()
        strSQL = "select distinct left(VCRNO,3) " & strDK & " and substring(vcrno,4,1) like '[0-9]' UNION select distinct left(VCRNO,4) " & _
            strDK & " and substring(vcrno,4,1) like '[A-Z]'"
        LoadCmb_MSC(Me.CmbPrj, strSQL)
        LoadCmb_VAL(Me.CmbVCRNo, "select RecID as VAL, VCRNo as DIS from discountVCR where VCRNO<>'' and status ='QQ'")
    End Sub
    Private Sub GridFOP_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridFOP.CellEnter
        Dim R As Integer, varROE As Decimal, varCurr As String, BasedCurr As String
        Dim i As Integer = 0
        BasedCurr = Curr
        R = e.RowIndex
        Me.LblUpdate.Visible = False
        For i = Me.GridFOP.RowCount - 1 To R + 1 Step -1
            Me.GridFOP.Item(3, i).Value = 0
            Me.GridFOP.Item(4, i).Value = ""
        Next
        If e.ColumnIndex = 1 Then
            If InStr("DEB_PSP", Me.GridFOP.Item(0, R).Value) > 0 Then
                Me.GridFOP.Item(1, R).Value = Curr
                Me.GridFOP.Item(1, R).ReadOnly = True
            Else
                Me.GridFOP.Item(1, R).ReadOnly = False
            End If
        ElseIf e.ColumnIndex = 2 Then
            varCurr = Me.GridFOP.Item(1, R).Value
            varROE = IIf(varCurr = "VND", 1, ForEX_12(Now.Date, varCurr, "BSR", "YY"))
            Me.GridFOP.Item(2, R).Value = varROE
            Me.GridFOP.Item(3, R).Value = 0
        ElseIf e.ColumnIndex = 3 Then
            If R = 0 Then
                Me.TxtPaid.Text = 0
                If Me.GridFOP.Item(3, R).Value = 0 Then
                    Me.GridFOP.Item(3, R).Value = CDec(Me.txtInVND.Text) / Me.GridFOP.Item(2, R).Value
                End If
                Exit Sub
            End If
            If R > 0 Then
                tmpDaTra = 0
                For i = 0 To R - 1
                    tmpDaTra = tmpDaTra + Me.GridFOP.Item(3, i).Value * Me.GridFOP.Item(2, i).Value
                Next
            End If
            Me.TxtPaid.Text = Format(tmpDaTra, "#,##0.00")
            Me.GridFOP.Item(3, R).Value = (CDec(Me.TxtINVND.Text) - tmpDaTra) / Me.GridFOP.Item(2, R).Value
        ElseIf e.ColumnIndex = 4 Then
            tmpDaTra = 0
            For i = 0 To R
                tmpDaTra = tmpDaTra + Me.GridFOP.Item(3, i).Value * Me.GridFOP.Item(2, i).Value
            Next
            Me.TxtPaid.Text = Format(tmpDaTra, "#,##0.00")
        End If
        Me.GridFOP.Columns(2).DefaultCellStyle.Format = "#,##0.00"
        Me.GridFOP.Columns(3).DefaultCellStyle.Format = "#,##0.00"
    End Sub
    Private Sub GridFOP_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles GridFOP.CellValidating

        If e.ColumnIndex = 3 Or e.ColumnIndex = 2 Then
            Me.GridFOP.Rows(e.RowIndex).ErrorText = ""
            Dim newInteger As Decimal, ErrText As String = "Invalid Input"
            If GridFOP.Rows(e.RowIndex).IsNewRow Then Return
            If Not Decimal.TryParse(e.FormattedValue.ToString(), newInteger) Then
                e.Cancel = True
                Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
            End If
        End If
        If e.ColumnIndex = 4 Then
            Me.GridFOP.Rows(e.RowIndex).ErrorText = ""
            Dim ErrText As String = "Invalid Document Number Format"
            If GridFOP.Rows(e.RowIndex).IsNewRow Then Return
            If (e.FormattedValue.ToString().Length <> 17 And _
                InStr("CRD", Me.GridFOP.Item(0, e.RowIndex).Value) > 0) Then
                e.Cancel = True
                Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
            End If
            If (e.FormattedValue.ToString().Length < 5 And _
                InStr("BTF_CHQ_ITP_VCR", Me.GridFOP.Item(0, e.RowIndex).Value) > 0) Then
                e.Cancel = True
                Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
            End If
        End If
    End Sub

    Private Sub GridFOP_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles GridFOP.RowsAdded
        Me.GridFOP.Item(0, e.RowIndex).Value = "CSH"
        Me.GridFOP.Item(1, e.RowIndex).Value = "VND"
        Me.GridFOP.Item(2, e.RowIndex).Value = 1
    End Sub

    Private Sub LblSearch_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSearch.LinkClicked
        Dim tmpROE As Decimal, PrjAmt As Decimal
        strSQL = " select recID, VCRNo,Curr, Amount from DiscountVCR where status='QQ' and "
        If Me.OptVCR.Checked Then
            strSQL = strSQL & " VCRNo='" & Me.CmbVCRNo.Text & "'"
        Else
            strSQL = strSQL & " vcrno like '" & Me.CmbPrj.Text & "%'"
        End If
        Me.GrdVCRinPrj.DataSource = GetDataTable(strSQL)
        Me.GrdVCRinPrj.Columns(0).Visible = False
        Me.GrdVCRinPrj.Columns(2).Width = 36
        Me.GrdVCRinPrj.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GrdVCRinPrj.Columns(3).DefaultCellStyle.Format = "#,##0.00"
        For i As Int16 = 0 To Me.GrdVCRinPrj.RowCount - 1
            tmpROE = IIf(Me.GrdVCRinPrj.Item(2, i).Value = "VND", 1, ForEX_12(Now.Date, Me.GrdVCRinPrj.Item(2, i).Value, "BSR", "YY"))
            PrjAmt = PrjAmt + Me.GrdVCRinPrj.Item(3, i).Value * tmpROE
            Me.LblNoOfVCR.Text = "No. Of VCR in Project: " & (i + 1).ToString
        Next
        Me.TxtAmt.Text = Format(PrjAmt, "#,##0.00")
        Me.TxtINVND.Text = PrjAmt
        Me.LblDelete.Visible = IIf(Me.GrdVCRinPrj.RowCount > 0, True, False)
    End Sub
    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        If Me.GrdVCRinPrj.RowCount = 0 Then Exit Sub
        Dim pmtID As Integer = 0
        cmd.CommandTimeout = 512
        Try
            For r As Int16 = 0 To Me.GridFOP.RowCount - 1
                If Me.GridFOP.Item("Amount", r).Value > 0 Or Me.GridFOP.Item("FOP", r).Value = "RND" Then
                    StrSQL = "insert FLX_FOP (OrderID, FOP, Curr, Amount, ROE, Document, pmtID, Payer, status, FstUser) values (-99,'" & _
                    Me.GridFOP.Item("FOP", r).Value & "','" & _
                    Me.GridFOP.Item("Currency", r).Value & "','" & _
                    Me.GridFOP.Item("Amount", r).Value & "','" & _
                    Me.GridFOP.Item("ROE", r).Value & "','" & _
                    Me.GridFOP.Item("Document", r).Value.ToString.Replace("--", "") & "','" & _
                    pmtID & "',N'[VCR]:" & Me.TxtPayer.Text.Replace("--", "") & "/" & _
                        IIf(Me.OptPrj.Checked, Me.CmbPrj.Text, Me.CmbVCRNo.Text) & "','" & _
                    IIf(Me.GridFOP.Item("FOP", r).Value = "DEB", "QQ", "OK") & "','" & myStaff.SICode & "')"
                    cmd.CommandText = StrSQL
                    cmd.ExecuteNonQuery()
                    If pmtID = 0 Then
                        cmd.CommandText = "SELECT top 1 RecID from flx_FOP where left(payer,5)='[VCR]' order by recid desc"
                        pmtID = cmd.ExecuteScalar
                        cmd.CommandText = "update FLX_FOP set pmtid=" & pmtID & " where recid=" & pmtID
                        cmd.ExecuteNonQuery()
                    End If
                End If
            Next
            cmd.CommandText = ""
            For i As Int16 = 0 To Me.GrdVCRinPrj.RowCount - 1
                cmd.CommandText = cmd.CommandText & "; update DiscountVCR set status='OK' where recid=" & Me.GrdVCRinPrj.Item(0, i).Value
            Next
            cmd.CommandText = cmd.CommandText.Substring(1)
            cmd.ExecuteNonQuery()
            MsgBox("Payment Has Been Updated and VCR(s) is Ready 2B Redeemed", MsgBoxStyle.Information, msgTitle)
            Me.LblUpdate.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub
    Private Sub TxtPayer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPayer.TextChanged
        If tmpDaTra = CDec(Me.TxtINVND.Text) Then
            Me.LblUpdate.Visible = True
        Else
            Me.LblUpdate.Visible = False
        End If
    End Sub

    Private Sub LblUpdate_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblUpdate.VisibleChanged
        Me.LblNotice.Visible = Me.LblUpdate.Visible
    End Sub

    Private Sub OptPrj_CheckedChanged(sender As Object, e As EventArgs) Handles OptPrj.CheckedChanged, OptVCR.CheckedChanged
        Me.CmbPrj.Enabled = Me.OptPrj.Checked
        Me.CmbVCRNo.Enabled = Me.OptVCR.Checked

    End Sub
    Private Sub LblDelete_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        cmd.CommandText = ""
        For i As Int16 = 0 To Me.GrdVCRinPrj.RowCount - 1
            cmd.CommandText = cmd.CommandText & "; update DiscountVCR set status='XX' where recid=" & Me.GrdVCRinPrj.Item(0, i).Value
        Next
        cmd.CommandText = cmd.CommandText.Substring(1)
        cmd.ExecuteNonQuery()
        MsgBox("VCRs of This Project Have Been Deleted", MsgBoxStyle.Information, msgTitle)
        Me.LblDelete.Visible = False
    End Sub

    Private Sub GridDEB_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridDEB.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblClearDEB.Visible = True
    End Sub

    Private Sub LblClearDEB_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblClearDEB.LinkClicked
        Me.LblClearDEB.Visible = False
        cmd.CommandText = "update FLX_FOP set status='OK', payer=payer+'|" & Me.TxtNoteClearDeb.Text.Replace("--", "") & _
            "' where RecID=" & Me.GridDEB.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridDEB()
    End Sub
End Class