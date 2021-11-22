Public Class BulkPrint_New
    Private Const fName As String = ""
    Private MyCust As New objCustomer
    Private VN_CRS_Type As String
    Private Function GetSelectedAL() As String
        Dim KQ As String = ""
        For i As Int16 = 0 To LstAL.Items.Count - 1
            If Me.LstAL.GetItemChecked(i) Then
                KQ = KQ & "," & Me.LstAL.Items(i).ToString
            End If
        Next
        If KQ.Length > 2 Then
            KQ = KQ.Substring(1)
            KQ = KQ.Replace(",", "','")
            KQ = "('" & KQ & "')"
        End If
        Return KQ
    End Function
    Private Sub LoadPendingTRX()
        If Me.CmbCounter.Text = "" Then Exit Sub
        Me.LblPreview.Visible = False
        Me.CmbPrintToWhom.Text = "Customer"
        Dim strDK As String = " RecID not in (select TKTID from INV_TVTR where status='OK') and status='OK' and SRV in ('S','R')"
        If Me.txtTCocde.Text <> "" Then
            strDK = strDK & " and RCPID in (select RcpID from FOP where status='OK' and Document='" & Me.txtTCocde.Text.Replace("--", "") & "')"
        End If
        Dim strSQL As String = "Select RecID, RCPID, Charge_D, RCPNO, SRV, TKNO, DOI, Currency as Curr, Fare, NetToAL, Tax, Charge, ChargeTV, 0 as LCL_C, " & _
            "'' as TCode from TKT where " & strDK & " and RCPID in (select RecID from #" & myStaff.SICode & " where CustID=" & _
            Me.cmbCust.SelectedValue & ")"
        Me.GridTKT.DataSource = GetDataTable(strSQL)
        Me.GridTKT.Columns("recID").Visible = False
        Me.GridTKT.Columns("RCPID").Visible = False
        Me.GridTKT.Columns("Charge_D").Visible = False
        Me.GridTKT.Columns("SRV").Width = 32
        Me.GridTKT.Columns("Curr").Width = 32
        Me.GridTKT.Columns("DOI").Width = 56
        Me.GridTKT.Columns("Fare").Width = 64
        Me.GridTKT.Columns("NetToAL").Width = 75
        Me.GridTKT.Columns("Tax").Width = 56
        Me.GridTKT.Columns("Charge").Width = 56
        Me.GridTKT.Columns("ChargeTV").Width = 56
        Me.GridTKT.Columns("LCL_C").Width = 56
        Me.GridTKT.Columns("Fare").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("Fare").DefaultCellStyle.Format = "#,##0.00"
        Me.GridTKT.Columns("NetToAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("NetToAL").DefaultCellStyle.Format = "#,##0.00"
        Me.GridTKT.Columns("Tax").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("Tax").DefaultCellStyle.Format = "#,##0.00"
        Me.GridTKT.Columns("Charge").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("Charge").DefaultCellStyle.Format = "#,##0"
        Me.GridTKT.Columns("ChargeTV").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("ChargeTV").DefaultCellStyle.Format = "#,##0.00"
        Me.GridTKT.Columns("LCL_C").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("LCL_C").DefaultCellStyle.Format = "#,##0"
        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            If Me.GridTKT.Item("Curr", i).Value = "USD" And Me.GridTKT.Item("Charge_D", i).Value.ToString.Contains("VND") Then ' dua SF ve nguyen te
                Me.GridTKT.Item("LCL_C", i).Value = Split_ChargeTV(Me.GridTKT.Item("Charge_D", i).Value, "VND")
                Me.GridTKT.Item("ChargeTV", i).Value = Split_ChargeTV(Me.GridTKT.Item("Charge_D", i).Value, "USD")
            End If
        Next
        If Me.cmbCust.Text = "TVSGN" Then
            Dim RCPID As Integer, TCode As String
            For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
                If Me.GridTKT.Item("TCode", i).Value = "" Then
                    RCPID = Me.GridTKT.Item("RCPID", i).Value
                    TCode = ScalarToString("FOP", "Document", "RCPID=" & RCPID & " and FOP='PSP' and Document <>''")
                    If String.IsNullOrEmpty(TCode) Then TCode = "."
                    For r As Int16 = 0 To Me.GridTKT.Rows.Count - 1
                        If Me.GridTKT.Item("RCPID", r).Value = RCPID Then
                            Me.GridTKT.Item("TCode", r).Value = TCode
                        End If
                    Next
                End If
            Next
        End If
    End Sub
    Private Function Split_ChargeTV(pChargeD As String, pCurr As String) As Decimal
        Dim Amt As Decimal = 0
        For i As Int16 = 0 To pChargeD.Split("|").Length - 1
            If pChargeD.Split("|")(i).Substring(0, 2) = "TV" Then
                If pChargeD.Split("|")(i).Contains(pCurr) Then
                    Amt = Amt + CDec(pChargeD.Split("|")(i).Split(":")(1).Substring(3))
                End If
            End If
        Next
        Return Amt
    End Function
    Private Sub BulkPrint_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub BulkPrint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        Dim tbl As DataTable = GetDataTable("select AL from airline where status='OK' and AL not in ('01','XX') order by AL")
        For i As Int16 = 0 To tbl.Rows.Count - 1
            If tbl.Rows(i)("AL") = "VN" Then
                Me.LstAL.Items.Add(tbl.Rows(i)("AL"))
            Else
                Me.LstAL.Items.Add(tbl.Rows(i)("AL"), True)
            End If
        Next
    End Sub
    Private Sub LblPreview_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblPreview.LinkClicked
        Exit Sub
        Dim InvNoToPrint As String
        Dim NewAmt As Decimal, frm As Int16, Tto As Int16
        Dim RCPNO As String, INVID As Integer, tmpStr As String
        Dim RCPCharge As Decimal, TKTCharge As Decimal, ROE As Decimal
        Dim tblINV As DataTable
        Dim HasTKNO_INVNO_Record As Integer
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand, myAnswer As Int16
        Dim t As SqlClient.SqlTransaction
        Dim CurrTaxCode As String, CurrINVAmt As Decimal
        For i As Int16 = frm To Tto
            HasTKNO_INVNO_Record = ScalarToInt("TKTNO_INVNO", "top 1 RecID", "status='OK' and RCPID=" & Me.GridTKT.Item("RecID", i).Value)
            If HasTKNO_INVNO_Record > 0 Then
                tblINV = GetDataTable("select RecID, INVNO, CustID, CustTaxCode, Amount from INV where RCPID=" & _
                                      Me.GridTKT.Item("RecID", i).Value & " and status='OK'")
                INVID = tblINV.Rows(0)("RecID")
                InvNoToPrint = tblINV.Rows(0)("INVNO")
                CurrTaxCode = tblINV.Rows(0)("CustTaxCode")
                CurrINVAmt = tblINV.Rows(0)("Amount")
                RCPNO = Me.GridTKT.Item("RCPNO", i).Value
                cmd.CommandText = UpdateTblINVHistory(INVID, InvNoToPrint, 1)
                cmd.ExecuteNonQuery()
                'InHoaDon(Application.StartupPath, fName, "O", InvNoToPrint, Now.Date, Now.Date, 0, IIf(Me.OptGSA.Checked, RCPNO.Substring(0, 2), "TS"), IIf(Me.OptGSA.Checked, "GSA", "TVS"), "")
            End If
        Next
        LoadPendingTRX()
        Me.LblPreview.Visible = False
        On Error GoTo 0
        Exit Sub
RollBackT:
        t.Rollback()
        On Error GoTo 0
    End Sub
    Private Sub GridTRX_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridTKT.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblPreview.Visible = True
    End Sub
    Private Sub LblCheckAll_unCheckAll_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCheckAll_unCheckAll_AL.LinkClicked
        If Me.LblCheckAll_unCheckAll_AL.Text = "CheckAll" Then
            Me.LblCheckAll_unCheckAll_AL.Text = "UnCheckAll"
            Check_UnCheckAL(True)
        Else
            Me.LblCheckAll_unCheckAll_AL.Text = "CheckAll"
            Check_UnCheckAL(False)
        End If
    End Sub
    Private Sub Check_UnCheckAL(pCheck As Boolean)
        For i As Int16 = 0 To Me.LstAL.Items.Count - 1
            If Me.LstAL.Items(i).ToString <> "VN" Or Not pCheck Then Me.LstAL.SetItemChecked(i, pCheck)
        Next
    End Sub

    Private Sub CmbCounter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCounter.SelectedIndexChanged
        If Me.CmbCounter.Text = "" Then
            LoadCmb_VAL(Me.cmbCust, "select 0 as VAL, '' as DIS")
            Exit Sub
        End If
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim strDK_TKT As String = GetSelectedAL()
        If strDK_TKT = "" Then Exit Sub
        If strDK_TKT = "('VN')" Then
            Me.CmbPrintToWhom.Enabled = True
        Else
            Me.CmbPrintToWhom.Enabled = False
        End If
        strDK_TKT = " status='OK' and AL in " & strDK_TKT
        Dim strDK_Channel As String = "('CS','LC')"
        If Me.CmbCounter.Text = "TVS" Then strDK_Channel = "('TA','TO','CA','WK')"
        Try
            cmd.CommandText = "drop table #" & myStaff.SICode
            cmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        cmd.CommandText = "select RecID, CustID into #" & myStaff.SICode & _
            " from RCP where recID in (select RCPID from tkt where " & strDK_TKT & " and DOI between'" & _
            Format(Me.txtDOIFrm.Value, "dd-MMM-yy") & "' and '" & Format(Me.TxtDOIto.Value, "dd-MMM-yy") & " 23:56')"
        cmd.ExecuteNonQuery()
        LoadCmb_VAL(Me.cmbCust, "select RecID as VAL, CustShortName as DIS from customerList where RecID>0 and status='OK' and " & _
                     " recID in (select CustID from Cust_detail where status='OK' and cat='Channel' and VAL in " & strDK_Channel & ")" & _
                     " and recID in (select CustID from #" & myStaff.SICode & ")")
    End Sub
    Private Sub LblLoadTKT_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblLoadTKT.LinkClicked
        LoadPendingTRX()
        Me.TabControl1.SelectTab("TabPage2")
    End Sub

    Private Sub cmbCust_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCust.SelectedIndexChanged
        Me.LblLoadTKT.Visible = True
    End Sub
    Private Function DefineVN_CRS_Type(pTKNO As String) As String
        If ScalarToInt("MISC", "count(*)", " cat='BSPSTOCK' and val='" & pTKNO.Substring(4, 4) & "'") <> 0 Then Return "BSP"
        Return "NonBSP"
    End Function

    Private Sub GridTRX_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles GridTKT.CellValueChanged
        If e.RowIndex = -1 Then Exit Sub
        Dim Tcode As String = Me.GridTKT.Item("TCode", e.RowIndex).Value
        Dim chkStatus As Boolean = Me.GridTKT.Item("S", e.RowIndex).Value
        If Me.cmbCust.Text = "TVSGN" AndAlso Tcode <> "." Then
            For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
                If Me.GridTKT.Item("TCode", i).Value = Tcode Then Me.GridTKT.Item("S", i).Value = chkStatus
            Next
        End If

    End Sub

    Private Sub GridTRX_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles GridTKT.CurrentCellDirtyStateChanged
        Me.GridTKT.CommitEdit(DataGridViewDataErrorContexts.Commit)
    End Sub
    Private Sub LBlChk_UnChkAll_TKT_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LBlChk_UnChkAll_TKT.LinkClicked
        If Me.LBlChk_UnChkAll_TKT.Text = "CheckAll" Then
            Me.LBlChk_UnChkAll_TKT.Text = "UnCheckAll"
            Check_UnCheckTKT(True)
        Else
            Me.LBlChk_UnChkAll_TKT.Text = "CheckAll"
            Check_UnCheckTKT(False)
        End If
    End Sub
    Private Sub Check_UnCheckTKT(pStatus As Boolean)
        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            Me.GridTKT.Item("S", i).Value = pStatus
        Next
    End Sub
    Private Sub txtDOIFrm_ValueChanged(sender As Object, e As EventArgs) Handles txtDOIFrm.ValueChanged, TxtDOIto.ValueChanged
        Me.CmbCounter.Text = ""
    End Sub

    Private Sub LstAL_LostFocus(sender As Object, e As EventArgs) Handles LstAL.LostFocus
        Me.CmbCounter.Text = ""
    End Sub

    Private Sub LblCalcRASInvAmt_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCalcRASInvAmt.LinkClicked
        Me.LblPreview.Enabled = False
        Dim VNDTotal As Decimal, ROE As Decimal, FTC As Decimal
        VN_CRS_Type = ""
        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            If Me.GridTKT.Item("S", i).Value Then
                If VN_CRS_Type = "" Then
                    VN_CRS_Type = DefineVN_CRS_Type(Me.GridTKT.Item("TKNO", i).Value)
                Else
                    If VN_CRS_Type <> DefineVN_CRS_Type(Me.GridTKT.Item("TKNO", i).Value) Then
                        Me.txtSummary.Text = "Err. Selected TKTs Issued By Different CRS"
                        Exit Sub
                    End If
                End If
                ROE = 1
                If Me.GridTKT.Item("Curr", i).Value <> "VND" Then ROE = ScalarToDec("RCP", "ROE", "RecID=" & Me.GridTKT.Item("RCPID", i).Value)
                FTC = Me.GridTKT.Item("Tax", i).Value + Me.GridTKT.Item("Charge", i).Value + Me.GridTKT.Item("ChargeTV", i).Value

                If VN_CRS_Type = "BSP" Then
                    FTC = FTC + Me.GridTKT.Item("Fare", i).Value
                Else
                    FTC = FTC + Me.GridTKT.Item("NetToAL", i).Value
                End If
                VNDTotal = VNDTotal + FTC * ROE + Me.GridTKT.Item("LCL_C", i).Value
            End If
        Next
        Me.txtSummary.Text = IIf(VN_CRS_Type = "NonBSP", "Net2AL", "Fare") & "-Based Total: " & Format(VNDTotal, "#,##0")
        Me.LblPreview.Enabled = True
    End Sub
End Class