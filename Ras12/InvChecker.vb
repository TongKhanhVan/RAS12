Imports SharedFunctions.MySharedFunctionsWzConn
Imports SharedFunctions.MySharedFunctions
Public Class InvChecker
    Dim dTable As DataTable
    Dim strDK_InvTheoThang As String
    Dim DKThang As String, DK_BSP_AL As String
    Dim DKDOI As String
    Private Sub GenCmbThang()
        Dim tmpdate As Date = DateSerial(Now.Year, Now.Month, 1)
        Dim tmpThang As String
        For i As Int16 = -4 To 0
            If tmpdate.AddMonths(1) > CDate("31-Aug-14") Then
                tmpThang = Format(tmpdate.AddDays(i), "MMM-yy")
                Me.CmbMonth.Items.Add(tmpThang)
            End If
        Next
        Me.CmbMonth.Text = Me.CmbMonth.Items(0).ToString
        Me.CmbINVType.Text = "ORG"
    End Sub

    Private Sub InvChecker_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub InvChecker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If HasNewerVersion_R12(Application.ProductVersion) Or SysDateIsWrong(Conn) Then
            Me.Close()
            End
        End If
        GenCmbThang()
        LoadCmb_MSC(CmbAL, myStaff.TVA)
        Me.CmbAL.Text = Me.CmbAL.Items(0).ToString
    End Sub
    Private Sub CleanHoaDonDaHuyRoiDungLai()
        ' cac ban ghi logHd da huy, dung lai thi xoa xx khoi INV, dua vao actionlog
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim StrDK As String = " from INV where status<>'OK' and INVNO in (select INVNO from inv where status='OK')"
        Dim dTbl As DataTable = GetDataTable("select * " & StrDK)
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            cmd.CommandText = UpdateLogFile("INV", "LuuXX", dTbl.Rows(i)("RecID"), dTbl.Rows(i)("INVNO"), dTbl.Rows(i)("SRV"), _
                                            dTbl.Rows(i)("CustID"), dTbl.Rows(i)("CustShortName"), dTbl.Rows(i)("CustAddress"), _
                                            dTbl.Rows(i)("CustTaxCode"), dTbl.Rows(i)("CustFullName"), "RCPID" & dTbl.Rows(i)("RCPID"), _
                                            dTbl.Rows(i)("FstUser") & "/" & dTbl.Rows(i)("FstUpdate"), _
                                            dTbl.Rows(i)("LstUser") & "/" & dTbl.Rows(i)("LstUpdate"))
            cmd.ExecuteNonQuery()
        Next
        cmd.CommandText = "delete " & StrDK
        cmd.ExecuteNonQuery()
    End Sub

    Private Sub LoadGridHD()
        Dim TTLINVAmt As Decimal = 0
        CleanHoaDonDaHuyRoiDungLai()
        Me.LblUpdatePickUp.Visible = False
        Me.LblUpdateInfor.Visible = False
        Me.LblUpdateTKT.Visible = False
        Dim strSelect As String = "select RecID, PickUp, HCopy, InvNo, SRV, status, Amount, CustShortName, CustAddress, CustTaxCode," & _
            "CustFullName, PrintCopy, RCPID, fstUpdate as INVDate, FOP, '' as TRX "
        Dim strDK As String = "  left(invNo,2)='" & Me.CmbAL.Text & "' "
        If Me.CmbAL.Text = "TS" Then
            strDK = strDK & DK_BSP_AL
        End If
        strDK = strDK & " and RecID in " & strDK_InvTheoThang ' lay all inv thoa man

        If Me.CmbINVType.Text = "ORG" Then
            strDK = strDK & " and status='OK' and RCPID>0 "
        ElseIf Me.CmbINVType.Text = "RSV" Then
            If Not DK_BSP_AL Is Nothing Then
                strDK = strDK.Replace(DK_BSP_AL, "")
            End If
            strDK = strDK & " and (RCPID =-64 or status='XX')"
        End If
        dTable = GetDataTable(strSelect & " from INV where " & strDK & " order by INVNO")



        Me.GridHD.DataSource = dTable
        Me.GridHD.Columns(0).Visible = False
        Me.GridHD.Columns("RCPID").Visible = False
        Me.GridHD.Columns(1).Width = 32
        Me.GridHD.Columns(2).Width = 32
        Me.GridHD.Columns(4).Width = 32
        Me.GridHD.Columns(5).Width = 75
        Me.GridHD.Columns("PrintCopy").Width = 50
        Me.GridHD.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridHD.Columns(5).DefaultCellStyle.Format = "#,##0.00"
        Me.GridHD.Columns("INVDate").DefaultCellStyle.Format = "ddMMMyy"
        If Me.CmbINVType.Text = "ORG" Then
            Me.GridHD.Columns("SRV").ReadOnly = True
            If Now.Month > Month(CDate("01-" & Me.CmbMonth.Text)) Or Now.Year > Year(CDate("01-" & Me.CmbMonth.Text)) Then
                DienSoHDBiLungThanhReserve()
            End If
        ElseIf Me.CmbINVType.Text = "RSV" Then
            Me.GridHD.Columns("SRV").ReadOnly = False
        End If
        Me.LblAdj.Visible = False
        Me.LblView.Visible = False

        Dim tblRCP As DataTable
        For i As Int16 = 0 To Me.GridHD.RowCount - 1
            If Me.GridHD.Item("RCPID", i).Value > 0 Then
                tblRCP = GetDataTable("select RCPNO, FstUpdate, ROE from RCP where RecID=" & Me.GridHD.Item("RCPID", i).Value)
                Me.GridHD.Item("InvDate", i).Value = tblRCP.Rows(0)("FstUpdate")
                Me.GridHD.Item("TRX", i).Value = tblRCP.Rows(0)("RCPNo")
            End If
            Me.GridHD.Item("InvNo", i).Value = Me.GridHD.Item("InvNo", i).Value.ToString.Substring(2, 2) & "/" & _
                Me.GridHD.Item("InvNo", i).Value.ToString.Substring(4, 2) & "T" & Me.GridHD.Item("InvNo", i).Value.ToString.Substring(6)
            If Me.GridHD.Item("SRV", i).Value = "R" Or Me.GridHD.Item("SRV", i).Value = "C" Then Me.GridHD.Item("Amount", i).Value = Me.GridHD.Item("Amount", i).Value * -1
            Me.GridHD.Item("Amount", i).Value = Math.Round(Me.GridHD.Item("Amount", i).Value, 0)
            If Me.GridHD.Item("Status", i).Value = "XX" Then Me.GridHD.Item("Amount", i).Value = 0

            If Me.CmbINVType.Text = "ALL" Then
                TTLINVAmt = TTLINVAmt + Me.GridHD.Item("Amount", i).Value
                Me.TxtTTLINVByS1.Text = Format(TTLINVAmt, "#,##0")
            End If
        Next
        Me.GridHD.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridHD.Columns("Amount").DefaultCellStyle.Format = "#,##0"
        Me.LblRounding.Visible = False
    End Sub
    Private Function GetNextInvNo(pINVNo As String) As String
        Dim KQ As String = pINVNo.Substring(0, 7)
        Dim i As Integer = Strings.Right(pINVNo, 5)
        i = i + 1
        Return KQ & Format(i, "00000")
    End Function

    Private Sub DienSoHDBiLungThanhReserve()
        If String.IsNullOrEmpty(Me.CmbAL.Text) Then Exit Sub
        Dim CheckLungHD As Integer = ScalarToInt("ActionLog", "RecID", "TableName='CHCKLUNGHD' and DoWhat='" & Me.CmbAL.Text & Me.CmbMonth.Text & "'")
        If CheckLungHD > 0 Then Exit Sub
        Dim KyHieu As String = ScalarToString("Airline", "KyHieu", "AL='" & Me.CmbAL.Text & "'")
        Dim INVNo As String, INV_cuoi As String, strDK As String, PrevInvDate As Date
        strDK = " left(INVno,2)='" & Me.CmbAL.Text & "' and ((rcpid<0  and " & DKThang & _
            ") or (rcpid in (select rcpid from tkt where " & DKDOI & ")))"

        Dim InvID As Integer
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        INVNo = ScalarToString("INV", "INVNO", strDK & " order by InvNo")
        INV_cuoi = ScalarToString("INV", "INVNO", strDK & " order by InvNo desc")


        On Error Resume Next
        cmd.CommandText = "drop table #tmpINVNO_ChckLung"
        cmd.ExecuteNonQuery()
        On Error GoTo 0

        cmd.CommandText = "select top 1 INVNo into #tmpINVNO_ChckLung from INV where recid=0"
        cmd.ExecuteNonQuery()

        Do While INVNo <> INV_cuoi
            cmd.CommandText = "insert #tmpINVNO_ChckLung (INVNO) values ('" & INVNo & "')"
            cmd.ExecuteNonQuery()
            INVNo = GetNextInvNo(INVNo)
        Loop

        dTable = GetDataTable("select INVNO from #tmpINVNO_ChckLung where invNo not in (select INVNO from INV)")
        For i As Int16 = 0 To dTable.Rows.Count - 1
            INVNo = dTable.Rows(i)("INVNO")
            InvID = Insert_INV("E", INVNo, INVNo.Substring(0, 2), -64)
            PrevInvDate = ScalarToDate("INV", "fstUpdate", "invno='" & GetNextInvNo(INVNo) & "'")
            cmd.CommandText = "update INV set srv='V', custID=697, CustShortName='TEST', Amount=0, CustFullName='', FstUpdate='" & _
                Format(PrevInvDate, "dd-MMM-yy") & "' where recID=" & InvID
            cmd.ExecuteNonQuery()
        Next
        cmd.CommandText = UpdateLogFile("CHCKLUNGHD", Me.CmbAL.Text & Me.CmbMonth.Text, "", "", "", "", "")
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub GridHD_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridHD.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblUpdatePickUp.Visible = True
        Me.LblUpdateInfor.Visible = True
        If Me.CmbINVType.Text = "ORG" Then Me.LblRounding.Visible = True
        Me.LblView.Visible = True
        If Me.CmbINVType.Text = "RSV" Then
            Me.LblAdj.Visible = True
            Me.LblUpdateTKT.Visible = True
        End If
    End Sub
    Private Sub LblUpdate1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdatePickUp.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.Parameters.Add("@PickUp", SqlDbType.Int)
        cmd.Parameters.Add("@HCopy", SqlDbType.Int)
        For i As Int16 = 0 To dTable.Rows.Count - 1
            If dTable.Rows(i).RowState = DataRowState.Modified Then
                cmd.CommandText = UpdateLogFile("INV", "PICKUP", dTable.Rows(i)("RecID").ToString, dTable.Rows(i)("PickUp").ToString, "", "", "") & _
                    "; Update INV set PickUp=@Pickup, HCopy=@HCopy where recid=" & dTable.Rows(i)("RecID")
                cmd.Parameters("@PickUp").Value = dTable.Rows(i)("PickUp")
                cmd.Parameters("@HCopy").Value = dTable.Rows(i)("HCopy")
                cmd.ExecuteNonQuery()
            End If
        Next
    End Sub
    Private Sub CmbAL_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbAL.SelectedIndexChanged
        Dim vThang As Date = CDate("01-" & Me.CmbMonth.Text)
        If Me.CmbAL.Text <> "TS" Then
            Me.CmbBSPAL.Items.Clear()
            Me.CmbBSPAL.Items.Add(Me.CmbAL.Text)
            Me.CmbBSPAL.Enabled = False
        Else
            LoadCmb_MSC(CmbBSPAL, "select AL as VAL from airline where vat like '%TY'")
            Me.CmbBSPAL.Enabled = True
        End If
        Me.CmbBSPAL.Text = Me.CmbBSPAL.Items(0).ToString
    End Sub

    Private Sub CmbMonth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbMonth.SelectedIndexChanged
        Dim vDauThang As Date = CDate("01-" & Me.CmbMonth.Text)
        Dim vCuoiThang As Date = vDauThang.AddMonths(1)
        vCuoiThang = vCuoiThang.AddDays(-1)
        Dim strDK As String = " between '" & Format(vDauThang, "dd-MMM-yy") & "' and '" & Format(vCuoiThang, "dd-MMM-yy") & " 23:59'"

        DKThang = " fstupdate " & strDK
        DKDOI = " DOI " & strDK
        strDK_InvTheoThang = " (select RecID from INV where " & _
                                    "RCPID in (select RCPID from TKT where substring(rcpno,3,4)='" & _
                                            Format(vDauThang, "MMyy") & "' and left(tkno,3) not in ('GRP','MCO','HTL','CAR','INS','AHC') )" & _
                                    " OR (" & DKThang & " and RCPID=-64)" & _
                                    " OR (" & DKThang & " and status<>'OK')" & _
                                ")"
    End Sub
    Private Sub CalculateDiff()
        Dim Amt As Decimal, tmpStr As String
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim strDK As String = "tableName='InvChecker' and DoWhat='Amt2AL' and F1='" & Me.CmbAL.Text & "' and F2='" & Me.CmbMonth.Text & "'"
        tmpStr = ScalarToString("ActionLog", "F3", strDK)
        If tmpStr = "" Then tmpStr = "0"
        Amt = CDec(tmpStr)
        Me.txtADM_VND.Text = Format(Amt, "#,##0")

        tmpStr = ScalarToString("ActionLog", "F4", strDK)
        If tmpStr = "" Then tmpStr = "0"
        Amt = CDec(tmpStr)
        Me.Txt2AL_VND.Text = Format(Amt, "#,##0")

        cmd.CommandText = "select isnull(sum(F_VND),0) as FareToAL from TKTNO_INVNO  " & _
            " where status='OK' and  left(INVNO,2)='" & Me.CmbAL.Text & "' and INVID in " & strDK_InvTheoThang
        If Me.CmbAL.Text = "TS" Then
            cmd.CommandText = cmd.CommandText & DK_BSP_AL
        End If
        Amt = cmd.ExecuteScalar
        Amt = Math.Round(Amt, 0)
        Me.TxtInvFare.Text = Format(Amt, "#,##0")

        'Me.TxtComm_VND.Text = Format(amt) / (1 - CInt(Me.TxtComm.Text) / 100) - amt, "#,##0")
        Me.TxtComm_VND.Text = Format(Amt * CInt(Me.TxtComm.Text) / 100, "#,##0")

        cmd.CommandText = "select isnull(sum (T_VND+C_VND),0) as ToAL from TKTNO_INVNO  " & _
            " where status='OK' and left(INVNO,2)='" & Me.CmbAL.Text & "' and invid in " & strDK_InvTheoThang
        If Me.CmbAL.Text = "TS" Then
            cmd.CommandText = cmd.CommandText & DK_BSP_AL
        End If
        Amt = cmd.ExecuteScalar
        Amt = Math.Round(Amt, 0)
        Me.TxtInvTaxCharge.Text = Format(Amt, "#,##0")

        Me.TxtTTLINV.Text = Format(CDec(Me.TxtInvFare.Text) + CDec(Me.TxtInvTaxCharge.Text), "#,##0")

        Me.TxtDiff_VND.Text = Format(CDec(Me.TxtTTLINV.Text) - CDec(Me.TxtComm_VND.Text) - CDec(Me.txtADM_VND.Text) - CDec(Me.Txt2AL_VND.Text), "#,##0")

    End Sub
    Private Sub LblCalc_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCalc.LinkClicked
        CalculateDiff()
    End Sub

    Private Sub LblView_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblView.LinkClicked
        Dim RPCID As Integer = Me.GridHD.CurrentRow.Cells("RCPID").Value
        Me.GridRCP.DataSource = GetDataTable("select * from rcp where recid=" & RPCID)
        Me.GridTKT.DataSource = GetDataTable("select * from tkt where rcpid=" & RPCID)
        Me.GridFOP.DataSource = GetDataTable("select * from fop where rcpid=" & RPCID)
        Me.TabControl1.SelectTab("TabPage2")
    End Sub

    Private Sub CmbBSPAL_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbBSPAL.SelectedIndexChanged
        DK_BSP_AL = " and RCPID in (select RecID from RCP where Stock='" & Me.CmbBSPAL.Text & "')"
    End Sub

    Private Sub LblAdj_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAdj.LinkClicked
        Dim dTbl As DataTable = GetDataTable("select TKNO, Fare, Tax, Charge from TKTNO_INVNO where status='OK' and  INVID=" & Me.GridHD.CurrentRow.Cells("RecID").Value)
        If dTbl.Rows.Count = 0 Then
            Me.txtFare.Text = 0
            Me.TxtTax.Text = 0
            Me.TxtRTG.Text = ScalarToString("TKT", "top 1 Itinerary", "AL='" & Me.CmbBSPAL.Text & "' order by NetToAL desc")
            Me.txtTKNO.Text = ScalarToString("Airline", "DocCode", "AL='" & Me.CmbBSPAL.Text & "'") & " " & Format(Now, "MMyy") & " " & Format(Now, "ddHHmm")
            Me.CmbCurr.Text = "VND"
        Else
            Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
            Try
                cmd.CommandText = "Drop Table #INV"
                cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            cmd.CommandText = "select * into #INV from ActionLog where tableName='INV' and DoWhat='RSVINV'"
            cmd.ExecuteNonQuery()
            Dim tblINVInfor As DataTable = GetDataTable("select * from #INV where f5 <>'XX' and F1=" & _
                                                        Me.GridHD.CurrentRow.Cells("RecID").Value)
            Me.txtFare.Text = dTbl.Rows(0)("Fare")
            Me.TxtTax.Text = dTbl.Rows(0)("tax")
            Me.TxtRTG.Text = tblINVInfor.Rows(0)("F2")
            Me.CmbCurr.Text = tblINVInfor.Rows(0)("F3")
            Me.TxtVND.Text = tblINVInfor.Rows(0)("F4")
            Me.txtTKNO.Text = dTbl.Rows(0)("TKNO")
        End If
        Me.TabControl1.SelectTab("TabPage3")
    End Sub

    Private Sub LblUpdateTKT_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdateTKT.LinkClicked
        If Me.CmbINVType.Text <> "RSV" Then Exit Sub

        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim INVID As Integer = Me.GridHD.CurrentRow.Cells("RecID").Value
        Dim INVNO As String = ScalarToString("INV", "INVNO", " RecID=" & INVID)
        Dim TKNO As String = Me.txtTKNO.Text
        Dim HeSo As Int16 = IIf(Me.GridHD.CurrentRow.Cells("SRV").Value = "R", -1, 1)
        Dim ActionLogID As Integer = ScalarToInt("ActionLog", "recID", "dowhat='RSVINV' and F1='" & INVID.ToString.Trim & "'")
        If ActionLogID > 0 Then
            cmd.CommandText = "update Actionlog set F5='XX' where RecID=" & ActionLogID
            cmd.ExecuteNonQuery()
        End If
        cmd.CommandText = ChangeStatus_ByDK("TKTNO_INVNO", "XX", "INVID=" & INVID) & _
            ";" & UpdateLogFile("INV", "RSVINV", Me.GridHD.CurrentRow.Cells("RecID").Value, Me.TxtRTG.Text, Me.CmbCurr.Text, CDec(Me.TxtVND.Text) * HeSo, "") & _
            "; update INV set Amount=" & CDec(Me.TxtVND.Text) & ", FOP='" & Me.CmbMonth.Text & "' where recid=" & INVID
        cmd.ExecuteNonQuery()
        For i As Int16 = 1 To CInt(Me.txtQty.Text)
            cmd.CommandText = "insert TKTNO_INVNO (RCPID, TKNO, INVNO, FstUser, Fare, Tax, INVID, F_VND, T_VND) values (-64,'" & TKNO & _
                    "','" & INVNO & "','" & myStaff.SICode & "'," & CDec(Me.txtFare.Text) * HeSo & "," & CDec(Me.TxtTax.Text) * HeSo & _
                    "," & CDec(Me.txtFare.Text) * HeSo & "," & CDec(Me.TxtTax.Text) * HeSo & "," & INVID & ")"
            cmd.ExecuteNonQuery()
            TKNO = DefineNextTKNO(TKNO, False)
        Next

    End Sub
    Private Sub CmbINVType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbINVType.SelectedIndexChanged
        LoadGridHD()
        If Me.CmbINVType.Text <> "ORG" Then
            Me.LblUpdatePickUp.Enabled = False
        Else
            Me.LblUpdatePickUp.Enabled = True
        End If
    End Sub

    Private Sub TxtVND_Enter(sender As Object, e As EventArgs) Handles TxtVND.Enter
        Dim aa As Decimal = CDec(Me.txtFare.Text) + CDec(Me.TxtTax.Text)
        aa = aa * CInt(Me.txtQty.Text)
        Me.TxtVND.Text = Format(aa, "#,##0")
    End Sub
    Private Sub LblLoadGridTRXwoINV_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblLoadGridTRXwoINV.LinkClicked
        Dim vDauThang As Date = CDate("01-" & Me.CmbMonth.Text)
        Dim strDk As String = Me.CmbAL.Text & Format(vDauThang, "MMyy")
        Dim strSQL As String = ""
        strSQL = "select RecID, RCPNo, CustID, CustShortName, TTLDue, Currency, ROE from RCP where ttldue<>0 and " & _
            " left(rcpno,6)='" & strDk & "' and status='OK'  and SRV <>'V' and RecID not in (select RCPID from INV where status='OK')" & _
            " and RecID not in (select RCPID from tkt where left(tkno,3) in ('GRP','MCO','HTL','AHC','INS','CAR'))"
        If Me.CmbAL.Text = "TS" Then
            strSQL = strSQL & " and RecID In (select RCPID from TKT where substring(TKNO,5,4) in (select val from MISC where cat='BSPSTOCK'))"
            strSQL = strSQL & " and Stock In (select AL from Airline where VAT like '%TY')"
        End If
        Me.GridTRXwoINV.DataSource = GetDataTable(strSQL)
        Me.LblIssueINV.Visible = False
    End Sub

    Private Sub GridTRXwoINV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridTRXwoINV.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblIssueINV.Visible = True
    End Sub

    Private Sub LblIssueINV_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblIssueINV.LinkClicked
        Dim f As New InvoicePrinting("BO", Me.GridTRXwoINV.CurrentRow.Cells("RCPNO").Value)
        f.ShowDialog()
    End Sub

    Private Sub txtADM_VND_LostFocus(sender As Object, e As EventArgs) Handles txtADM_VND.LostFocus, Txt2AL_VND.LostFocus
        Dim txt As TextBox = CType(sender, TextBox)
        If txt.Text.Trim = "" Then txt.Text = "0"
        txt.Text = Format(CDec(txt.Text), "#,##0")
        Me.TxtDiff_VND.Text = Format(CDec(Me.TxtTTLINV.Text) - CDec(Me.TxtComm_VND.Text) - CDec(Me.txtADM_VND.Text) - CDec(Me.Txt2AL_VND.Text), "#,##0")
    End Sub
    Private Sub LblRounding_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblRounding.LinkClicked
        Dim CurrAmt As Decimal = ScalarToDec("INV", "Amount", "RecID=" & Me.GridHD.CurrentRow.Cells("RecID").Value)
        Dim InvID As Integer = Me.GridHD.CurrentRow.Cells("RecID").Value
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        CurrAmt = CurrAmt + CDec(Me.CmbRoundingAmt.Text)
        cmd.CommandText = UpdateLogFile("INV", "Rounding", InvID.ToString, Me.CmbRoundingAmt.Text, "", "", "") & _
            "; Update INV set AMount=" & CurrAmt & " where recid=" & InvID
        cmd.ExecuteNonQuery()
        LoadGridHD()
    End Sub

    Private Sub LblRounding_VisibleChanged(sender As Object, e As EventArgs) Handles LblRounding.VisibleChanged
        Me.CmbRoundingAmt.Visible = Me.LblRounding.Visible
    End Sub

    Private Sub LblUpdateInfor_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdateInfor.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.Parameters.Add("@SRV", SqlDbType.VarChar)
        cmd.Parameters.Add("@CustTaxCode", SqlDbType.VarChar)
        cmd.Parameters.Add("@CustFullName", SqlDbType.NVarChar)
        cmd.Parameters.Add("@CustAddress", SqlDbType.NVarChar)
        For i As Int16 = 0 To dTable.Rows.Count - 1
            If dTable.Rows(i).RowState = DataRowState.Modified Then
                cmd.CommandText = UpdateLogFile("INV", "INFOR", dTable.Rows(i)("RecID").ToString, "", _
                    dTable.Rows(i)("SRV"), dTable.Rows(i)("CustFullName"), dTable.Rows(i)("CustAddress"), dTable.Rows(i)("CustTaxCode")) & _
                    "; Update INV set SRV=@SRV, CustFullName=@CustFullName, CustAddress=@CustAddress, CustTaxCode=@CustTaxCode " & _
                    "where recID=" & dTable.Rows(i)("RecID")
                cmd.Parameters("@SRV").Value = dTable.Rows(i)("SRV")
                cmd.Parameters("@CustTaxCode").Value = dTable.Rows(i)("CustTaxCode")
                cmd.Parameters("@CustFullName").Value = dTable.Rows(i)("CustFullName")
                cmd.Parameters("@CustAddress").Value = dTable.Rows(i)("CustAddress")
                cmd.ExecuteNonQuery()
            End If
        Next
        CalculateDiff()
    End Sub
    Private Sub Label10_DoubleClick(sender As Object, e As EventArgs) Handles Label10.DoubleClick
        Me.TxtVND.ReadOnly = False
    End Sub
    Private Sub LblSaveAmtToAL_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSaveAmtToAL.LinkClicked
        Dim strDK As String = "tableName='InvChecker' and DoWhat='Amt2AL' and F1='" & Me.CmbAL.Text & "' and F2='" & Me.CmbMonth.Text & "'"
        Dim chckID As Integer = ScalarToInt("ActionLog", "RecID", strDK)
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        If chckID = 0 Then
            cmd.CommandText = UpdateLogFile("InvChecker", "Amt2AL", Me.CmbAL.Text, Me.CmbMonth.Text, Me.txtADM_VND.Text, Me.Txt2AL_VND.Text, "")
        Else
            cmd.CommandText = "update Actionlog set F3='" & Me.txtADM_VND.Text & "',F4='" & Me.Txt2AL_VND.Text & "', F5='" & myStaff.SICode & _
                "', F6='" & Now.ToShortDateString & "' where " & strDK
        End If
        cmd.ExecuteNonQuery()
    End Sub
End Class