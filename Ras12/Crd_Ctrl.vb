Imports SharedFunctions.MySharedFunctionsWzConn
Public Class Crd_Ctrl
    Public Shared Function defineVND_Avail(ByVal pCustID As Integer, ByVal parType As String, ByVal parCutOverDate As Date, ByVal parCutOverDateNew As Date, ByVal pRMK As String, ByVal ppConn As SqlClient.SqlConnection, ByVal pSICode As String, ByVal pCnString As String) As Decimal
        Dim UnUse As Decimal = 0, Unpaid As Decimal = 0, UnInvoice As Decimal = 0, KQ As Decimal = 0
        Dim BG As Decimal = 0, TinChap As Decimal = 0, strSQL As String, FT_Sale As Decimal
        Dim CustName As String, SabreSale As Decimal
        UnUse = DefineUnUse(pCustID, parType, parCutOverDateNew, parCutOverDate, ppConn) ' voi PPD thi tong khach deposit deu la unused
        UnInvoice = DefineUnInvoice(pCustID, parType, parCutOverDateNew, parCutOverDate, ppConn) ' PSP chua inv hoac all PPD (xem y tren)
        If parType = "PSP" Then
            BG = DefineBG_TinChap(pCustID, "BG", ppConn)
            TinChap = DefineBG_TinChap(pCustID, "CR", ppConn)
            Unpaid = DefineUnPaid(pCustID, parCutOverDateNew, parCutOverDate, ppConn)
        End If

        FT_Sale = Define_SaleNotInRAS(pCustID, ppConn, "FT")
        SabreSale = Define_SaleNotInRAS(pCustID, ppConn, "'1S")

        Dim cmd As SqlClient.SqlCommand = ppConn.CreateCommand
        cmd.CommandText = "select CustShortName from cc_setting where custid=" & pCustID
        CustName = cmd.ExecuteScalar

        strSQL = "Insert cc_BLC (FstUser, CustShortName, CustID, CRCoef, PPCoef, FT_Sale, SabreSale, "
        If parType = "PSP" Then
            strSQL = strSQL & " PSP_UnUsed, PSP_UnInv, PSP_UnPaid, BG, CR "
        ElseIf parType = "PPD" Then
            strSQL = strSQL & " PPD_Depo, PPD_Used"
        End If
        strSQL = strSQL & ", RMK) select '" & pSICode & "','" & CustName & "', custID, CrCoef, PPCoef," & FT_Sale
        strSQL = strSQL & "," & SabreSale & "," & UnUse & "," & UnInvoice
        If parType = "PSP" Then
            strSQL = strSQL & "," & Unpaid & "," & BG & "," & TinChap
        End If
        strSQL = strSQL & ",'" & pRMK & "' from cc_Setting where status <>'XX' and custid=" & pCustID
        cmd.CommandText = strSQL
        cmd.ExecuteNonQuery()
        If parType = "PSP" Then
            strSQL = "select top 1 VND_PSP_AVail"
        ElseIf parType = "PPD" Then
            strSQL = "select top 1 VND_PPD_AVail"
        End If
        cmd.CommandText = strSQL & " from cc_BLC where custid=" & pCustID & " order by recid desc"
        KQ = cmd.ExecuteScalar

        Return KQ
    End Function
    Private Shared Function Define_SaleNotInRAS(ByVal pCustID As Integer, ByVal ppConn As SqlClient.SqlConnection, ByVal pFT_1S As String) As Decimal
        Dim KQ As Decimal, PRGlist As String
        If pFT_1S = "FT" Then
            PRGlist = "('TTQ','TTP','A1S')"
        Else
            PRGlist = "('M1S')"
        End If
        Dim cmd As SqlClient.SqlCommand = ppConn.CreateCommand
        cmd.CommandText = "select isnull(sum(CreditAmt*ROE*qty),0) from SalesNotInRas where CustID=" & pCustID & " and PRG in " & PRGlist
        KQ = cmd.ExecuteScalar
        Return KQ
    End Function
    Private Shared Function DefineBG_TinChap(ByVal pCust As Integer, ByVal BG_CR As String, ByVal ppConn As SqlClient.SqlConnection) As Decimal
        Dim KQ As Decimal, strSQL As String
        Dim cmd As SqlClient.SqlCommand = ppConn.CreateCommand
        strSQL = "select isnull(sum(BGAmount),0) as TTL from bg where status='OK' and custid=" & pCust
        strSQL = strSQL & " and '" & Now.Date & "' between BGValidFrm and BGExpireDate "
        If BG_CR = "BG" Then
            strSQL = strSQL & " and Bank not in ('CRD','CSH')"
        Else
            strSQL = strSQL & " and Bank in ('CRD','CSH')"
        End If
        cmd.CommandText = strSQL
        KQ = cmd.ExecuteScalar
        Return KQ
    End Function
    Private Shared Function DefineUnUse(ByVal pCust As Integer, ByVal ParPmtType As String, ByVal pCutOverDateNew As Date, ByVal pCutOverDate As Date, pppConn As SqlClient.SqlConnection) As Decimal
        Dim KQ As Decimal = 0, DuKyTruoc As Decimal
        Dim cmd1 As SqlClient.SqlCommand = pppConn.CreateCommand

        If pCutOverDateNew = "00:00" Then pCutOverDateNew = pCutOverDate

        cmd1.CommandText = "select isnull(sum(conlai*ROE),0) as Conlai from TT_KhachTra where status <>'XX' and " & _
            "PmtType ='" & ParPmtType & "' and CustID=" & pCust & " and isClosed=0 and fstUpdate >'" & pCutOverDateNew & _
            "' group by OrgCurr"
        KQ = cmd1.ExecuteScalar

        If ParPmtType = "PPD" Then
            cmd1.CommandText = "select VND_Avail from tt_ChotCongNo where status='OK' and custid=" & pCust & _
                " and asof='" & pCutOverDateNew & "' "
            DuKyTruoc = cmd1.ExecuteScalar
            KQ = KQ + DuKyTruoc
        End If
        Return KQ
    End Function
    Private Shared Function DefineUnPaid(ByVal pCust As Integer, ByVal pCutOverDateNew As Date, ByVal pCutOverDate As Date, pppConn As SqlClient.SqlConnection) As Decimal
        Dim KQ As Decimal = 0
        Dim cmd1 As SqlClient.SqlCommand = pppconn.CreateCommand
        If pCutOverDateNew = "00:00" Then pCutOverDateNew = pCutOverDate
        cmd1.CommandText = "select isnull(sum(conNo * BSR),0) as ConNo from tt_GhiNoKhach where status <>'XX' and " & _
            " DebType ='PSP' and CustID=" & pCust & "  and invdate >'" & _
            pCutOverDateNew & "' and conNo <>0 "
        KQ = cmd1.ExecuteScalar
        Return KQ
    End Function
    Private Shared Function DefineUnInvoice(ByVal pCust As Integer, ByVal ParPmtType As String, ByVal pCutOverDateNew As Date, ByVal pCutOverDate As Date, pppConn As SqlClient.SqlConnection) As Decimal
        Dim KQ As Decimal = 0
        Dim tblUnInv As New DataTable
        If pCutOverDateNew = "00:00" Then pCutOverDateNew = pCutOverDate
        Dim adapter As New SqlClient.SqlDataAdapter("select SRV, isnull(sum(Amount*ROE),0) as Amt from func_CC_PSP " & _
                    "(" & pCust & ",'" & pCutOverDateNew & "','" & ParPmtType & "')  group by SRV", pppConn)
        adapter.Fill(tblUnInv)
        For i As Int16 = 0 To tblUnInv.Rows.Count - 1
            If tblUnInv.Rows(i)("SRV") = "R" Then
                KQ = KQ - tblUnInv.Rows(i)("Amt")
            Else
                KQ = KQ + tblUnInv.Rows(i)("Amt")
            End If
        Next
        Return KQ
    End Function
    Public Shared Function RefreshBalance(ByVal pCustType As String, ByVal pCutDateNew As Date, ByVal pCustID As Integer, ByVal ShowResult As Boolean, ByVal pRMK As String, ByVal pConn As SqlClient.SqlConnection, ByVal pSICode As String, Optional pCnString As String = "") As Decimal
        Dim VNDAvail As Decimal, cutDate As Date
        If InStr("PSP_PPD", pCustType) = 0 Then Return 0
        cutDate = GetCutOverDate(pCustType, pConn)
        VNDAvail = defineVND_Avail(pCustID, pCustType, cutDate, pCutDateNew, pRMK, pConn, pSICode, pCnString)
        If ShowResult Then MsgBox("Current Balance is " & Format(VNDAvail, "#,##0.00"), MsgBoxStyle.Information, "TransViet Airlines :: RAS :. System Message")
        Return VNDAvail
    End Function
    Public Shared Function CheckOverDue(ByVal pCustID As Integer, ByVal pConn As SqlClient.SqlConnection) As String
        Dim cmd As SqlClient.SqlCommand = pConn.CreateCommand
        Dim vConNo As Decimal = 0, NewOverDue As Int16, CurrOverDue As Int16
        Dim KyBCID As Integer, BLCMonitor As String
        cmd.CommandText = "select RecID from KyBaoCao where custid=" & pCustID & " and status='OK'"
        KyBCID = cmd.ExecuteScalar
        If KyBCID = 0 Then Return ""
        cmd.CommandText = "select BLCMonitor from KyBaoCao where RecID=" & KyBCID
        BLCMonitor = cmd.ExecuteScalar
        If BLCMonitor = "Y" Then
            cmd.CommandText = "select OverDue from KyBaoCao where RecID=" & KyBCID
            CurrOverDue = cmd.ExecuteScalar
            cmd.CommandText = "select isnull(sum(conno),0) from GhiNoKhach where custID=" & pCustID & _
                " and DueDate<getdate()"
            vConNo = cmd.ExecuteScalar
            NewOverDue = IIf(vConNo > 0, -1, 0)
            If NewOverDue <> CurrOverDue Then
                cmd.CommandText = "update KyBaoCao set OverDue=" & NewOverDue & ", LstUpdate=getdate(), LstUser='" & _
                    Environment.MachineName.ToString() & "_" & vConNo.ToString.Trim & "' where RecID=" & KyBCID
                cmd.ExecuteNonQuery()
                cmd.CommandText = "insert OverDueLog (CustId, StatusB4, StatusAfter, PendingAmt) values (" & _
                    pCustID & "," & CurrOverDue & "," & NewOverDue & "," & vConNo & ")"
                cmd.ExecuteNonQuery()
            End If
        End If
        Return IIf(vConNo > 0, "Err. OverDue Amount " & vConNo.ToString, "")
    End Function
End Class
