Module mdlTables
    Private strSQL As String
    Private MyCust As New objCustomer
    Private Cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Function Update_INV(ByVal pSRV As String, ByVal pFullName As String, ByVal pAddr As String, ByVal pTaxCode As String, ByVal pAmt As Decimal _
                               , ByVal pFOP As String, ByVal pSoBanIn As Integer, ByVal pInvID As Integer, pCustID As Integer) As String
        MyCust.CustID = pCustID
        strSQL = "update INV set "
        strSQL = strSQL & "SRV='" & pSRV & "', "
        strSQL = strSQL & "CustID=" & MyCust.CustID & ", "
        strSQL = strSQL & "CustShortName='" & MyCust.ShortName.Replace("--", "") & "', "
        strSQL = strSQL & "CustFullName=N'" & Replace(pFullName, "--", "") & "', "
        strSQL = strSQL & "CustAddress=N'" & Replace(pAddr, "--", "") & "', "
        strSQL = strSQL & "CustTaxCode='" & Replace(pTaxCode, "--", "") & "', "
        strSQL = strSQL & "Amount=" & Math.Round(pAmt, 0) & ", "
        strSQL = strSQL & "PrintCopy=PrintCopy+" & pSoBanIn & ", "
        strSQL = strSQL & "FstUser='" & myStaff.SICode & "', "
        strSQL = strSQL & "FOP='" & pFOP & "' "
        strSQL = strSQL & " Where RecID=" & pInvID
        Return strSQL
    End Function
    Public Function Insert_INV(ByVal SQL_Exec As String, ByVal pINVNo As String, ByVal pAL As String, ByVal pRCPID As Integer, Optional pNgayTaoDautien As Date = Nothing) As String
        Dim KQ As Integer
        If pNgayTaoDautien = Nothing Then pNgayTaoDautien = Now
        strSQL = "insert into INV (InvNo,AL, RCPID, city, FstUser, fstUpdate) values ('" & pINVNo & "','" & pAL & _
            "'," & pRCPID & ",'" & myStaff.City & "','" & myStaff.SICode & "','" & Format(pNgayTaoDautien, "dd-MMM-yy HH:mm") & "')"
        If SQL_Exec = "S" Then
            Return strSQL
        Else
            Cmd.CommandText = strSQL & "; SELECT SCOPE_IDENTITY() AS [RecID]"
            KQ = Cmd.ExecuteScalar
            Return KQ.ToString
        End If
    End Function
    Public Function Insert_FOP(ByVal pRCPID As Integer, ByVal pRCPNO As String, ByVal pFOP As String _
                               , ByVal pCurr As String, ByVal pROE As Decimal, ByVal pAmt As Decimal _
                               , ByVal pDoc As String, ByVal pRMK As String, pCustID As Integer _
                               , intCcId As Integer) As String
        If String.IsNullOrEmpty(pRMK) Then pRMK = ""
        If pDoc Is Nothing Then pDoc = ""

        'bo sung 16APR15, neu quay CWT SGN nhap CRD thi status=QQ de KT theo doi ca the
        Dim vStatus As String = "OK"
        If pFOP = "CRD" And myStaff.City = "SGN" And myStaff.Counter = "CWT" Then vStatus = "QQ"
        'End  bo sung
        strSQL = "insert FOP (RCPID, RCPNO, FOP, Currency, ROE, Amount, Document, RMK, CustomerID, FstUser" _
            & ", Status, CcId) values (" & pRCPID & ",'" & pRCPNO & "','" & _
            pFOP & "','" & pCurr & "'," & pROE & "," & pAmt & ",'" & pDoc.Replace("--", "") & "','" _
            & pRMK.Replace("--", "") & "'," & _
            pCustID & ",'" & myStaff.SICode & "','" & vStatus & "'," & intCcId & ")"
        Return strSQL
    End Function
    Public Function Insert_RCP(ByVal pRCPNO As String, ByVal pAL As String) As Integer
        Try
            Cmd.CommandText = "insert RCP (RCPNO, AL, SBU, FstUser) values ('" & pRCPNO & "','" & pAL & "','" & _
                MySession.Domain & "','" & myStaff.SICode & "') ; SELECT SCOPE_IDENTITY() AS [RecID]"
            Return Cmd.ExecuteScalar
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function Insert_KhachTra(ByVal SQL_Exec As String, ByVal orgDate As Date, ByVal Curr As String, ByVal FOP As String, ByVal OrgDocs As String, ByVal orgAmt As Decimal, ByVal note As String, ByVal Payer As String, ByVal ROE As Decimal, ByVal Description As String, ByVal pPmtType As String, pCustID As Integer) As String
        Dim KQ As Integer
        MyCust.CustID = pCustID
        strSQL = "insert tt_Khachtra (custID, custshortname, Orgdate, OrgCurr, FOP,PmtType, "
        strSQL = strSQL & "OrgAmt, receiveby, Note, custname, ROE, description, FstUser) values ("
        strSQL = strSQL & MyCust.CustID & ",'"
        strSQL = strSQL & MyCust.ShortName & "','"
        strSQL = strSQL & orgDate & "','"
        strSQL = strSQL & Curr & "','"
        strSQL = strSQL & FOP & "','"
        strSQL = strSQL & pPmtType & "',"
        strSQL = strSQL & orgAmt & ",'"
        strSQL = strSQL & "BO-BackOFC" & "','"
        strSQL = strSQL & note & " ','"
        strSQL = strSQL & Replace(Payer, "--", "") & "',"
        strSQL = strSQL & ROE & " ,'"
        strSQL = strSQL & Replace(Description, "--", "") & "','"
        strSQL = strSQL & myStaff.SICode & "')"
        If SQL_Exec = "S" Then
            Return strSQL
        Else
            Cmd.CommandText = strSQL & "; SELECT SCOPE_IDENTITY() AS [RecID]"
            KQ = Cmd.ExecuteScalar
            Return KQ.ToString
        End If
    End Function
    Public Function Insert_GhiNoKhach(ByVal SQL_Exec As String, ByVal InvCurr As String, ByVal InvDate As Date, ByVal InvAmt As Decimal, ByVal pStatus As String, ByVal note As String, ByVal DueDate As Date, ByVal BSR As Decimal, pCustID As Integer) As String
        Dim KQ As Integer
        MyCust.CustID = pCustID
        strSQL = "insert tt_GhiNoKhach (custID, custShortname, DebType, InvCurr, Invdate, InvAmt, Status, NOte, DueDate, "
        strSQL = strSQL & " BSR, FstUser) values ('"
        strSQL = strSQL & MyCust.CustID & "','"
        strSQL = strSQL & MyCust.ShortName & "','"
        strSQL = strSQL & "PSP" & "','"
        strSQL = strSQL & InvCurr & "','"
        strSQL = strSQL & InvDate & "','"
        strSQL = strSQL & InvAmt & "','"
        strSQL = strSQL & pStatus & "','"
        strSQL = strSQL & Replace(note, "--", "") & "','"
        strSQL = strSQL & Format(DueDate, "dd-MMM-yy") & " 23:59',"
        strSQL = strSQL & BSR & ",'"
        strSQL = strSQL & myStaff.SICode & "')"
        If SQL_Exec = "S" Then
            Return strSQL
        Else
            Cmd.CommandText = strSQL & "; SELECT SCOPE_IDENTITY() AS [RecID]"
            KQ = Cmd.ExecuteScalar
            Return KQ.ToString
        End If
    End Function
    Public Function Insert_ApplyPayment(ByVal SQL_Exec As String, ByVal GhiNoID As Integer, ByVal KhachTraID As Integer, ByVal AmtInDebCurr As Decimal, ByVal Curr As String, ByVal ROE As Decimal, ByVal CrdDoc As String, ByVal Note As String) As String
        Dim KQ As Integer
        strSQL = "insert into tt_ApplyPayment (GhiNoID, KhachTraID,  AmtInDebCurr, Currency, ROE, CrdDocs, Note, FstUser) Values ("
        strSQL = strSQL & GhiNoID & ","
        strSQL = strSQL & KhachTraID & ","
        strSQL = strSQL & AmtInDebCurr & ",'"
        strSQL = strSQL & Curr & "',"
        strSQL = strSQL & ROE & ",'"
        strSQL = strSQL & CrdDoc & "','"
        strSQL = strSQL & Replace(Note, "--", "") & "','"
        strSQL = strSQL & myStaff.SICode & "')"
        If SQL_Exec = "S" Then
            Return strSQL
        Else
            Cmd.CommandText = strSQL & "; SELECT SCOPE_IDENTITY() AS [RecID]"
            KQ = Cmd.ExecuteScalar
            Return KQ.ToString
        End If
    End Function
End Module
