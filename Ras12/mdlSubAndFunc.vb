Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.MySharedFunctionsWzConn
Imports System.IO
Imports System.Text.RegularExpressions
Module mdlSubAndFunc
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Function GetLastFileName_FullPath(pDDan As String, pPattern As String) As String
        Dim DDan As String = pDDan
        Dim FileNamePattern As String = DDan & pPattern
        Dim LstModif As Date = CDate("01-Jan-2010")
        Dim tmpFName As String = Dir(FileNamePattern)
        Dim LstFileName As String = ""
        Do While tmpFName <> ""
            If LstModif < File.GetCreationTime(DDan & tmpFName) Then
                LstModif = File.GetCreationTime(DDan & tmpFName)
                LstFileName = DDan & tmpFName
            End If
            tmpFName = Dir()
        Loop
        Return LstFileName
    End Function
    Public Sub Book_a_MSGR(pSI As String, pPSW As String, pApp As String, pTC As String, pID As Integer)
        Dim Fname As String = GetLastFileName_FullPath("X:\RAS2K7\", "MSGR*.exe")
        Fname = Fname & " " & pSI & "|" & pPSW & "|" & pApp & "|" & pTC & "|" & pID
        On Error Resume Next
        Shell(Fname)
        On Error GoTo 0
    End Sub
    Public Function getNextXX_Status(pRCPID As Integer) As String
        Dim KQ As String = ScalarToString("INV", "top 1 Status", "RCPID=" & pRCPID & " order by Status desc")
        Dim ChckChar As String = KQ.Substring(1, 1)
        If KQ = "OK" Then
            Return "X0"
        ElseIf InStr("012345678", ChckChar) > 0 Then
            Return "X" & (CInt(ChckChar) + 1).ToString.Trim
        ElseIf ChckChar = "9" Then
            Return "XA"
        End If
        Return "X" & Chr(Asc(ChckChar) + 1)
    End Function
    Public Sub AutoUploadPPD2TransVietVN()
        Dim DDanPPD As String = ScalarToString("MISC", "VAL", "cat='TRXListingPath'")
        Dim FName As String = Dir(DDanPPD & "TRXlisting*PPD_RPT.xls")
        On Error Resume Next
        Do While FName <> ""
            UploadFileToFtp(DDanPPD & FName, "ftp://transviet.vn/Upload/", "transviet.vn", "Abcd@123456789", "APP")
            Kill(DDanPPD & FName)
            FName = Dir()
        Loop
        On Error GoTo 0
    End Sub

    Public Function HasNewerVersion_R12(pProductVersion As String) As Boolean
        Dim VerConn As New SqlClient.SqlConnection
        If My.Computer.Name = "5-086" Then Return False
        Dim AppName As String = Application.ExecutablePath
        If AppName.ToLower.Contains("_0.exe") Then
            MsgBox("Please Use RAS12-RunMe.exe file to Start Application", MsgBoxStyle.Critical, msgTitle)
            Return True
        End If

        VerConn.ConnectionString = "server=" & MySession.ServerIP & ";uid=reporter;pwd=reporter;database=RAS12"
        VerConn.Open()
        Dim cmd As SqlClient.SqlCommand = VerConn.CreateCommand
        cmd.CommandText = "Select VAL from MISC where cat='APPVERSION'"
        Dim vProductVersion As String = cmd.ExecuteScalar
        VerConn.Close()
        If vProductVersion <> pProductVersion Then
            MsgBox("Newer Version Available. Please Quit This Application and Run it Again.", MsgBoxStyle.Critical, msgTitle)
            Return True
        End If
        Return False
    End Function
    Public Function isExistExeNewerThanMe(pCurrentExeName As String) As Boolean
        If My.Computer.Name = "5-086" Then Return False
        If Not pCurrentExeName.Contains("_") Then Return True ' buoc ten file co dang AppYY_*.exe
        Dim CurrentExeDate As Date = File.GetLastWriteTime(DDAN & "\" & pCurrentExeName)
        Dim tmpFileName As String = Dir(DDAN & pCurrentExeName.Substring(0, 7) & "*.exe")
        Do While tmpFileName <> ""
            If tmpFileName.Contains("_") Then
                If File.GetLastWriteTime(DDAN & "\" & tmpFileName) > CurrentExeDate Then Return True
            End If
            tmpFileName = Dir()
        Loop
        Return False
    End Function

    Public Function GetFareTaxChargeInVND(pSRV As String, pINVID As Integer) As Decimal
        Dim KQ As Decimal
        Dim dTbl As DataTable = GetDataTable("Select F_VND, T_VND, C_VND from TKTNO_INVNO where status='OK' and INVID=" & pINVID)
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            If pSRV = "S" Then
                KQ = KQ + dTbl.Rows(i)("F_VND") + dTbl.Rows(i)("T_VND") + dTbl.Rows(i)("C_VND")
            ElseIf pSRV = "R" Then
                KQ = KQ + Math.Abs(dTbl.Rows(i)("F_VND")) + Math.Abs(dTbl.Rows(i)("T_VND")) - dTbl.Rows(i)("C_VND")
            End If
        Next
        Return KQ
    End Function
    Public Function FillRBD_byRTG(RTG As String, pRBD As String) As String
        Dim KQ As String = pRBD
        RTG = RTG.Replace(" ", "")
        Dim NoOfSeg As Int16, tmpRBD As String = ""
        NoOfSeg = (RTG.Length - 3) / 5
        If KQ = "" Then KQ = StrDup(NoOfSeg, "Y")
        If NoOfSeg > KQ.Length Then
            tmpRBD = StrDup(NoOfSeg - KQ.Length, KQ.Substring(0))
            KQ = KQ & tmpRBD
        ElseIf NoOfSeg < KQ.Length Then
            KQ = KQ.Substring(0, NoOfSeg)
        End If
        If NoOfSeg > 2 Then
            For i As Int16 = 1 To NoOfSeg - 2
                If RTG.Substring(5 * i + 3, 2) = "//" Then
                    KQ = KQ.Substring(0, i) & "-" & KQ.Substring(i + 1)
                End If
            Next
        End If
        Return KQ
    End Function
    Public Function FillFB_byRTG(RTG As String, pFB As String) As String
        Dim tmpFB As String = pFB, DefaultFB As String
        If tmpFB = "" Then tmpFB = "Y"
        tmpFB = tmpFB.Replace("++", "+//+")
        RTG = RTG.Replace(" ", "")
        Dim NoOfSeg As Int16, NoOfPlus As Int16, FBi(16) As String
        NoOfSeg = (RTG.Length - 3) / 5 - 1
        NoOfPlus = UBound(tmpFB.Split("+"))
        If NoOfSeg > NoOfPlus Then
            DefaultFB = tmpFB.Split("+")(0)
            For i As Int16 = 1 To NoOfSeg - NoOfPlus
                tmpFB = tmpFB & "+" & DefaultFB
            Next
        ElseIf NoOfSeg < NoOfPlus Then
            For i As Int16 = 0 To NoOfSeg - 1
                tmpFB = tmpFB & "+" & tmpFB.Split("+")(i)
                tmpFB = tmpFB.Substring(1)
            Next
        End If
        FBi = Split(tmpFB, "+")
        For i As Int16 = 1 To NoOfSeg - 1
            If RTG.Substring(5 * i + 3, 2) = "//" Then
                FBi(i) = "//"
            End If
        Next
        tmpFB = FBi(0)
        For i As Int16 = 1 To NoOfSeg
            tmpFB = tmpFB + "+" + FBi(i)
        Next
        Return tmpFB
    End Function
    Public Sub TaoBanGhiTKTNO_INVNO_Standard(pRCPID As Integer, pINVNO As String, pINVID As Integer, pWzOrWoCharge As String)
        Dim ROE As Decimal = ScalarToDec("RCP", "ROE", "RecID=" & pRCPID)
        cmd.CommandText = "insert TKTNO_INVNO (INVNO, INVID, FstUser, RCPID, TKNO, F_VND, T_VND, C_VND, " & _
            " CTV_VND) select '" & pINVNO & "'," & pINVID & ", FstUser, RCPID, TKNO, " & _
            " (Fare-AgtDisctVAL)*qty*" & ROE & ", Tax*Qty*" & ROE & ", Charge*" & ROE
        If pWzOrWoCharge = "WO" Then
            cmd.CommandText = cmd.CommandText & ", 0"
        Else
            cmd.CommandText = cmd.CommandText & ", ChargeTV*" & ROE
        End If
        cmd.CommandText = cmd.CommandText & " from TKT where status='OK' and RCPID=" & pRCPID
        cmd.ExecuteNonQuery()
    End Sub
    Public Function Invalid3Rd(pDocNo As String, pCustID As Integer) As Boolean
        If pCustID <> 8085 Or pDocNo <> "TS24" Then Return True
        Return False
    End Function
    Public Function InvalidTourCode(ByVal pTC As String, ByVal pCustID As Integer, pSRV As String, pTKNO As String, pIsNew As Boolean, pDOI As Date, Optional pExDoc As String = "") As Boolean
        Dim RecNo As Integer, RCPID As Integer, S_TCode As String, vWhat As String = IIf(pSRV = "R", pTKNO, pExDoc)
        Dim strDK As String
        If MySession.Counter <> "CWT" Then Return True
        strDK = " custid=" & pCustID & " and TCode='" & pTC & "' and BillingBy in ('Event','Bundle') and status not in ('XX','RR')"
        If pSRV = "R" Or pExDoc <> "" Then
            RCPID = ScalarToInt("TKT", "RCPID", "TKNO='" & vWhat & "' and srv='S' and status<>'XX'")
            S_TCode = ScalarToString("FOP", "top 1 Document", "RCPID=" & RCPID & " and status='OK'")
            If pTC <> S_TCode Then Return True
            RecNo = ScalarToInt("DuToan_Tour", "RecID", strDK)
        ElseIf pSRV = "S" And pExDoc = "" Then
            If pIsNew Then strDK = strDK & " and edate >='" & pDOI & "'"
            RecNo = ScalarToInt("DuToan_Tour", "RecID", strDK)
        End If
        If RecNo = 0 Then Return True
        Return False
    End Function

    Public Function GenPseudoTKT(ByVal pDoc As String, ByVal pAL As String) As String
        Dim KQ As String, i As Integer
        KQ = ScalarToString("TKT", "top 1 TKNO", "left(tkno,6)='" & pDoc & " " & pAL & "' order by TKNO desc")
        If KQ = "" Then
            KQ = pDoc & " " & pAL & "00 000001"
        Else
            i = CInt(Strings.Right(KQ, 6)) + 1
            KQ = pDoc & " " & pAL & "00 " & Format(i, "000000")
        End If
        Return KQ
    End Function
    Public Sub LoadCombo(ByRef cboInput As ComboBox, ByVal strQuerry As String _
                         , objConn As SqlClient.SqlConnection)
        Dim daConditions As SqlClient.SqlDataAdapter
        Dim dsConditions As New DataSet

        daConditions = New SqlClient.SqlDataAdapter(strQuerry, objConn)
        If daConditions.Fill(dsConditions, "RESULT") > 0 Then
            cboInput.DataSource = dsConditions.Tables("RESULT")
            cboInput.DisplayMember = "Value"
            cboInput.ValueMember = "Value"
            'LoadCombo = cboInput
            dsConditions.Dispose()
            daConditions.Dispose()
        End If
    End Sub
    Public Sub LoadComboDisplay(ByRef cboInput As ComboBox, ByVal strQuerry As String _
                         , objConn As SqlClient.SqlConnection)
        Dim daConditions As SqlClient.SqlDataAdapter
        Dim dsConditions As New DataSet

        daConditions = New SqlClient.SqlDataAdapter(strQuerry, objConn)
        If daConditions.Fill(dsConditions, "RESULT") > 0 Then
            cboInput.DataSource = dsConditions.Tables("RESULT")
            cboInput.DisplayMember = "Display"
            cboInput.ValueMember = "Value"
            'LoadCombo = cboInput
            dsConditions.Dispose()
            daConditions.Dispose()
        End If
    End Sub
    Public Function LoadDataGridView(ByRef dgInput As DataGridView, strQuerry As String _
                                     , objConn As SqlClient.SqlConnection) As Boolean
        Dim daConditions As SqlClient.SqlDataAdapter
        Dim dsConditions As New DataSet

        daConditions = New SqlClient.SqlDataAdapter(strQuerry, objConn)
        daConditions.Fill(dsConditions, "Result")
        dgInput.DataSource = dsConditions.Tables("Result")
        dgInput.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgInput.AutoResizeColumns()
        dsConditions.Dispose()
        daConditions.Dispose()
    End Function
    Public Function GetDataTable(ByVal pStrCmd As String, Optional ByVal pConn As SqlClient.SqlConnection = Nothing) As DataTable
        Dim tblResults As New DataTable
        'Try
        If pConn Is Nothing Then
            Dim adapter As New SqlClient.SqlDataAdapter(pStrCmd, Conn)
            adapter.Fill(tblResults)
        Else
            Dim adapter As New SqlClient.SqlDataAdapter(pStrCmd, pConn)
            adapter.Fill(tblResults)
        End If
        'Catch ex As Exception
        '    MsgBox("SQL ERROR:" & pStrCmd & vbNewLine & ex.Message)
        'End Try
        
        Return tblResults
    End Function
    Public Sub CheckRightForALLForm(ByVal frm As Form)
        Dim Ctrl As Control
        myStaff.CurrObj = frm.Name
        If myStaff.URights = "" Then Exit Sub
        For i As Int16 = 0 To myStaff.URights.Split("|").Length - 1
            Ctrl = findControl(myStaff.URights.Split("|")(i), frm)
            Ctrl.Enabled = False
        Next
    End Sub
    Private Function findControl(ByVal pName As String, ByVal pCtrl As Control) As Control
        Dim ReturnControl As Control = Nothing
        For Each Ctrl_i As Control In pCtrl.Controls
            If Ctrl_i.Name.ToUpper = pName.ToUpper Then
                ReturnControl = Ctrl_i
                Exit For
            ElseIf Ctrl_i.Controls.Count > 0 Then
                ReturnControl = findControl(pName, Ctrl_i)
                If ReturnControl IsNot Nothing Then Exit For
            End If
        Next
        Return ReturnControl
    End Function

    Public Function GetDomainNameFromTRXCode(ByVal pTRXCode As String) As String
        If pTRXCode = "XX" Then
            Return "EDU"
        ElseIf pTRXCode = "TS" Then
            Return "TVS"
        Else
            Return "GSA"
        End If
    End Function
    Public Function CityAPTToCountry_Area_City(ByVal pReturnWhat As String, ByVal pInput As String) As String
        Dim KQ As String = ScalarToString("cityCode", pReturnWhat, " airport='" & pInput & "' OR CITY='" & pInput & "'")
        If KQ Is Nothing Then
            KQ = IIf(pReturnWhat = "Country", "??", "???")
        End If
        Return KQ
    End Function
    Public Sub LoadCmb_VAL(ByVal cmb As ComboBox, ByVal pQry As String)
        If InStr(pQry.ToUpper, "ORDER BY") = 0 Then
            pQry = pQry & " order by DIS"
        End If
        cmb.DataSource = GetDataTable(pQry)
        cmb.ValueMember = "VAL"
        cmb.DisplayMember = "DIS"
    End Sub
    Public Sub LoadCmbAL(ByVal cmb As ComboBox)
        If MySession.Domain = "GSA" Then
            LoadCmb_MSC(cmb, myStaff.GSA)
        ElseIf MySession.Domain = "TVS" Then
            LoadCmb_MSC(cmb, myStaff.ALList)
        Else
            LoadCmb_MSC(cmb, "select 'XX' as VAL")
        End If
        MyAL.Domain = MySession.Domain
        MyAL.ALCode = ""
    End Sub
    Public Sub LoadCmb_MSC(ByVal cmb As ComboBox, ByVal pCat As String)
        Dim dTable As DataTable
        If pCat.Length < 16 Then
            dTable = GetDataTable("select VAL from MISC where status='OK' and CAT='" & pCat & "' order by val ")
        Else
            If InStr(pCat.ToUpper, "ORDER BY") = 0 Then pCat = pCat & " order by VAL"
            dTable = GetDataTable(pCat)
        End If
        cmb.Items.Clear()
        For i As Int32 = 0 To dTable.Rows.Count - 1
            cmb.Items.Add(dTable.Rows(i)("VAL"))
        Next
        If cmb.Items.Count > 0 Then cmb.Text = cmb.Items(0).ToString
    End Sub
    Public Function CheckTKTformat(ByVal vTKNO As String) As String
        Dim KQ As String = ""
        If vTKNO.Length <> 13 And vTKNO.Length <> 15 Then
            KQ = "Error. Invalid Ticket Number!  "
            Return KQ
        End If
        If (InStr(vTKNO, " ") > 0 And vTKNO.Length = 13) Or _
            (InStr(vTKNO, " ") = 0 And vTKNO.Length = 15) Then
            KQ = "Error. Invalid Ticket Number!  "
            Return KQ
        End If
        If InStr(vTKNO, " ") = 0 And vTKNO.Length = 13 Then
            KQ = AddSpace2TKNO(vTKNO)
        ElseIf InStr(vTKNO, " ") > 0 And vTKNO.Length = 15 Then
            KQ = vTKNO
        End If
        If Not MyAL.ValidDocCode.Contains(KQ.Substring(0, 3)) Then
            KQ = "Error. Invalid Airline Document Code!  "
        End If
        Return KQ
    End Function

    Public Function GenInvNo_QD153(ByVal pRCP As String, ByVal pKyHieu As String) As String
        Dim KQ As String = "", strDK As String
        Dim strPrefix As String = pRCP.Substring(0, 2) + pKyHieu + pRCP.Substring(4, 2) + MySession.POSCode 'AL yy POS
        strDK = " left(invno,7)='" & strPrefix & "' "

        If Now() > CDate("01-aug-16") AndAlso myStaff.City = "HAN" AndAlso InStr("UA_RJ_TS", MySession.TRXCode) > 0 Then
            strDK = strDK & " AND FSTUPDATE>'01-AUG-16'  " ' THEM DO HAN danh lai so hd tu so 1 cho 3 hang nay. co the bo sau 1Jan17
        End If

        strDK = strDK & "  order by invno desc"
        KQ = ScalarToString("INV", "top 1 INVNO", strDK)
        If KQ <> "" Then
            KQ = strPrefix & Format(CInt(Strings.Right(KQ, 5)) + 1, "00000")
        Else
            KQ = strPrefix & "00001"
        End If
        Return KQ
    End Function
    Public Function CheckEditROE(ByVal pAL As String) As Boolean
        If myStaff.SupOf <> "" Then Return True
        If pubVarSRV = "O" Then
            Return True
        Else
            If MySession.TRXCode <> "TS" Then
                Return False
            Else
                Return Not MyAL.isTVA
            End If
        End If
        Return False
    End Function
    Public Function DefineDefauSF(ByVal pAL As String) As Decimal
        Return ScalarToDec("MISC", "top 1 VAL ", "cat ='MINSF' or CAT='MINSF" & pAL & "' order by CAT DESC")
    End Function
    Public Function ForEX_12(ByVal DOS As Date, ByVal pCurr As String, ByVal pType As String, ByVal pAL As String _
                             , Optional ByVal parQuay As String = "**") As clsROE
        Dim objROE As New clsROE
        Dim surCharge As Decimal
        Dim dTable As DataTable
        dTable = GetDataTable("select * from ForEx where Currency='" & pCurr & "' and Status='OK' order by EffectDate DESC, recid desc ")
        For i As Int16 = 0 To dTable.Rows.Count - 1
            If dTable.Rows(i)("EffectDate") <= DOS And _
                (dTable.Rows(i)("ApplyROEto") = "YY" Or InStr(dTable.Rows(i)("ApplyROEto"), pAL) > 0 _
                 Or pAL = "YY" Or InStr(dTable.Rows(i)("ApplySCto"), pAL) > 0) Then
                objROE.Amount = dTable.Rows(i)(pType)
                objROE.Id = dTable.Rows(i)("RecId")

                If pType = "RECID" Then Exit For
                surCharge = dTable.Rows(i)("SurCharge")
                If pType = "BBR" Then surCharge = -surCharge
                If (parQuay <> "**" And InStr(dTable.Rows(i)("ApplySCto"), parQuay) > 0) Or _
                    InStr(dTable.Rows(i)("ApplySCto"), pAL) > 0 Then
                    objROE.Amount = objROE.Amount + surCharge
                End If
                Exit For
            End If
        Next
        Return objROE
    End Function
    Public Function getHotStr(ByVal vAL As String, ByVal vHotKey As String, ByVal vtxtBox As String) As String
        Return ScalarToString("MISC", "details", "cat+VAL='HOTKEY" & vAL & "' and val1='" & vtxtBox & "' and VAL2='" & vHotKey & "'")
    End Function

    Public Function XDHoaHongGSA(ByVal varAL As String, ByVal varProduct As String, ByVal pCurr As String) As Decimal
        Dim KQ As Decimal = 0
        Dim dTable As DataTable
        dTable = GetDataTable("select * from Charge_Comm where cat='COMM' and status='OK' and  AL='" & varAL & _
        "' and currency='" & pCurr & "' and Type='" & varProduct & "' and ('" & Now.Date & "' between ValidFrom and ValidThru)")
        For i As Int16 = 0 To dTable.Rows.Count - 1
            KQ = dTable.Rows(i)("Amount")
            If dTable.Rows(i)("AmtType") = "PCT" Then
                KQ = KQ / 100
            End If
        Next
        XDHoaHongGSA = KQ
    End Function
    Private Sub TaoDataChoALRPT(ByVal ppAL As String, ByVal ppFrm As Date, ByVal ppThru As Date)
        Dim strDKDate As String, QryForFOPAmt As String, QryForFOPdocs As String
        Dim tblFOP As String = "##ztmpal_fop"
        Dim tblRCP As String = "##ztmpal_RCP"
        Dim tblTKT As String = "##ztmpal_TKT"
        Dim tblTrans As String = "##ztmpal_tkt_trans"

        On Error Resume Next
        cmd.CommandText = "drop table " & tblFOP
        cmd.ExecuteNonQuery()
        cmd.CommandText = "drop table " & tblRCP
        cmd.ExecuteNonQuery()
        cmd.CommandText = "drop table " & tblTKT
        cmd.ExecuteNonQuery()
        cmd.CommandText = "drop table " & tblTrans
        cmd.ExecuteNonQuery()
        On Error GoTo 0

        strDKDate = " SRV <>'A' and  DOI between '" & ppFrm & "' and '" & ppThru & " 23:59'"
        QryForFOPAmt = "(select sum(amount) from FOP  where status='OK' and "
        QryForFOPAmt = QryForFOPAmt & " t1.rcpid=FOP.rcpid and "
        QryForFOPdocs = "(select top 1 Document from FOP where status='OK' and "
        QryForFOPdocs = QryForFOPdocs & " t1.rcpid=FOP.rcpid and "

        cmd.CommandTimeout = 64
        cmd.CommandText = "select RCPNo, TKNO, SRV, FTKT, fare, ShownFare, itinerary, Currency, charge, CommVal, CommPCT, " & _
            " NetToAL, tax, tax_d , RCPID, stockCtrl, qty, doi, doctype, faretype into " & _
            tblTKT & " from TKT where statusAL='OK' and " & strDKDate & _
            " and (srv <>'A' or (srv='A' and Fare+tax+charge >0 )) and al='" & ppAL & _
            "' and TKNO not like '%TV%' and TKNO not like '%GRP%' and rcpid in (select recid from rcp where sbu='GSA')"

        cmd.ExecuteNonQuery()

        cmd.CommandText = "select  FOP, Currency, sum(Amount) as Amount into " & tblFOP & " from FOP " & _
                " where fop in ('MCO','PTA','EXC','UCF') and status ='OK' and RCPID in  (select RCPID from " & tblTKT & " t " & _
                " where  RECID in (Select top 1 recid from " & tblTKT & "  b where b.tkno= t.tkno and b.srv=t.SRV" & _
                " and document not like 'GRP%' order by fstupdate desc)) group by FOP, Currency "

        cmd.ExecuteNonQuery()

        cmd.CommandText = "Select RCPID, TKNO, FTKT, SRV, Qty, DOI, fare, tax, charge, CommVal, nettoAL, itinerary, Currency, " & _
            " Tax_D, (select sum(qty) from " & tblTKT & " t2 where  t2.rcpid=t1.rcpid) as TTLPax," & QryForFOPAmt & "FOP = 'MCO') as MCO, " & _
            QryForFOPAmt & "FOP = 'PTA') as PTA, " & QryForFOPAmt & "FOP = 'EXC' and document not like 'GRP%') as EXC, " & QryForFOPAmt & "FOP = 'UCF') as UCF, " & _
            QryForFOPdocs & " FOP = 'MCO') as MCO_doc, " & QryForFOPdocs & " FOP = 'PTA') as PTA_doc, " & QryForFOPdocs & _
            " FOP = 'EXC' and document not like 'GRP%') as EXC_doc, " & QryForFOPdocs & " FOP = 'UCF') as UCF_doc, DocType, RCPNO, Faretype, ShownFare,StockCtrl " & _
            " into " & tblTrans & " from " & tblTKT & " t1"
        cmd.ExecuteNonQuery()

        cmd.CommandText = "select RecID, RCPNo, SRV, TTLDue, Status, CustID, Currency, ROE, VNDPCT into " & tblRCP & " from RCP " & _
            " where al='" & ppAL & "' and status+city='OK" & MySession.City & "' and recID in (select RCPID from " & tblTKT & ")"

        On Error Resume Next
        cmd.ExecuteNonQuery()
        On Error GoTo 0
    End Sub
    Private Sub TaoDataChoDailyRPT(ByVal ppAL As String, ByVal ppFrm As Date, ByVal ppThru As Date, ByVal pBoPhan As String)
        Dim strDKDate As String, StrTKTFieldList As String, StrTKTdk As String, strSQL As String
        If MySession.Domain = "TVS" Then ppAL = "TS"
        Dim KTFOP = "##ztmpkt_fop_" & ppAL.ToLower & "_" & MySession.Counter, KTRCP = "##ztmpkt_rcp_" & ppAL.ToLower & "_" & MySession.Counter
        Dim KTTKT = "##ztmpkt_tkt_" & ppAL.ToLower & "_" & MySession.Counter
        KTFOP = KTFOP.Replace("-", "")
        KTRCP = KTRCP.Replace("-", "")
        KTTKT = KTTKT.Replace("-", "")
        On Error Resume Next
        cmd.CommandText = "drop table " & KTFOP
        cmd.ExecuteNonQuery()
        cmd.CommandText = "drop table " & KTRCP
        cmd.ExecuteNonQuery()
        cmd.CommandText = "drop table " & KTTKT
        cmd.ExecuteNonQuery()
        On Error GoTo 0
        strDKDate = " fstUpdate between '" & Format(ppFrm, "dd-MMM-yy") & "' and '" & Format(ppThru, "dd-MMM-yy") & " 23:59'"
        StrTKTFieldList = " Select RCPNo, SRV, TKNO, FTKT, Qty, RCPID, Fare, tax, Charge, ChargeTV, CommVAL, "
        StrTKTFieldList = StrTKTFieldList & " NetToAL, AgtDisctPCT, AgtDisctVAL, Itinerary, Promocode, DocType, DOF "
        StrTKTdk = " RCPID in (Select RecID from " & KTRCP & ")"

        strSQL = "Select RecID, RCPNo, SRV, CustShortName, TTLDue, Discount, Charge, Status, CustID, Currency, " & _
            " ROE, RPTNo, City, Counter, cast(Stock as varchar(4)) as AL into " & KTRCP & _
            " from RCP where " & strDKDate & " and status <>'NA' and al='" & ppAL & "'" & _
            " and City='" & MySession.City & "'"
        Select Case myStaff.Counter
            Case "CWT", "HAN", "ALL"
            Case Else
                strSQL = strSQL & " and location ='" & myStaff.Location & "'"
        End Select

        If pBoPhan = "C" Then
            strSQL = strSQL & " and Counter='" & MySession.Counter & "'"
        End If

        strSQL = strSQL & "  order by RCPNO"
        cmd.CommandText = strSQL
        Append2TextFile(strSQL)
        cmd.ExecuteNonQuery()

        cmd.CommandText = "Select  RCPID, RCPNo, FOP, Currency, Amount, ROE, 'NEWS' as PmtType, Document  into " & KTFOP & _
                " from FOP where FOP <>'RND' and left(rmk,8)<>'BO_CLEAR' and status in ('OK','QQ') and RCPID in " & _
                "(select RecID from " & KTRCP & " where status <> 'XX')"
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        cmd.CommandText = " insert into " & KTFOP & " Select  RCPID, RCPNo, FOP, Currency, Amount, ROE, 'EDIT', Document " & _
                " from FOP where FOP <>'RND' and status in('OK','QQ') and " & strDKDate & " and RCPID NOT in " & _
                " (select RecID from " & KTRCP & " where status <> 'XX') " & _
                " and left(rcpno,2)='" & ppAL & "' and left(rmk,8)<>'BO_CLEAR' " & _
                " and RCPID NOT in (select recid from rcp where status='NA') "
        If MySession.Domain = "TVS" And InStr("CN", pBoPhan) > 0 Then
            cmd.CommandText = cmd.CommandText & " and rcpid in (select recid from rcp where counter='" & MySession.Counter & "')"
        End If
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        'Lay cac ve SA-R bt co statusAL = OK
        strSQL = StrTKTFieldList & " into " & KTTKT & " from TKT"
        strSQL = strSQL & " Where SRV in ('S','A','R') and StatusAL='OK' and " & StrTKTdk
        If MySession.Domain = "GSA" Then
            strSQL = strSQL & " and al='" & ppAL & "'"
        End If

        cmd.CommandText = strSQL
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        'lay cac ve R cua void deu, tuc la SRV=R va statusNoibo=OK  va phai thay so RCP tuong ung co TKT.SRV=V statusAL =OK
        strSQL = "Insert into " & KTTKT & StrTKTFieldList & " from TKT Where SRV='R' and Status<>'XX'"
        If MySession.Domain = "GSA" Then
            strSQL = strSQL & " and AL='" & ppAL & "' "
        End If
        strSQL = strSQL & " and " & StrTKTdk & " and RCPID in (Select RCPID from TKT"
        strSQL = strSQL & " Where SRV='V' and statusAL='OK' and " & StrTKTdk
        If MySession.Domain = "GSA" Then
            strSQL = strSQL & " and al='" & ppAL & "'"
        End If
        strSQL = strSQL & " )"
        cmd.CommandText = strSQL
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        ' Lay cac ve ban ma sau do bi void deu, khi void deu statusal=XX nen no ko vao khi lay lan 1
        'DK la SRV=S + StatusNoiBo=OK va phai thay 1 so ve tuong ung bi SRV=V va statusAL=OK

        strSQL = "Insert into " & KTTKT & StrTKTFieldList & " from TKT Where SRV='S' and Status<>'XX'"
        If MySession.Domain = "GSA" Then
            strSQL = strSQL & " and al='" & ppAL & "'"
        End If
        strSQL = strSQL & " and " & StrTKTdk & " and TKNO in (Select TKNO from TKT Where SRV='V' and statusAL='OK'"
        If MySession.Domain = "GSA" Then
            strSQL = strSQL & " and al='" & ppAL & "'"
        End If
        strSQL = strSQL & " )"

        cmd.CommandText = strSQL
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        ' voi KT nen chi Lay cac ve Void Xin, tuc la ko nam trong loai co RCP tuong ung chua ve R 
        strSQL = "Insert into " & KTTKT & StrTKTFieldList & " from TKT Where SRV='V' and "
        strSQL = strSQL & " StatusAL='OK' and " & StrTKTdk
        If MySession.Domain = "GSA" Then
            strSQL = strSQL & " and al='" & ppAL & "'"
        End If
        strSQL = strSQL & " and RCPID not in (Select RCPID from TKT"
        strSQL = strSQL & " Where SRV='R' and status<>'XX' and " & StrTKTdk + ")"
        cmd.CommandText = strSQL
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        'thay cac ve void Deu trong ngay thanh SRV=C de khau tru tien trong bao cao ngay  
        strSQL = "update " & KTTKT & " set SRV='C' where SRV='R' and "
        strSQL = strSQL & " TKNO in (select tkno from TKT where srv='V' and statusal='OK' "
        If MySession.Domain = "GSA" Then
            strSQL = strSQL & " and al='" & ppAL & "'"
        End If
        strSQL = strSQL & " ) and "
        strSQL = strSQL & " TKNO in (select tkno from " & KTTKT & " where srv='S') "
        cmd.CommandText = strSQL
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        'thay cac RCP void Deu trong ngay thanh SRV=C de khau tru tien trong bao cao ngay  
        cmd.CommandText = "update " & KTRCP & " set SRV='C' where SRV='R' and RECID in (select RCPID from " & KTTKT & " where srv='C')"
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        cmd.CommandText = "update " & KTRCP & " set AL='1A' where RecID in (select RCPID from " & KTTKT & _
            " where substring(tkno,5,4) in (select VAL from misc where cat='BSPSTOCK'))"
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

        cmd.CommandText = "update " & KTRCP & " set AL=AL+'1A' where RecID in (select RCPID from " & KTTKT & _
            " where left(tkno,3) in (select SecondCode from Airline where secondCode <>''))"
        Append2TextFile(cmd.CommandText)
        cmd.ExecuteNonQuery()

    End Sub
    Public Sub InHoaDon(ByVal strPath As String, ByVal parFileName As String, ByVal parViewPrint As String, ByVal parRCPNO As String, ByVal parFrm As Date, ByVal parTo As Date, ByVal ParNewValue As Decimal, ByVal pAL As String, ByVal pDomain As String, Optional ByVal ParLoaiHD As String = "", Optional AnyInt1 As Integer = 0)
        Dim AppXls As Microsoft.Office.Interop.Excel.Application, WkBook As Microsoft.Office.Interop.Excel.Workbook, WkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim varAL As String
        If parFileName = "" Then
            MsgBox("Please Select Invoice Type", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If parFileName.Substring(0, 3) <> "R12" Then parFileName = "R12_" & parFileName
        varAL = Dir(strPath & "\" & parFileName)
        If varAL = "" Then
            MsgBox("Template File Not Found. Plz Check " & parFileName _
                   , MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        On Error Resume Next
        AppXls = CreateObject("Excel.Application")
        On Error GoTo 0
        If parRCPNO.Length > 2 Then
            varAL = parRCPNO.Substring(0, 2)
        Else
            varAL = pAL
        End If
        If InStr(parFileName.ToUpper, "SR_") > 0 Or InStr(parFileName.ToUpper, "AR_") > 0 Or _
            InStr(parFileName.ToUpper, "DISCREPANCY") > 0 Then
            If InStr(parFileName.ToUpper, "DISCREPANCY") > 0 Then
                cmd.CommandText = "Exec DiscrepancyRPT '" & varAL & "','" & Format(parFrm, "dd MMM yyyy") & "','" & _
                    Format(parTo, "dd MMM yyyy") & "','" & MySession.Domain & "'"
                cmd.ExecuteNonQuery()
            ElseIf InStr(parFileName.ToUpper, "AR_") > 0 Then
                TaoDataChoDailyRPT(varAL, parFrm, parTo, ParLoaiHD)
            ElseIf InStr(parFileName.ToUpper, "SR_") > 0 Then
                TaoDataChoALRPT(varAL, parFrm, parTo)
            End If
        End If
        On Error GoTo CloseXLS
        AppXls.Visible = True
        WkBook = AppXls.Workbooks.Open(strPath & "\" & parFileName, , , , "aibiet", , , , , True)
        WkSheet = WkBook.Worksheets("Para")
        WkSheet.Cells.Range("B1").Value = parRCPNO
        WkSheet.Cells.Range("B2").Value = "'" & varAL
        WkSheet.Cells.Range("B3").Value = myStaff.SICode
        WkSheet.Cells.Range("B4").Value = parFrm
        WkSheet.Cells.Range("B5").Value = parTo
        WkSheet.Cells.Range("B6").Value = ParNewValue
        WkSheet.Cells.Range("B7").Value = pDomain
        WkSheet.Cells.Range("B8").Value = ParLoaiHD
        WkSheet.Cells.Range("B9").Value = parViewPrint
        WkSheet.Cells.Range("B10").Value = AnyInt1

        If parFileName = "R12_CC_CT_TKT Listing CWT.xlt" Then
            WkSheet.ComboBox1.Text = ScalarToString("CustomerList", "CustShortName", "Status<>'XX' and RecId=" & ParNewValue)
        End If

        WkSheet.Cells.Range("B15").Value = "YES"

        If InStr("PFVO", parViewPrint) = 0 Then GoTo CloseXLS
        WkSheet = WkBook.Worksheets("RPT")
        If InStr("V", parViewPrint) > 0 Then
            AppXls.Visible = True
            If InStr(parFileName.ToUpper, "RECEIPT") Or InStr(parFileName.ToUpper, "VAT") Then
                WkSheet.Cells.Range("D6").Value = "P R E V I E W"
            ElseIf InStr(parFileName.ToUpper, "SR_") + InStr(parFileName.ToUpper, "AR_") > 0 Then
                WkSheet.Cells.Range("G2").Value = "P R E V I E W.   N O T  F O R  R E P O R T I N G  P U R P O S E"
            End If
            WkSheet.PrintPreview(vbNo)
        ElseIf parViewPrint = "P" Then
            AppXls.Visible = True
            WkSheet.PrintPreview(vbNo)
        ElseIf parViewPrint = "O" Then
            AppXls.Visible = False
            WkSheet.PrintOut()
        End If
CloseXLS:
        WkBook.Close(SaveChanges:=False)
        AppXls.Quit()
        AppXls = Nothing
        On Error GoTo 0
    End Sub
    Public Function CalcCharge(ByVal pCharge As String, ByVal varWho As String, ByVal pTRXCurr As String, ByVal pROE As Decimal, ByVal pDOS As Date, ByVal pAL As String) As Decimal
        Dim KQ As Decimal, c As Decimal, Curri As String, tmpROE As Decimal
        For i As Int16 = 0 To UBound(pCharge.Split("|"))
            If pCharge.Split("|")(i).Substring(0, 2) = varWho Then
                Curri = pCharge.Split("|")(i).Split(":")(1).Substring(0, 3)
                c = CDec(pCharge.Split("|")(i).Split(":")(1).Substring(3))
                If pTRXCurr <> "VND" And Curri = "VND" Then
                    c = c / pROE
                ElseIf pTRXCurr = "VND" And Curri <> "VND" Then
                    tmpROE = ForEX_12(pDOS, Curri, "BSR", pAL).Amount
                    c = c * tmpROE
                End If
                KQ = KQ + c
            End If
        Next
        Return KQ
    End Function
    Public Function DefineNextTKNO(ByVal parThisTKT As String, ByVal IsTKTless As Boolean) As String
        Dim lstTKT As Double, KQ As String
        If Not parThisTKT.Contains("Z") Or Left(parThisTKT, 3) = "GRP" Then
            lstTKT = parThisTKT.Substring(9, 6)
            lstTKT = CDbl(lstTKT) + 1
            KQ = parThisTKT.Substring(0, 9) & Format(lstTKT, "000000")
        Else
            lstTKT = parThisTKT.Substring(parThisTKT.Length - 3, 2)
            lstTKT = CDbl(lstTKT) + 1
            KQ = parThisTKT.Substring(0, parThisTKT.Length - 2) & Format(lstTKT, "00")
        End If
        DefineNextTKNO = KQ
    End Function
    Public Function UpdateTblINVHistory(ByVal ParInvID As Integer, ByVal parInvNo As String, pInt As Int16) As String
        Return "update INV set printCopy=printCopy +" & pInt & " where RecID=" & ParInvID & " ; " & _
            UpdateLogFile("INV", "INPPrint", parInvNo, ParInvID, "", "", "", "", "", "")
    End Function
    Public Function UpdateLogFile(ByVal pTbl As String, ByVal pAction As String, ByVal pF1 As String, ByVal pF2 As String, ByVal pF3 As String, ByVal pF4 As String, ByVal pF5 As String, Optional ByVal pF6 As String = "", Optional ByVal pF7 As String = "", Optional ByVal pF8 As String = "", Optional ByVal pF9 As String = "", Optional ByVal pF10 As String = "", Optional ByVal pF11 As String = "", Optional ByVal pF12 As String = "") As String
        Dim KQ As String = "insert ActionLog (TableName, doWhat, F1, F2, F3, F4, F5, F6, F7, F8, f9,f10, F11,F12,city, ActionBy) Values ('"
        KQ = KQ & pTbl & "','" & pAction & "',N'" & pF1 & "',N'" & pF2 & "','" & pF3 & "','" & pF4 & "','" & pF5 & "','" & pF6 & "','" & _
                pF7 & "','" & pF8 & "','" & pF9 & "','" & pF10 & "','" & pF11 & "','" & pF12 & "','" & MySession.City & "','" & myStaff.SICode & "')"
        Return KQ
    End Function
    Public Function ConvertDomainAccess2SqlList(strDomainAccess As String) As String
        Return "('" & Replace(strDomainAccess, "_", "','") & "')"
    End Function
    Public Function CreateFromDate(dteInput As Date) As String
        Return Format(dteInput, "dd MMM yy 00:00")
    End Function
    Public Function CreateToDate(dteInput As Date) As String
        Return Format(dteInput, "dd MMM yy 23:59")
    End Function
    Public Function ExecuteNonQuerry(strQuerry As String, objConn As SqlClient.SqlConnection) As Boolean
        Dim objCmd As SqlClient.SqlCommand = objConn.CreateCommand
        objCmd.CommandText = strQuerry
        Try
            objCmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Append2TextFile(vbNewLine & "ERROR|" & myStaff.UID & "|" & Now & vbNewLine & strQuerry & vbNewLine & ex.Message)
        End Try

    End Function
    Public Function UpdateListOfQuerries(lstQuerries As List(Of String), objConn As SqlClient.SqlConnection _
                                         , Optional blnGetLastInsertedRecId As Boolean = False _
                                         , Optional ByRef intLastInsertedRecId As Integer = 0) As Boolean
        Dim i As Integer
        Dim strQuerry As String = String.Empty
        Dim trcSql As SqlClient.SqlTransaction = objConn.BeginTransaction()
        Dim objCmd As SqlClient.SqlCommand = objConn.CreateCommand
        objCmd.Transaction = trcSql

        Try
            For i = 0 To lstQuerries.Count - 1
                strQuerry = lstQuerries(i)
                objCmd.CommandText = strQuerry
                If Not String.IsNullOrEmpty(strQuerry) Then
                    objCmd.ExecuteNonQuery()
                    If blnGetLastInsertedRecId AndAlso UCase(strQuerry).StartsWith("INSERT") Then
                        objCmd.CommandText = "select SCOPE_IDENTITY()"
                        intLastInsertedRecId = objCmd.ExecuteScalar
                    End If
                End If
            Next
            trcSql.Commit()
            Return True
        Catch ex As Exception
            trcSql.Rollback()
            Append2TextFile(vbNewLine & "ERROR|" & myStaff.SICode & "|" & Now & vbNewLine & strQuerry & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Public Function Append2TextFile(ByVal strText As String) As Boolean
        Dim strLogFile As String = My.Application.Info.DirectoryPath & "\" _
                                            & Format(Today, "yyMMdd") & pstrPrg & ".txt"

        Dim objLogFile As New System.IO.StreamWriter(strLogFile, True)
        objLogFile.WriteLine(strText)
        objLogFile.Close()
        objLogFile = Nothing
        Return True
    End Function
    Public Function ClearLogFile(ByVal intNbrOfDays As Int16) As Boolean
        ' make a reference to a directory
        Dim objDir As New IO.DirectoryInfo(Directory.GetCurrentDirectory)
        Dim objFileInfo As IO.FileInfo
        Dim intMinDay As Integer
        intMinDay = Format(DateAdd(DateInterval.Day, -intNbrOfDays, Now), "yyMMdd")

        For Each objFileInfo In objDir.GetFiles
            If IsNumeric(Mid(objFileInfo.Name, 1, 6)) _
                AndAlso Mid(objFileInfo.Name, 1, 6) < intMinDay _
                AndAlso Mid(objFileInfo.Name, 7) = pstrPrg & ".txt" Then
                objFileInfo.Delete()
            End If
        Next

        Return True
    End Function
    Public Function ChangeGridViewSelectedColumn(ByRef dgInput As DataGridView, blnSelected As Boolean) As Boolean
        For Each objdgrCust As DataGridViewRow In dgInput.Rows
            With objdgrCust
                .Cells("Selected").Value = blnSelected
            End With
        Next
    End Function
    Public Function CheckFormatTextBox(ByRef txtInput As System.Windows.Forms.TextBox _
                                        , Optional ByVal blnNumeric As Boolean = False _
                                        , Optional ByVal intMinLength As Int16 = 0 _
                                         , Optional ByVal intMaxLength As Int16 = 0) As Boolean
        Dim strName As String
        If txtInput.Tag = "" Then
            strName = Mid(txtInput.Name, 4)
        Else
            strName = txtInput.Tag
        End If
        If txtInput.Text = "" Then
            MsgBox("Invalid value for " & strName)
            txtInput.Focus()
            Return False
        End If
        If intMaxLength > 0 AndAlso txtInput.Text.Length > intMaxLength Then
            MsgBox("Invalid value for " & strName)
            txtInput.Focus()
            Return False
        End If
        If intMinLength > 0 AndAlso txtInput.Text.Length < intMinLength Then
            MsgBox("Invalid value for " & strName)
            txtInput.Focus()
            Return False
        End If
        If blnNumeric AndAlso Not IsNumeric(txtInput.Text) Then
            MsgBox("Invalid value for " & strName)
            txtInput.Focus()
            Return False
        End If
        Return True
    End Function
    Public Function CheckFormatComboBox(ByRef cboInput As System.Windows.Forms.ComboBox _
                                        , Optional ByVal blnNumeric As Boolean = False _
                                        , Optional ByVal intMinLength As Int16 = 0 _
                                        , Optional ByVal intMaxLength As Int16 = 0) As Boolean
        Dim strName As String
        If cboInput.Tag = "" Then
            strName = Mid(cboInput.Name, 4)
        Else
            strName = cboInput.Tag
        End If

        If (intMaxLength > 0 AndAlso cboInput.Text.Length > intMaxLength) _
            Or (intMinLength > 0 AndAlso cboInput.Text.Length < intMinLength) _
            Or (blnNumeric AndAlso Not IsNumeric(cboInput.Text)) Then
            GoTo Quit
        End If
        Return True
Quit:
        MsgBox("Invalid value for " & strName)
        cboInput.Focus()

    End Function
    Public Function GetCardTypeByCardNbr(strCardNbr As String) As String
        Select Case Mid(strCardNbr, 1, 2)
            Case "40", "41", "42", "43", "44", "45", "46", "47", "48", "49"
                Return "VI"
            Case "34", "37", "35"
                Return "AX"
            Case "36"
                Return "DC"
            Case "51", "52", "53", "54", "55"
                Return "CA"
            Case Else
                Return "XX"
                'Throw New SystemException("Unable to find CreditCard Type for ")
        End Select
    End Function
    
    Public Function GetCDRs(ByVal objSqlConx As SqlClient.SqlConnection, ByVal strCmc As String _
                            , Optional ByVal blnBeforeRas As Boolean = False _
                            , Optional blnIncludeConditionalCDR As Boolean = False) As Collection

        Dim colCDRs As New Collection
        Dim strQry As String
        Dim drResult As SqlClient.SqlDataReader
        Dim cmdSql As New SqlClient.SqlCommand
        cmdSql.Connection = objSqlConx

        strQry = "select * from cwt.dbo.GO_CDRs"
        strQry = strQry & " where CMC='" & strCmc & "' and status='OK'"
        If blnBeforeRas Then
            strQry = strQry & " and CollectionMethod<>'RAS'"
        End If
        If Not blnIncludeConditionalCDR Then
            strQry = strQry & " and Mandatory='M'"
        End If
        strQry = strQry & " order by cdrnbr"
        cmdSql.CommandText = strQry
        drResult = cmdSql.ExecuteReader
        If Not drResult Is Nothing Then
            Do While drResult.Read
                Dim objCdr As New clsCwtCdr
                objCdr.Nbr = drResult("CdrNbr")
                objCdr.CdrName = drResult("CdrName")
                objCdr.CharType = drResult("CharType")
                objCdr.MinLength = drResult("MinLength")
                objCdr.MaxLength = drResult("MaxLength")
                objCdr.Mandatory = drResult("Mandatory")
                colCDRs.Add(objCdr, objCdr.Nbr)
            Loop
        End If
        drResult.Close()
        Return colCDRs
    End Function
    Public Function IsAlphaOnly(ByRef strText As String) As Boolean
        Dim rgAlpha As New Regex("\d")
        If rgAlpha.IsMatch(strText) Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Function GenerateBankPaymentBatchNbr(ByVal objSqlConx As SqlClient.SqlConnection _
                                                , strBankName As String) As String
        Dim strBatchNbr As String
        Dim strQuerry As String = "Select top 1 SCBNo from UNC_Payments where substring(SCBNo,1,3)='" & strBankName _
                                  & "' and substring(SCBNo,4,6)=CONVERT(varchar,getdate(),12)" _
                                  & " order by substring(SCBNo,10,2) desc"
        Dim cmdSql As New SqlClient.SqlCommand
        cmdSql.Connection = objSqlConx
        cmdSql.CommandText = strQuerry
        strBatchNbr = cmdSql.ExecuteScalar

        If strBatchNbr = "" Then
            strBatchNbr = strBankName & Format(Now, "yyMMdd") & "01"
        Else
            strBatchNbr = Mid(strBatchNbr, 1, strBankName.Length) & (Mid(strBatchNbr, strBankName.Length + 1) + 1)
        End If
        Return strBatchNbr

    End Function
    Public Function ConvertToLetter(iCol As Integer) As String
        Dim iAlpha As Integer
        Dim iRemainder As Integer
        Dim strResult As String = String.Empty
        iAlpha = Int(iCol / 27)
        iRemainder = iCol - (iAlpha * 26)
        If iAlpha > 0 Then
            strResult = Chr(iAlpha + 64)
        End If
        If iRemainder > 0 Then
            strResult = strResult & Chr(iRemainder + 64)
        End If
        Return strResult
    End Function
    Public Function DefineDomRtg(strDomCities As String, strRtg As String)
        Dim strResult As String = "DOM"
        Dim arrCities As String() = strRtg.Split(" ")
        For Each strCity As String In arrCities
            If strCity.Length = 3 Then
                If Not strDomCities.Contains(strCity) Then
                    strResult = ""
                    Exit For
                End If
            End If
        Next
        Return strResult
    End Function
    Public Function GetTaxAmtFromTaxDetails(strTaxCode As String, strTaxDetails As String) As Decimal
        Dim decResult As Decimal = 0

        If strTaxDetails <> "" Then
            Dim arrTaxes() As String = strTaxDetails.Split("|")
            Dim i As Integer
            For i = 0 To arrTaxes.Length - 1
                If Mid(arrTaxes(i), 1, 2) = strTaxCode Then
                    decResult = decResult + Mid(arrTaxes(i), 3)
                End If
            Next
        End If
        Return decResult
    End Function
    
    Public Function CheckData(ByVal colRequiredData As Collection _
                                , ByVal colAvailableData As Collection _
                                , ByVal intCustId As Integer, dteDOI As Date) As Collection
        Dim colResult As New Collection
        Dim i As Integer

        For i = 1 To colRequiredData.Count
            Dim objRequiredData As clsRequiredData = colRequiredData(i)

            With objRequiredData
                .ClearOldCheck()
                Dim objAvaiData As New clsAvailableData

                If Not colAvailableData.Contains(.DataCode) Then
                    If objRequiredData.DefaultValue = "" Then
                        objRequiredData.ErrMsg = "MISSING"
                    Else
                        AddRequiredData(colAvailableData, .DataCode, .DefaultValue)
                    End If
                Else
                    objAvaiData = colAvailableData(.DataCode)
                    If IsSpecialValue(intCustId, .DataCode, objAvaiData.DataValue) Then
                        'pass object nay
                    ElseIf IsNormalValue(intCustId, .DataCode, objAvaiData.DataValue) Then
                        'pass object nay
                    ElseIf .CheckValues Then
                        objRequiredData.ErrMsg = "INVALID VALUE"

                    ElseIf objAvaiData.DataValue.Length < .MinLength Then
                        objRequiredData.ErrMsg = "INVALID MIN LENGTH"
                    ElseIf objAvaiData.DataValue.Length > .MaxLength Then
                        objRequiredData.ErrMsg = "INVALID MAX LENGTH"
                    ElseIf .CharType = "NUMERIC" AndAlso Not IsNumeric(objAvaiData.DataValue) Then
                        objRequiredData.ErrMsg = "VALUE IS NOT NUMERIC"
                    ElseIf .CharType = "ALPHA" AndAlso Not AlphaOnly(objAvaiData.DataValue) Then
                        objRequiredData.ErrMsg = "VALUE IS NOT ALPHA ONLY"
                    End If
                End If
                If .ErrMsg <> "" Then
                    .AvailableValue = objAvaiData.DataValue
                    colResult.Add(objRequiredData)
                End If

            End With

        Next
        Return colResult
    End Function
    Public Function AddRequiredData(ByVal colResult As Collection _
                    , ByRef strDataCode As String, ByRef strValue As String) As Boolean
        Try
            Dim objAvaiData As New clsAvailableData
            objAvaiData.DataCode = strDataCode
            objAvaiData.DataValue = strValue
            colResult.Add(objAvaiData, objAvaiData.DataCode)
        Catch ex As Exception
            If ex.Message.Contains("Duplicate") Then
                MsgBox("Duplicate RM*" & strDataCode)
                Return False
            End If
        End Try

        Return True
    End Function
    Public Function IsSpecialValue(ByVal intCustId As Integer, ByVal strDataCode As String _
                                    , ByVal strValue As String) As Boolean
        Dim intResult As Integer
        intResult = ScalarToInt("cwt.dbo.GO_RequiredDataValues", "top 1 RecId", " DataType='SPECIAL'" _
                    & " and Value='" & strValue & "' and CustId=" & intCustId)
        If intResult > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function IsNormalValue(ByVal intCustId As Integer, ByVal strDataCode As String _
                                    , ByVal strValue As String) As Boolean
        Dim intResult As Integer
        intResult = ScalarToInt("cwt.dbo.GO_RequiredDataValues", "top 1 RecId", " DataType='NORMAL'" _
                    & " and Value='" & strValue & "' and CustId=" & intCustId)
        If intResult > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function AlphaOnly(ByVal strText As String) As Boolean
        Dim rgCheck As New Regex("\d")

        If rgCheck.IsMatch(strText) Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Function GetDataRequirement(ByVal intCustId As Integer _
                                        , ByVal strMandatoryType As String _
                                        , Optional ByVal blnPosOnly As Boolean = False _
                                        , Optional strAppyTo As String = "") As Collection
        Dim colResult As New Collection
        Dim strQuerry As String
        Dim tblResult As DataTable

        strQuerry = "Select * from GO_RequiredData" _
                        & " where Status='OK' and CustId =" & intCustId
        If strMandatoryType <> "" Then
            strQuerry = strQuerry & " and Mandatory='" & strMandatoryType & "'"
        End If

        If blnPosOnly Then
            strQuerry = strQuerry & " and CollectionMethod='AGTINPUT'"
        End If

        If strAppyTo <> "" Then
            strQuerry = strQuerry & " and ApplyTo in ('ALL','" & strAppyTo & "')"
        End If

        For Each objRow As DataRow In tblResult.Rows
            Dim objRqData As New clsRequiredData
            With objRqData
                .DataCode = objRow("DataCode")
                .NameByCustomer = objRow("NameByCustomer")
                .MinLength = objRow("MinLength")
                .MaxLength = objRow("MaxLength")
                .Mandatory = objRow("Mandatory")
                .ConditionOfUse = objRow("ConditionOfUse")
                .CollectionMethod = objRow("CollectionMethod")
                .DefaultValue = objRow("DefaultValue")
                .CheckValues = objRow("CheckValues")
                .AllowSpecialValues = objRow("AllowSpecialValues")
                .CharType = objRow("CharType")
                colResult.Add(objRqData, objRqData.DataCode)
            End With
        Next

        Return colResult

    End Function
    Public Function CreateSms4UNC(strRecId As String, strBankName As String, blnUseBatchNbr As Boolean) As Boolean
        Dim strSql As String
        If Conn_Web.State = ConnectionState.Closed Then Conn_Web.Open()
        Dim arrMobiles() As String = {"0909250946", "0908900131"}
        Dim tblUnc As DataTable
        If blnUseBatchNbr Then
            tblUnc = GetDataTable("Select Curr,Sum(Amount) as Amount from UNC_Payments where SCBNo='" _
                                  & strRecId & "' group by Curr", Conn)
        Else
            tblUnc = GetDataTable("Select * from UNC_Payments where RecId=" & strRecId, Conn)
        End If

        Dim strMsg As String = strBankName & " " & tblUnc.Rows(0)("Curr") _
                                & " " & Format(tblUnc.Rows(0)("Amount"), "#,###.00")

        For Each strMobile As String In arrMobiles
            strSql = "Insert SMSLog (CustID, SMSText, Location, MobileNbr) values (-1,'" _
                & strMsg & "','SGN','" & strMobile & "')"
            ExecuteNonQuerry(strSql, Conn_Web)
        Next
        Conn_Web.Close()
    End Function
    

End Module
