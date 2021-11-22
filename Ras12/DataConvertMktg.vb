Imports SharedFunctions.MySharedFunctions
Module DataConvertMktg
    Private Cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Sub GanShortNameChoBSP()
        Dim dTable As DataTable, CustShortName As String, CustCity As String, testIATACodeExist As Integer
        dTable = GetDataTable("select ID, AGTN from mktg_MIDT.dbo.ua_hot where id in (select tkid from reportdata_BSP where custshortName='')")
        For i As Integer = 0 To dTable.Rows.Count - 1
            CustShortName = ScalarToString("MISC", "VAL1", "cat='BSPAGT' and VAL='" & dTable.Rows(i)("AGTN") & "'")
            If Not String.IsNullOrEmpty(CustShortName) Then
                CustCity = ScalarToString("MISC", "VAL2", "cat='BSPAGT' and VAL='" & dTable.Rows(i)("AGTN") & "'")
                Cmd.CommandText = "update ReportData_BSP set CustShortName ='" & CustShortName & "', City='" & CustCity & "' where TKID=" & dTable.Rows(i)("ID")
                Cmd.ExecuteNonQuery()
            Else
                testIATACodeExist = ScalarToInt("MISC", "RecID", "VAL='" & dTable.Rows(i)("AGTN") & "'")
                If testIATACodeExist = 0 Then
                    Cmd.CommandText = "insert misc (cat, val) values ('BSPAGT','" & dTable.Rows(i)("AGTN") & "')"
                    Cmd.ExecuteNonQuery()
                End If
            End If
        Next
    End Sub
    Public Sub GoDataConvert()
        If Conn.State <> ConnectionState.Open Then Conn.Open()
        Dim myRptData As New RptData, RTG As String, RecID As Integer, Amt As Decimal, ROE As Decimal
        Dim dTable As DataTable, Qry(2) As String, TblName(2) As String
        Dim ChayBSP As Int16 = 1
        Cmd.CommandText = "set datefirst 5"
        Cmd.ExecuteNonQuery()
        Dim strVnDomCities As String = GetColumnValuesAsString("CityCode", "Airport", "where Country='VN'", "_")

        TblName(0) = " reportData"
        TblName(1) = TblName(0) & "_BSP"
        'Lay du lieu TKT do sang ReportData
        Qry(0) = "select t.recid, Itinerary, Farebasis, DOI, ROEID, ROE, (fare+tax)*Qty + t.Charge as Amt, CustID, Fare*qty as Fare, Tax*qty as Tax, " & _
            "t.Charge, ChargeTV, CommVAL*qty as CommVAL, NetToAL*qty as NetToAL, r.Currency, left(t.RCPNO,2) as Counter, BkgClass, t.AL, AgtDisctVAL " _
            & ",DocType,Tkno,r.Counter as BizUnit,t.RMK " & DKDataConvertMktg_RAS

        'Lay du lieu hotData do sang ReportData_BSP
        Qry(1) = "select ID as RecID, Routing as Itinerary, Farebasis, DAIS, [Class] as BkgClass, TDNR, TRNC, NTFA, COBL, FareCur, 0 as CustID, NetRev as Amt, AL" & _
            DKDataConvertMktg_BSP
        If myStaff.City = "HAN" Then ChayBSP = 0

        For j As Int16 = 0 To ChayBSP
            dTable = GetDataTable(Qry(j))
            For i As Int16 = 0 To dTable.Rows.Count - 1
                Dim strDomRtg As String = String.Empty     'Xac dinh hanh trinh DOM cho CWT
                Amt = dTable.Rows(i)("Amt")
                RTG = dTable.Rows(i)("Itinerary").ToString.Trim


                If RTG.Trim = "" Then
                    If j = 1 Then
                        RTG = "XXX YY XXX"
                        'ElseIf dTable.Rows(i)("DocType") = "MCO" AndAlso dTable.Rows(i)("BizUnit") = "CWT" _
                        '    AndAlso dTable.Rows(i)("RMK").ToString.StartsWith("BIZT|") _
                        '    AndAlso myStaff.Counter = "CWT" Then
                        '    MsgBox("Missing Itinerary for ticket " & dTable.Rows(i)("TKNO") & ".Please report NMK!")
                        '    GoTo resumeHere
                    Else
                        GoTo resumeHere
                    End If
                End If
                If InStr(RTG, " ") = 0 Then
                    RTG = AddSpace2Rtg(RTG).Trim
                End If
                If j = 0 Then
                    myRptData.AL = dTable.Rows(i)("AL")
                    myRptData.DOI = dTable.Rows(i)("DOI")
                    myRptData.Counter = dTable.Rows(i)("Counter")
            
                    ROE = dTable.Rows(i)("ROE")
                    myRptData.VND = Amt * ROE
                    myRptData.F_VND = dTable.Rows(i)("Fare") * ROE
                    myRptData.T_VND = dTable.Rows(i)("Tax") * ROE
                    myRptData.C_VND = dTable.Rows(i)("Charge") * ROE
                    myRptData.TVCharge_VND = dTable.Rows(i)("ChargeTV") * ROE
                    myRptData.Comm_VND = dTable.Rows(i)("CommVAL") * ROE
                    myRptData.toAL_VND = dTable.Rows(i)("NetToAL") * ROE
                    myRptData.Disct_VND = dTable.Rows(i)("AgtDisctVAL") * ROE

                    ROE = ROEOfTixByVND(dTable.Rows(i)("Currency"), dTable.Rows(i)("ROEID"))
                    myRptData.USD = Amt / ROE
                    myRptData.F_USD = dTable.Rows(i)("Fare") / ROE
                    myRptData.T_USD = dTable.Rows(i)("Tax") / ROE
                    myRptData.C_USD = dTable.Rows(i)("Charge") / ROE
                    myRptData.TVCharge_USD = dTable.Rows(i)("ChargeTV") / ROE
                    myRptData.Comm_USD = dTable.Rows(i)("CommVAL") / ROE
                    myRptData.toAL_USD = dTable.Rows(i)("NetToAL") / ROE
                    myRptData.Disct_USD = dTable.Rows(i)("AgtDisctVAL") / ROE
                    strDomRtg = DefineDomRtg(pstrVnDomCities, RTG)
                Else
                    If dTable.Rows(i)("DAIS").ToString.Trim = "" Then GoTo resumeHere
                    'If (RTG.Length - 3) Mod 7 <> 0 Then GoTo resumeHere
                    If (RTG.Length - 3) Mod 7 <> 0 Then RTG = "XXX YY XXX"
                    myRptData.AL = dTable.Rows(i)("AL")
                    If myRptData.AL.Trim = "" Then myRptData.AL = ScalarToString("Airline", "AL", " docCode='" & dTable.Rows(i)("TDNR").ToString.Substring(0, 3) & "'")
                    myRptData.Counter = myRptData.AL
                    myRptData.DOI = DateSerial(dTable.Rows(i)("DAIS").ToString.Substring(0, 2), dTable.Rows(i)("DAIS").ToString.Substring(2, 2), dTable.Rows(i)("DAIS").ToString.Substring(4, 2))
                    'Amt = IIf(dTable.Rows(i)("NTFA") > 0, dTable.Rows(i)("NTFA"), dTable.Rows(i)("COBL"))
                    ROE = ForEX_12(myRptData.DOI, "USD", "BSR", myRptData.AL).Amount
                    If dTable.Rows(i)("FareCur") = "VND" Then
                        myRptData.VND = Amt
                        myRptData.USD = Amt / ROE
                    Else
                        myRptData.VND = Amt * ROE
                        myRptData.USD = Amt
                    End If
                End If
                RecID = dTable.Rows(i)("RecID")
                myRptData.OrgCity = RTG.Substring(0, 3)
                myRptData.Tuan = DatePart(DateInterval.WeekOfYear, myRptData.DOI)
                myRptData.Thang = Month(myRptData.DOI)
                myRptData.Nam = Year(myRptData.DOI)
                myRptData.Ngay = myRptData.DOI.Day

                If RTG = "XXX YY XXX" Then ' Empty route
                    myRptData.OrgCountry = ".."
                    myRptData.OrgArea = "..."
                    myRptData.DomInt = "..."
                    myRptData.OWRT = ".."
                    myRptData.Dest = "..."
                    myRptData.DestCity = "..."
                    myRptData.Country = ".."
                    myRptData.Area = "..."
                Else
                    myRptData.OrgCountry = CityAPTToCountry_Area_City("Country", myRptData.OrgCity)
                    myRptData.OrgArea = CityAPTToCountry_Area_City("Area", myRptData.OrgCity)
                    myRptData.DomInt = DefindDomInt(myRptData.OrgCountry, RTG)
                    myRptData.OWRT = DefineOWRT_New(RTG, myRptData.DomInt, myRptData.OrgCountry)
                    myRptData.Dest = DefineDest_new(RTG, myRptData.OWRT, myRptData.OrgCountry, myRptData.DomInt)
                    myRptData.DestCity = CityAPTToCountry_Area_City("City", myRptData.Dest)
                    myRptData.Country = CityAPTToCountry_Area_City("Country", myRptData.Dest)
                    myRptData.Area = CityAPTToCountry_Area_City("Area", myRptData.Dest)
                End If
                myRptData.Cabin = DefineCabin(dTable.Rows(i)("BkgClass"), myRptData.AL)
                myRptData.stFB = DefineFB(dTable.Rows(i)("FareBasis"), myRptData.Counter)
                Cmd.CommandText = "insert " & TblName(j) & " (TKID, Dest, OWRT, DomInt, DeKho, FB, Cabin, USD, VND, Country, Area, tuan, " & _
                    "Thang, Nam, Counter, OrgCity, OrgCountry, OrgArea, DestCity, Ngay"
                If j = 0 Then
                    Cmd.CommandText = Cmd.CommandText & ", F_USD, T_USD, C_USD, TVCharge_USD, Comm_USD, Net2AL_USD" &
                        ", F_VND, T_VND, C_VND, TVCharge_VND, Comm_VND, Net2AL_VND, Disct_USD, Disct_VND"
                End If
                Cmd.CommandText = Cmd.CommandText & ") values (" & RecID & ",'" & myRptData.Dest & "','" & myRptData.OWRT & "','" & _
                    myRptData.DomInt & "','" & strDomRtg & "','" & myRptData.stFB & "','" & myRptData.Cabin & "'," & _
                    myRptData.USD & "," & myRptData.VND & ",'" & myRptData.Country & "','" & myRptData.Area & "'," & myRptData.Tuan & _
                    "," & myRptData.Thang & "," & myRptData.Nam & ",'" & myRptData.Counter & "','" & myRptData.OrgCity & _
                    "','" & myRptData.OrgCountry & "','" & myRptData.OrgArea & "','" & myRptData.DestCity & "'," & myRptData.Ngay
                If j = 0 Then
                    Cmd.CommandText = Cmd.CommandText & "," & myRptData.F_USD & "," & myRptData.T_USD & "," & myRptData.C_USD & "," & _
                        myRptData.TVCharge_USD & "," & myRptData.Comm_USD & "," & myRptData.toAL_USD & "," & _
                        myRptData.F_VND & "," & myRptData.T_VND & "," & myRptData.C_VND & "," & myRptData.TVCharge_VND & "," & _
                        myRptData.Comm_VND & "," & myRptData.toAL_VND & "," & myRptData.Disct_USD & "," & myRptData.Disct_VND
                End If
                Cmd.CommandText = Cmd.CommandText & ")"
                Cmd.ExecuteNonQuery()
resumeHere:
            Next
        Next
        If myStaff.City = "SGN" Then
            GanShortNameChoBSP()
            Call BreakBySegment()
        End If
        CalcVat4Cwt()
    End Sub
    Private Function CalcVat4Cwt() As Boolean
        Dim tblReportData As System.Data.DataTable
        Dim i As Integer

        Dim strQuerry As String = "SELECT T.Tax_D, t.Charge_D,T.SRV, R.*, C.CustShortName " _
                & " FROM REPORTDATA R" _
                & " LEFT JOIN TKT T ON R.TKID=T.RecID" _
                & " LEFT JOIN RCP C ON C.RECID=T.RCPID" _
                & " WHERE T.Status<>'XX' AND T.DOI>'31 DEC 2016 23:59'" _
                & " AND R.SfNoVat_VND IS NULL AND R.Counter='TS' AND R.DEKHO<>''" _
                & " AND C.Status<>'XX'"

        tblReportData = GetDataTable(strQuerry, Conn)
        With tblReportData
            For i = 0 To .Rows.Count - 1
                Dim decVat4Fare As Decimal
                Dim decDomAptSvc As Decimal
                Dim decVat4DomAptSvc As Decimal
                Dim decVat4AlCharge As Decimal
                Dim decVat4TvCharge As Decimal
                Dim decVatTotal As Decimal
                Dim decSfNoVat As Decimal

                If .Rows(i)("DEKHO") = "DOM" Then
                    decVat4Fare = GetTaxAmtFromTaxDetails("UE", .Rows(i)("Tax_D"))
                    decDomAptSvc = (.Rows(i)("T_VND") - decVat4Fare)
                    decVat4DomAptSvc = decDomAptSvc - (decDomAptSvc / 1.1)
                    If .Rows(i)("SRV") = "R" Then
                        decVat4Fare = -decVat4Fare
                        decVat4DomAptSvc = -decVat4DomAptSvc
                    Else
                        decVat4AlCharge = .Rows(i)("C_VND") - (.Rows(i)("C_VND") / 1.1)
                    End If
                End If

                'Khong thu thue VAT cho GEVNHPH 
                If .Rows(i)("CustShortName") = "GEVNHPH" Then
                    decSfNoVat = .Rows(i)("TVCharge_VND")
                Else
                    decSfNoVat = .Rows(i)("TVCharge_VND") / 1.1
                End If
                decVat4TvCharge = .Rows(i)("TVCharge_VND") - decSfNoVat

                decVatTotal = decVat4Fare + decVat4DomAptSvc + decVat4AlCharge + decVat4TvCharge
                strQuerry = "Update ReportData set VAT_VND=" & decVatTotal _
                    & ",SfNoVat_VND=" & decSfNoVat _
                    & " where Tkid=" & .Rows(i)("TKID")
                ExecuteNonQuerry(strQuerry, Conn)
            Next
        End With

    End Function
    Private Structure RptData
        Public OWRT As String
        Public Dest As String
        Public Country As String
        Public Area As String
        Public DomInt As String
        Public stFB As String
        Public Cabin As String
        Public Counter As String
        Public AL As String
        Public USD As Decimal
        Public VND As Decimal
        Public Ngay As Int16
        Public Tuan As Int16
        Public Thang As Int16
        Public Nam As Int16
        Public DOI As Date
        Public DestCity As String
        Public OrgCountry As String
        Public OrgArea As String
        Public OrgCity As String
        Public Qty As Int16
        Public F_USD As Decimal
        Public T_USD As Decimal
        Public C_USD As Decimal
        Public TVCharge_USD As Decimal
        Public Comm_USD As Decimal
        Public toAL_USD As Decimal
        Public F_VND As Decimal
        Public T_VND As Decimal
        Public C_VND As Decimal
        Public TVCharge_VND As Decimal
        Public Comm_VND As Decimal
        Public toAL_VND As Decimal
        Public Disct_VND As Decimal
        Public Disct_USD As Decimal
    End Structure
    Private Function DefineCabin(pRBD As String, pAL As String) As String
        Dim KQ As String = "Y", FC As String, f As String, C As String
        FC = ScalarToString("RBD_Cabin", " F + '_' + C ", " al='" & pAL & "' and status='OK' ")
        pRBD = pRBD.Replace("+", "")
        If FC <> "" AndAlso FC <> "_" Then
            f = FC.Split("_")(0)
            C = FC.Split("_")(1)
            For i As Int16 = 0 To Len(pRBD) - 1
                If InStr(f, pRBD.Substring(i, 1)) > 0 Then
                    KQ = "F"
                    Exit For
                ElseIf InStr(C, pRBD.Substring(i, 1)) > 0 Then
                    KQ = "C"
                    Exit For
                End If
            Next
        End If
        Return KQ
    End Function
    Private Function DefindDomInt(pOrgCountry As String, ByVal pRTG As String) As String
        Dim SoChang As Int16
        SoChang = (pRTG.Length - 3) / 7
        For i As Int16 = 1 To SoChang
            If CityAPTToCountry_Area_City("Country", pRTG.Substring(7 * i, 3)) <> pOrgCountry Then
                Return "INT"
            End If
        Next
        Return "DOM"
    End Function
    Private Function DefineOWRT_New(ByVal pRTG As String, ByVal pDomINT As String, pOrgCountry As String) As String
        Dim OWRT As String = "??", ViTriGiua As Integer, LstCountry As String
        Dim SoChang As Integer, TPi As String, TPDoiXung As String
        If pRTG.Length < 10 Then Return "??"
        If pRTG.Length = 10 Then Return "OW"
        If pDomINT = "DOM" Then
            If pRTG.Substring(0, 3) <> Strings.Right(pRTG, 3) Then
                Return "OW"
            Else
                Return "RT"
            End If
        Else
            LstCountry = CityAPTToCountry_Area_City("Country", Strings.Right(pRTG, 3))
            If pOrgCountry <> LstCountry Then
                Return "OW"
            Else
                SoChang = (pRTG.Length - 3) / 7
                If SoChang / 2 > Int(SoChang / 2) Then
                    OWRT = "CT"
                Else
                    OWRT = "RT"
                    ViTriGiua = (SoChang / 2) + 1
                    For i As Int16 = 2 To ViTriGiua - 1
                        TPi = pRTG.Substring(7 * i - 6 - 1, 3)
                        TPDoiXung = pRTG.Substring(7 * (SoChang + 2 - i) - 6 - 1, 3)
                        If TPi <> TPDoiXung Then
                            OWRT = "CT"
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        Return OWRT
    End Function
    Private Function CleanRTG(ppRTG As String) As String
        Dim KQ As String, SoChang As Int16, ChangI As String, CT As String
        SoChang = (ppRTG.Length - 3) / 7
        If Not ppRTG.Contains("/") Or SoChang < 3 Then Return ppRTG
        KQ = ppRTG
        For i As Int16 = 2 To SoChang - 1
            ChangI = ppRTG.Substring(7 * (i - 1), 10)
            If ChangI.Contains("/") Then
                CT = CityAPTToCountry_Area_City("City", ChangI.Substring(0, 3))
                If CT = CityAPTToCountry_Area_City("City", Strings.Right(ChangI, 3)) Then
                    KQ = KQ.Replace(ChangI, CT)
                End If
            End If
        Next
        Return KQ
    End Function
    Private Function DefineDest_new(ByVal pRTG As String, ByVal pOWRT As String, pOrgCountry As String _
                                    , pDomInt As String) As String
        Dim Dest As String = "???", ViTriGiua As Integer, SoChang As Integer, isChanChang As Boolean
        Dim tblCityCode As New System.Data.DataTable
        Dim strFirstArea As String = String.Empty

        If pRTG.Length > 9 Then
            pRTG = CleanRTG(pRTG)
            If pOWRT = "OW" Then Return Strings.Right(pRTG, 3)
            SoChang = (pRTG.Length - 3) / 7
            isChanChang = IIf(SoChang Mod 2 = 0, True, False)
            ViTriGiua = (SoChang / 2) + 1
            If pOWRT = "RT" Then Return pRTG.Substring(7 * ViTriGiua - 7, 3)
            If pDomInt = "DOM" Then
                Return pRTG.Substring(7 * ViTriGiua - 7, 3)
            Else
                For i As Int16 = 1 To SoChang
                    tblCityCode = GetDataTable("Select * from CityCode where Airport='" & pRTG.Substring(7 * i, 3) & "'")
                    If tblCityCode.Rows.Count = 0 Then
                        MsgBox("Unable to find airport code " & pRTG.Substring(7 * i, 3))
                        Return ""
                    End If
                    If tblCityCode.Rows(0)("Country") <> pOrgCountry Then
                        If Dest = "???" Then
                            Dest = pRTG.Substring(7 * i, 3) ' tam thoi gan dest = tp qt 1. neu ko co tp qt khac thi ra khoi va return 
                            strFirstArea = tblCityCode.Rows(0)("Area")
                        ElseIf pOrgCountry = "VN" AndAlso tblCityCode.Rows(0)("Area") = "SEA" _
                            AndAlso strFirstArea = "EUR" Then       'Neu thanh pho thu 2 la SEA ma thanh pho 1 la EUR thi lay thanh pho 1
                            Return Dest
                        Else
                            Return pRTG.Substring(7 * i, 3) ' co tp qt thu 2 lay tp thu 2
                        End If
                    End If
                Next
            End If
        End If
        Return Dest
    End Function
    Private Function ROEOfTixByVND(ByVal pCurr As String, ByVal pROEID As Integer) As Decimal
        Dim KQ As Decimal
        If pCurr <> "VND" Then
            Return 1
        Else
            Cmd.CommandText = "select BSR from forex where recid=" & pROEID
            KQ = Cmd.ExecuteScalar
            Return IIf(KQ = 0, 21000, KQ)
        End If
    End Function
    Private Function DefineFB(ByVal pFB As String, ByVal pAL As String) As String
        If pFB = "" Then pFB = "Y"
        Cmd.CommandText = "insert into FB_RBD_QT (CAT, AL, Raw) values ('FB','" & pAL & "','" & pFB.Split("+")(0) & "')"
        Cmd.ExecuteNonQuery()
        Return pFB.Split("+")(0)
    End Function
    Private Sub BreakBySegment()
        Cmd.CommandText = "Delete from Segment where tkID in (select recid from tkt where status ='XX')"
        Cmd.ExecuteNonQuery()
        Dim strQry As String = "select RecID, Itinerary, BkgClass, FareBasis from tkt " & _
            " where al='VN' and qty <>0 and status <>'XX' and RecID not in (select TKID from Segment) and len(itinerary) >8 "
        Dim Rtg As String, Seg As Int16, RBD As String, FB As String
        Dim dTable As DataTable = GetDataTable(strQry)
        For i As Int16 = 0 To dTable.Rows.Count - 1
            Rtg = dTable.Rows(i)("Itinerary").trim
            If Strings.Right(Rtg, 3).Contains(" ") Then
                Rtg = Rtg.Replace(" ", "")
                Rtg = Rtg.Replace("NHACXR", "CXR")
                Rtg = Rtg.Replace("TMKVCL", "TMK")
                Rtg = AddSpace2Rtg(Rtg).Trim
            End If
            RBD = dTable.Rows(i)("BkgClass")
            FB = dTable.Rows(i)("FareBasis")
            Seg = (Rtg.Length - 3) / 7
            If FB.Split("+").Length < Seg Then FB = FillFB_byRTG(Rtg, FB)
            If Seg > RBD.Length Then RBD = FillRBD_byRTG(Rtg, RBD)
            strQry = ""
            For j As Int16 = 1 To Seg
                If Rtg.Substring(7 * j - 7, 10).Substring(4, 2) <> "//" Then
                    strQry = strQry & "; insert Segment (TKID, Segment, RBD, FB) values (" & dTable.Rows(i)("RecID") & _
                        ",'" & Rtg.Substring(7 * j - 7, 10) & "','" & RBD.ToString.Substring(j - 1, 1) & "','" & _
                        FB.Split("+")(j - 1) & "')"
                End If
            Next
            Cmd.CommandText = strQry.Substring(1)
            If strQry.Length > 2 Then Cmd.ExecuteNonQuery()
        Next
    End Sub
End Module
