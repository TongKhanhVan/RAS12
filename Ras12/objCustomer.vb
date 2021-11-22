Public Class objCustomer
    Private _CustID As Integer, _DueIn As Int16, _DayToCFM As Int16
    Private _ShortName As String, _FullName As String, _Addr As String, _TaxCode As String, _KyBaoCao As String
    Private _CustType As String, _DelayType As String, _AdhType As String, _LstReconcile As Date, _eMail As String, _OID As String
    Private _CrditOnAL As String, _CustOf_AL As String, _FoxCoef As Decimal, _MinBLC As Decimal, _Coef As Decimal
    Private _SIR As String, _HasSIR As Boolean, _TayBa As String, _CurrBLC As Decimal, _LstDue As Date, _isOverDue As Boolean
    Dim _CnStr As String
    Private _List_TA As String, _List_TO As String, _List_CA As String, _List_CS As String, _List_CT As String _
        , _List_WK As String, _List_All As String
    Private _List_CC As String, _List_CR As String, _List_PP As String, _List_Fox As String, _List_CWT As String
    Private _List_LC As String, _VAT_Company As String
    ReadOnly Property IsOverDue() As Boolean
        Get
            Return _isOverDue
        End Get
    End Property
    ReadOnly Property DayToCfm() As Integer
        Get
            Return _DayToCFM
        End Get
    End Property

    ReadOnly Property DueIn() As Integer
        Get
            Return _DueIn
        End Get
    End Property
    ReadOnly Property VAT_Company() As String
        Get
            Return _VAT_Company
        End Get
    End Property
    ReadOnly Property KyBaoCao() As String
        Get
            Return _KyBaoCao
        End Get
    End Property


    ReadOnly Property List_CWT() As String
        Get
            Return _List_CWT
        End Get
    End Property
    Property CnStr() As String
        Get
            Return _CnStr
        End Get
        Set(ByVal sCnStr As String)
            _CnStr = sCnStr
        End Set
    End Property

    ReadOnly Property TayBa() As String
        Get
            Return _TayBa
        End Get
    End Property

    ReadOnly Property List_Fox() As String
        Get
            Return _List_Fox
        End Get
    End Property

    ReadOnly Property List_CR() As String
        Get
            Return _List_CR
        End Get
    End Property
    ReadOnly Property List_PP() As String
        Get
            Return _List_PP
        End Get
    End Property
    ReadOnly Property List_LC() As String
        Get
            Return _List_LC
        End Get
    End Property

    ReadOnly Property List_TA() As String
        Get
            Return _List_TA
        End Get
    End Property
    ReadOnly Property List_TO() As String
        Get
            Return _List_TO
        End Get
    End Property
    ReadOnly Property List_CA() As String
        Get
            Return _List_CA
        End Get
    End Property
    ReadOnly Property List_CS() As String
        Get
            Return _List_CS
        End Get
    End Property
    ReadOnly Property List_CT() As String
        Get
            Return _List_CT
        End Get
    End Property
    ReadOnly Property List_WK() As String
        Get
            Return _List_WK
        End Get
    End Property
    ReadOnly Property List_CC() As String
        Get
            Return _List_CC
        End Get
    End Property
    ReadOnly Property List_All() As String
        Get
            Return _List_All
        End Get
    End Property
    ReadOnly Property FoxCoef() As Decimal
        Get
            Return _FoxCoef
        End Get
    End Property
    ReadOnly Property Coef() As Decimal
        Get
            Return _Coef
        End Get
    End Property

    ReadOnly Property CurrBLC() As Decimal
        Get
            Return _CurrBLC
        End Get
    End Property

    ReadOnly Property MinBLC() As Decimal
        Get
            Return _MinBLC
        End Get
    End Property

    ReadOnly Property HasSIR() As Boolean
        Get
            Return _HasSIR
        End Get
    End Property

    ReadOnly Property SIR() As String
        Get
            Return _SIR
        End Get
    End Property

    ReadOnly Property Email() As String
        Get
            Return _eMail
        End Get
    End Property

    ReadOnly Property OID() As String
        Get
            Return _OID
        End Get
    End Property
    ReadOnly Property LstDue() As Date
        Get
            Return _LstDue
        End Get
    End Property

    ReadOnly Property LstReconcile() As Date
        Get
            Return _LstReconcile
        End Get
    End Property
    ReadOnly Property CrditOnAL() As String
        Get
            Return _CrditOnAL
        End Get
    End Property
    ReadOnly Property CustOf_AL() As String
        Get
            Return _CustOf_AL
        End Get
    End Property
    ReadOnly Property AdhType() As String
        Get
            Return _AdhType
        End Get
    End Property

    Property CustID() As Integer
        Get
            Return _CustID
        End Get
        Set(ByVal iCustID As Integer)
            _CustID = iCustID
            _VAT_Company = ""
            _SIR = ""
            _CustOf_AL = ""
            _OID = ""
            _Coef = 0
            _MinBLC = 0
            _FoxCoef = 0
            _CrditOnAL = ""
            _AdhType = ""
            _TayBa = ""
            If _CustID = 0 Then Exit Property
            Dim dTable As DataTable

            dTable = GetDataTable("select CustShortName, CustFullName,CustAddress, CustTaxCode, Email from CustomerList " & _
                                  "where status<>'XX' and recID=" & _CustID)
            For i As Int16 = 0 To dTable.Rows.Count - 1
                _ShortName = dTable.Rows(i)("CustShortName")
                _FullName = dTable.Rows(i)("CustFullName")
                _FullName = _FullName.Replace("'", "")
                _Addr = dTable.Rows(i)("CustAddress")
                _Addr = _Addr.Replace("'", "")
                _TaxCode = dTable.Rows(i)("CustTaxCode")
                _eMail = dTable.Rows(i)("Email")
            Next
            _CustType = ScalarToString("Cust_detail", "VAL", "Custid=" & _CustID & " and status='OK' and cat='Channel'")
            dTable = GetDataTable("select val, val1 from cust_Detail where status='OK' and cat='EMAIL' and custid=" & _CustID)
            For i As Int16 = 0 To dTable.Rows.Count - 1
                If InStr(_eMail, dTable.Rows(i)("Val")) = 0 Then
                    _eMail = _eMail & ";" & dTable.Rows(i)("VAL1") & ": " & dTable.Rows(i)("VAL")
                End If
            Next
            If _ShortName <> "" Then
                dTable = GetDataTable("select CAT, VAL from Cust_Detail where custid=" & _CustID & " and status <>'XX'")
                For i As Int16 = 0 To dTable.Rows.Count - 1
                    If dTable.Rows(i)("CAT") = "AL" Then _CustOf_AL = _CustOf_AL & "_" & dTable.Rows(i)("VAL")
                    If dTable.Rows(i)("CAT") = "SIR" Then _SIR = _SIR & "|" & dTable.Rows(i)("VAL")
                Next
                If _SIR.Length > 1 Then _SIR = _SIR.Substring(1)
                If _CustOf_AL.Length > 1 Then _CustOf_AL = _CustOf_AL.Substring(1)
                _CustOf_AL = _CustOf_AL.Replace("_", "','")
                _CustOf_AL = "('" & _CustOf_AL & "')"

                dTable = GetDataTable("select CRCoef, PPCoef, Adh, FoxCoef, MinBLC, AL, ADH, VAT " & _
                                      "from CC_Setting where status='OK' and  custid=" & _CustID)
                _DelayType = "DEB"
                _CurrBLC = 0
                _isOverDue = False
                _DueIn = 0
                _KyBaoCao = ""
                _DayToCFM = 0
                _LstDue = Now.Date
                If dTable.Rows.Count = 1 Then
                    _DelayType = IIf(dTable.Rows(0)("CRCoef") > 0, "PSP", "PPD")
                    _Coef = IIf(dTable.Rows(0)("CRCoef") > 0, dTable.Rows(0)("CRCoef") > 0, dTable.Rows(0)("PPCoef") > 0)
                    _MinBLC = dTable.Rows(0)("MinBLC")
                    _FoxCoef = dTable.Rows(0)("FoxCoef")
                    _CrditOnAL = dTable.Rows(0)("AL")
                    _AdhType = dTable.Rows(0)("ADH")
                    _VAT_Company = dTable.Rows(0)("VAT")
                End If
                If InStr("PPD_PSP", _DelayType) > 0 Then
                    _CurrBLC = ScalarToDec("CC_BLC", IIf(_DelayType = "PPD", "top 1 VND_PPD_Avail", "top 1 VND_PSP_Avail"), " custID=" & _CustID & " order by recID desc")
                    If _DelayType = "PSP" Then
                        dTable = GetDataTable("select * from KyBaoCao where status='OK' and  custid=" & _CustID)
                        If dTable.Rows.Count = 1 Then
                            _isOverDue = dTable.Rows(0)("OverDue")
                            _DueIn = dTable.Rows(0)("DueIn")
                            _DayToCFM = dTable.Rows(0)("DaysToCfm")
                            _KyBaoCao = dTable.Rows(0)("Periods")
                            _KyBaoCao = _KyBaoCao.Replace("/0", "")
                        End If
                    End If
                    If _isOverDue Then
                        _LstDue = ScalarToDate("GhiNoKhach", "top 1 DueDate", " custID=" & _CustID & " and status <>'XX' and DueDate <getdate() and ConNo>0")
                    End If
                End If
            End If
            If InStr("PSP_PPD", _DelayType) > 0 Then
                _LstReconcile = ScalarToDate("ChotCongNo", "top 1 AsOf", "custid=" & _CustID & " and status <>'XX' order by asof desc")

                dTable = GetDataTable("select OfficeID from OfficeID where status in ('QQ','OK') and CustID=" & _CustID)
                For i As Int16 = 0 To dTable.Rows.Count - 1
                    _OID = _OID & "_" & dTable.Rows(i)("OfficeID")
                Next
                If _OID.Length > 2 Then _OID = _OID.Substring(1)
            End If

            dTable = GetDataTable("select substring(cat,9,2)+VAl as TAYBA from cust_Detail where custid=" & _CustID & _
                                  " and status ='OK' and left(cat,4)='TAY3' order by VAL")
            For i As Int16 = 0 To dTable.Rows.Count - 1
                _TayBa = _TayBa & "_" & dTable.Rows(i)("TAYBA")
            Next
            If _TayBa.Length > 2 Then _TayBa = TayBa.Substring(1)
        End Set
    End Property
    ReadOnly Property ShortName() As String
        Get
            Return _ShortName
        End Get
    End Property

    ReadOnly Property FullName() As String
        Get
            Return _FullName
        End Get
    End Property

    ReadOnly Property taxCode() As String
        Get
            Return _TaxCode
        End Get
    End Property
    ReadOnly Property Addr() As String
        Get
            Return _Addr
        End Get
    End Property

    Property CustType() As String
        Get
            Return _CustType
        End Get
        Set(ByVal sCustType As String)
            _CustType = sCustType
        End Set
    End Property

    ReadOnly Property DelayType() As String
        Get
            Return _DelayType
        End Get
    End Property
    Public Function AddCustomer(ByVal pCustShortName As String, ByVal pCustFullName As String, ByVal pCustTaxcode As String, ByVal pCustAddress As String, ByVal pEmail As String, ByVal pPhone As String, pstatus As String) As Integer
        Dim KQ As Integer
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "Insert into Customer (CustShortName, CustFullName, CustTaxcode, CustAddress, email, Phone, City, FstUser,Status) " & _
            "values (@CustShortName, @CustFullName, @CustTaxcode, @CustAddress, @email, @Phone, @City, @FstUser,@Status)" & _
            "; SELECT SCOPE_IDENTITY() AS [RecID]"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@CustShortName", SqlDbType.VarChar).Value = pCustShortName
        cmd.Parameters.Add("@CustFullName", SqlDbType.NVarChar).Value = pCustFullName
        cmd.Parameters.Add("@CustTaxcode", SqlDbType.VarChar).Value = pCustTaxcode
        cmd.Parameters.Add("@CustAddress", SqlDbType.NVarChar).Value = pCustAddress
        cmd.Parameters.Add("@email", SqlDbType.VarChar).Value = pEmail
        cmd.Parameters.Add("@Phone", SqlDbType.VarChar).Value = pPhone
        cmd.Parameters.Add("@Status", SqlDbType.VarChar).Value = pstatus
        cmd.Parameters.Add("@City", SqlDbType.VarChar).Value = MySession.City
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        KQ = cmd.ExecuteScalar
        Return KQ
    End Function
    Public Sub InsertCustDetail(ByVal pCustID As Integer, ByVal CAT As String, ByVal Val As String, pXoa As Boolean)
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        If pXoa Then
            cmd.CommandText = ChangeStatus_ByDK("cust_Detail", "XX", "custId=" & pCustID & " and cat='" & CAT & "' and status='OK'") & "; "
        End If
        cmd.CommandText = cmd.CommandText & " insert cust_Detail (CustID, CAT, VAL, FstUser) values (" & pCustID & ",'" & CAT & "','" & Val & "','" & myStaff.SICode & "')"
        cmd.ExecuteNonQuery()
    End Sub

    Public Sub SaveChange(ByVal pFullName As String, ByVal pTaxCode As String, ByVal pEmail As String, ByVal pPhone As String, ByVal pAddrress As String, ByVal pCustID As Integer, pLocation As String)
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand

        cmd.CommandText = UpdateLogFile("Customerlist", "Edit", pFullName, pTaxCode, pEmail, pPhone, pAddrress, pCustID)
        cmd.ExecuteNonQuery()
        cmd.CommandText = "update CustomerList set CustFullName=@CustFullName, CustTaxCode = @CustTaxCode, Email=@Email, Phone=@Phone, " & _
            " CustAddress=@Addrress, Location=@Location where RecID=@CustID"
        cmd.Parameters.Add("@CustFullName", SqlDbType.NVarChar).Value = pFullName
        cmd.Parameters.Add("@CustTaxCode ", SqlDbType.VarChar).Value = pTaxCode
        cmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = pEmail
        cmd.Parameters.Add("@Phone", SqlDbType.VarChar).Value = pPhone
        cmd.Parameters.Add("@Location", SqlDbType.VarChar).Value = pLocation
        cmd.Parameters.Add("@Addrress", SqlDbType.NVarChar).Value = pAddrress
        cmd.Parameters.Add("@Custid", SqlDbType.Int).Value = pCustID
        cmd.ExecuteNonQuery()
    End Sub

    Public Sub GenCustList()
        _List_All = "select RecID as VAL, CustShortName as DIS  from CustomerList l where status='OK' "
        Dim StrSQL As String = _List_All & " and RecID in (select CustID from cust_Detail where status+CAT+VAL='OKChannel"
        _List_TA = StrSQL & "TA')"
        _List_LC = StrSQL & "LC')"
        _List_CA = StrSQL & "CA')"
        _List_TO = StrSQL & "TO')"
        _List_CS = StrSQL & "CS')"
        _List_WK = StrSQL & "WK')"
        _List_CC = "Select CustID as VAL, CustShortName as DIS from cc_setting where status='OK' "
        _List_CR = _List_CC & " and CrCoef>0 "
        _List_PP = _List_CC & " and PPCoef>0 "
        _List_Fox = _List_CC & " and FoxCoef>0 "
        _List_CWT = "(select CustID from cust_Detail where status+cat='OKChannel' and val in ('CS','LC'))"
        _List_CT = _List_All & " and RecID in (select CustID from cust_Detail where status+CAT='OKChannel' and val in ('CS','LC'))"
    End Sub
    Public Sub GetSIR(ByVal pSVC As String)
        Dim dTable As DataTable
        If InStr("HTL_INS_VSA_CAR", pSVC) = 0 Then pSVC = "AIR"
        ' khai SIR cho Cust0 neu do la Spec Infor cua loai san pham nay
        dTable = GetDataTable("select CAT, VAL from Cust_Detail where custid in (0," & _CustID & ") and status <>'XX'" & _
                              " and cat='SIR' and VAL1='" & pSVC & "'")
        _SIR = ""
        For i As Int16 = 0 To dTable.Rows.Count - 1
            _SIR = _SIR & "|" & dTable.Rows(i)("VAL")
        Next
        If _SIR.Length > 2 Then
            _SIR = _SIR.Substring(1)
            _HasSIR = True
        Else
            _HasSIR = False
        End If
    End Sub
End Class
