Public Class objAL
    Private _ALCode As String, _DocCode As String, _QuocTich As String, _DefaultCurr As String, _VAT_Kyhieu As String, _VAT_MauSo As String
    Private _Domain As String, _ValidDocCode As String
    Private _isTKTLess As Boolean, _CanIssVAT As Boolean, _isMultiCurr As Boolean, _isTVA As Boolean
    Private StrSQL As String, _cnStr As String, _TVC As String, _Counter As String
    Property Domain() As String
        Get
            Return _Domain
        End Get
        Set(ByVal sDomain As String)
            _Domain = sDomain
        End Set
    End Property
    Property Counter() As String
        Get
            Return _Counter
        End Get
        Set(ByVal sCounter As String)
            _Counter = sCounter
        End Set
    End Property

    Property CnStr() As String
        Get
            Return _CnStr
        End Get
        Set(ByVal sCnStr As String)
            _CnStr = sCnStr
        End Set
    End Property
    Property ALCode() As String
        Get
            Return _ALCode
        End Get
        Set(ByVal sALCode As String)
            If InStr("EDU_TVS_GSA", _Domain) = 0 Then MsgBox("Plz Assign Domain For this AL")
            _ALCode = sALCode
            If _ALCode = "" Then Exit Property
            Dim VAT As String, SecondCode As String = ""
            Dim dTable As DataTable
            If _Domain = "TVS" Then
                _VAT_Kyhieu = ScalarToString("airline", "kyhieu", "AL='TS'")
                _VAT_MauSo = ScalarToString("airline", "mauso", "AL='TS'")
            ElseIf _Domain <> "EDU" Then
                _VAT_Kyhieu = "XX"
                _VAT_MauSo = "000"
            End If
            StrSQL = String.Format("select * from Airline where status='OK' and AL='{0}'", _ALCode)
            dTable = GetDataTable(StrSQL)
            For i As Int16 = 0 To dTable.Rows.Count - 1
                _DocCode = dTable.Rows(i)("DocCode")
                _QuocTich = dTable.Rows(i)("Country")
                _TVC = "TVTR"
                SecondCode = dTable.Rows(i)("SecondCode")
                If _Domain <> "TVS" Then
                    _VAT_Kyhieu = dTable.Rows(i)("KyHieu")
                    _VAT_MauSo = dTable.Rows(i)("MauSo")
                    _TVC = dTable.Rows(i)("TVC")
                End If
                _DefaultCurr = dTable.Rows(i)("DefauCurr")
                VAT = dTable.Rows(i)("VAT")
                _isMultiCurr = dTable.Rows(i)("MultiCurr")
                _isTVA = dTable.Rows(i)("isTVA")
                _isTKTLess = IIf(_DocCode.Substring(0, 1) = "Z", True, False)
                _CanIssVAT = IIf(InStr(VAT, _Domain.Substring(0, 1) & "Y") > 0, True, False)

                If _ALCode = "SU" And _Counter = "CWT" Then _CanIssVAT = True ' BT SU xuat ID HAN de gom incentive nen set  GNTN, rieng quay CWT dung ID SGN de tien rebook

            Next
            If _DocCode = "" Then _DocCode = ScalarToString("lib.dbo.AirlineList", "docCode", "ALcode='" & _ALCode & "'")
            _ValidDocCode = _DocCode
            If SecondCode <> "" Then _ValidDocCode = _ValidDocCode & "_" & SecondCode
        End Set
    End Property

    ReadOnly Property DocCode() As String
        Get
            Return _DocCode
        End Get
    End Property
    ReadOnly Property ValidDocCode() As String
        Get
            Return _ValidDocCode
        End Get
    End Property

    ReadOnly Property QuocTich() As String
        Get
            Return _QuocTich
        End Get
    End Property

    ReadOnly Property DefaultCurr() As String
        Get
            Return _DefaultCurr
        End Get
    End Property
    ReadOnly Property VAT_KyHieu() As String
        Get
            Return _VAT_Kyhieu
        End Get
    End Property
    ReadOnly Property VAT_MaSo() As String
        Get
            Return _VAT_MauSo
        End Get
    End Property
    ReadOnly Property TVC() As String
        Get
            Return _TVC
        End Get
    End Property

    ReadOnly Property isTKTLess() As Boolean
        Get
            Return _isTKTLess
        End Get
    End Property

    ReadOnly Property CanIssVAT() As Boolean
        Get
            Return _CanIssVAT
        End Get
    End Property

    ReadOnly Property isMultiCurr() As Boolean
        Get
            Return _isMultiCurr
        End Get
    End Property
    ReadOnly Property isTVA() As Boolean
        Get
            Return _isTVA
        End Get
    End Property

    Public Function DocToCode(ByVal pDoc As String) As String
        Return ScalarToString("lib.dbo.AirlineList", "ALCode", " Doccode='" & pDoc & "'")
    End Function
End Class
