Public Class objTerminal
    Private _Domain As String, _Location As String, _City As String, _Counter As String, _POSCode As String, _TRXCode As String
    Private _ServerIP As String
    Property City() As String
        Get
            Return _City
        End Get
        Set(ByVal sCity As String)
            _City = sCity
            _POSCode = IIf(_City = "SGN", "0", "3")
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
    Property Domain() As String
        Get
            Return _Domain
        End Get
        Set(ByVal sDomain As String)
            _Domain = sDomain
            If _Domain = "EDU" Then
                _TRXCode = "XX"
            ElseIf _Domain = "TVS" Then
                _TRXCode = "TS"
            Else
                _TRXCode = "YY"
            End If
        End Set
    End Property
    Property TRXCode() As String
        Get
            Return _TRXCode
        End Get
        Set(ByVal sTRXCode As String)
            _TRXCode = sTRXCode
        End Set
    End Property
    ReadOnly Property POSCode() As String
        Get
            Return _POSCode
        End Get
    End Property
    Property ServerIP() As String
        Get
            Return _ServerIP
        End Get
        Set(ByVal sServerIP As String)
            _ServerIP = sServerIP
        End Set

    End Property

    Property Location() As String
        Get
            Return _Location
        End Get
        Set(ByVal sLocation As String)
            _Location = sLocation
        End Set
    End Property

End Class
