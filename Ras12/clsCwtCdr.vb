Public Class clsCwtCdr
    Dim mstrNbr As String
    Dim mstrCdrName As String
    Dim mstrCharType As String
    Dim mintMinLength As Int16
    Dim mintMaxLength As Int16
    Dim mstrFormatError As String
    Private mstrMandatory As String

    Public Function CheckFormat(ByVal strValue As String) As Boolean
        Select Case mstrCharType
            Case "ALPHA"
                If Not IsAlphaOnly(strValue) Then
                    mstrFormatError = "Value is not Alpha only!"
                    Return False
                End If
            Case "NUMERIC"
                If Not IsNumeric(strValue) Then
                    mstrFormatError = "Value is not Numeric!"
                    Return False
                End If
        End Select
        If strValue.Length < mintMinLength Then
            mstrFormatError = "Invalid Minimum Length!"
            Return False
        End If
        If strValue.Length > mintMaxLength Then
            mstrFormatError = "Invalid Maximum Length!"
            Return False
        End If
        Return True
    End Function
    Public Property Nbr() As String
        Get
            Nbr = mstrNbr
        End Get
        Set(ByVal value As String)
            mstrNbr = value
        End Set
    End Property
    Public Property CdrName() As String
        Get
            CdrName = mstrCdrName
        End Get
        Set(ByVal value As String)
            mstrCdrName = value
        End Set
    End Property
    Public Property CharType() As String
        Get
            CharType = mstrCharType
        End Get
        Set(ByVal value As String)
            mstrCharType = value
        End Set
    End Property
    Public Property FormatError() As String
        Get
            FormatError = mstrFormatError
        End Get
        Set(ByVal value As String)
            mstrFormatError = value
        End Set
    End Property

    Public Property MinLength() As Int16
        Get
            MinLength = mintMinLength
        End Get
        Set(ByVal value As Int16)
            mintMinLength = value
        End Set
    End Property
    Public Property MaxLength() As Int16
        Get
            MaxLength = mintMaxLength
        End Get
        Set(ByVal value As Int16)
            mintMaxLength = value
        End Set
    End Property
    Public Property Mandatory() As String
        Get
            Mandatory = mstrMandatory
        End Get
        Set(ByVal value As String)
            mstrMandatory = value
        End Set
    End Property

End Class
