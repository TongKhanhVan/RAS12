Public Class clsRequiredData
    Dim mstrDataCode As String
    Dim mstrNameByCustomer As String
    Dim mintMinLength As Integer
    Dim mintMaxLength As Integer
    Dim mstrMandatory As String
    Dim mstrConditionOfUse As String
    Dim mstrCharType As String
    Dim mstrCollectionMethod As String
    Dim mstrDefaultValue As String
    Dim mblnCheckValues As Boolean
    Dim mblnAllowSpecialValues As Boolean
    Dim mstrAvailableValue As String
    Dim mstrErrMsg As String
    Public Sub ClearOldCheck()
        mstrAvailableValue = ""
        mstrErrMsg = ""
    End Sub
    Public Property DataCode() As String
        Get
            DataCode = mstrDataCode
        End Get
        Set(ByVal value As String)
            mstrDataCode = value
        End Set
    End Property

    Public Property NameByCustomer() As String
        Get
            NameByCustomer = mstrNameByCustomer
        End Get
        Set(ByVal value As String)
            mstrNameByCustomer = value
        End Set
    End Property
    Public Property MinLength() As Integer
        Get
            MinLength = mintMinLength
        End Get
        Set(ByVal value As Integer)
            mintMinLength = value
        End Set
    End Property
    Public Property MaxLength() As Integer
        Get
            MaxLength = mintMaxLength
        End Get
        Set(ByVal value As Integer)
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
    Public Property ConditionOfUse() As String
        Get
            ConditionOfUse = mstrConditionOfUse
        End Get
        Set(ByVal value As String)
            mstrConditionOfUse = value
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
    Public Property CollectionMethod() As String
        Get
            CollectionMethod = mstrCollectionMethod
        End Get
        Set(ByVal value As String)
            mstrCollectionMethod = value
        End Set
    End Property
    Public Property DefaultValue() As String
        Get
            DefaultValue = mstrDefaultValue
        End Get
        Set(ByVal value As String)
            mstrDefaultValue = value
        End Set
    End Property
    Public Property CheckValues() As Boolean
        Get
            CheckValues = mblnCheckValues
        End Get
        Set(ByVal value As Boolean)
            mblnCheckValues = value
        End Set
    End Property
    Public Property AllowSpecialValues() As Boolean
        Get
            AllowSpecialValues = mblnAllowSpecialValues
        End Get
        Set(ByVal value As Boolean)
            mblnAllowSpecialValues = value
        End Set
    End Property
    Public Property AvailableValue() As String
        Get
            AvailableValue = mstrAvailableValue
        End Get
        Set(ByVal value As String)
            mstrAvailableValue = value
        End Set
    End Property
    Public Property ErrMsg() As String
        Get
            ErrMsg = mstrErrMsg
        End Get
        Set(ByVal value As String)
            mstrErrMsg = value
        End Set
    End Property

End Class

