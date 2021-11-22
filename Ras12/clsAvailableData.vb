Public Class clsAvailableData
    Dim mstrDataCode As String
    Dim mstrDataValue As String
    
    Public Property DataCode() As String
        Get
            DataCode = mstrDataCode
        End Get
        Set(ByVal value As String)
            mstrDataCode = value
        End Set
    End Property

    Public Property DataValue() As String
        Get
            DataValue = mstrDataValue
        End Get
        Set(ByVal value As String)
            mstrDataValue = value
        End Set
    End Property
End Class


