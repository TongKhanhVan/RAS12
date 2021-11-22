Public Class clsROE
    Private mintID As Integer
    Private mdecAmount As Decimal

    Public Property Id As Int16
        Get
            Return mintID
        End Get
        Set(value As Int16)
            mintID = value
        End Set
    End Property
    Public Property Amount As Decimal
        Get
            Return mdecAmount
        End Get
        Set(value As Decimal)
            mdecAmount = value
        End Set
    End Property
End Class