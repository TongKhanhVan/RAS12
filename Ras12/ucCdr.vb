Public Class ucCdr
    Dim mstrCharType As String
    Dim mintMinLength As Int16
    Dim mstrMandatory As String
    Public Sub New(ByVal objCdr As clsCwtCdr, Optional strValue As String = "")

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        With objCdr
            lblName.Text = .Nbr & "." & .CdrName
            lblName.Tag = .Nbr
            txtValue.Tag = .CdrName
            txtValue.MaxLength = .MaxLength
            txtValue.Text = strValue
            mstrCharType = .CharType
            mintMinLength = .MinLength
            mstrMandatory = .Mandatory
        End With
    End Sub
    Public Property CharType() As String
        Get
            CharType = mstrCharType
        End Get
        Set(ByVal value As String)
            mstrCharType = value
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
    Public Property Mandatory() As String
        Get
            Mandatory = mstrMandatory
        End Get
        Set(ByVal value As String)
            mstrMandatory = value
        End Set
    End Property
End Class
