Public Class ucRqDataText
    Private mobjRow As DataRow
    Private mstrCurrentValue As String
    Public Function CheckInput() As String
        'return error
        Dim strError As String = String.Empty
        If mobjRow("Mandatory") = "M" Then
            If txtValue.Text = "" Then
                strError = mobjRow("NameByCustomer") & ": MissingValue"
            ElseIf mobjRow("AllowSpecialValues") = 1 _
                AndAlso IsSpecialValue(mobjRow("CustId"), mobjRow("AllowSpecialValue"), txtValue.Text) Then
                Return ""
            ElseIf txtValue.Text.Length < mobjRow("MinLength") Then
                strError = mobjRow("NameByCustomer") & ": Invalid MinLength"
            ElseIf txtValue.Text.Length > mobjRow("MaxLength") Then
                strError = mobjRow("NameByCustomer") & ": Invalid MaxLength"
            ElseIf mobjRow("CharType") = "NUMERIC" AndAlso Not IsNumeric(txtValue.Text) Then
                strError = mobjRow("NameByCustomer") & ": Value must be Numeric"
            End If
        End If
        Return strError
    End Function
    Private Sub txtValue_TextChanged(sender As Object, e As EventArgs) Handles txtValue.TextChanged
        mstrCurrentValue = txtValue.Text
    End Sub
    Public Property Row As DataRow
        Get
            Return mobjRow
        End Get
        Set(value As DataRow)
            mobjRow = value
        End Set
    End Property
    Public Property CurrentValue As String
        Get
            Return mstrCurrentValue
        End Get
        Set(value As String)
            mstrCurrentValue = value
        End Set
    End Property
End Class
