Public Class frmCcInfoEdit
    Public Sub New(strCustShortName As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        txtCustShortName.Text = strCustShortName


    End Sub
    Private Sub lbkSave_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkSave.LinkClicked
        Dim strQuerry As String = ""
        Dim intNewRecId As Integer
        Dim lstQuerries As New List(Of String)
        txtCardNbr.Text = txtCardNbr.Text.Replace(" ", "")
        txtCardHolder.Text = txtCardHolder.Text.Trim
        If Not CheckInputValues() Then
            Exit Sub
        End If
        lstQuerries.Add("Insert into Misc (Cat, Val, Val1, Val2, Details,Description, Status,FstUser) values ('CcNbr','" _
                        & txtCustShortName.Text & "','" & txtCardHolder.Text & "','" & cboCardType.Text _
                        & "','" & Mid(txtCardNbr.Text, txtCardNbr.Text.Length - 3) & "','" & txtExpDate.Text _
                        & "','--','" & myStaff.SICode & "')")
        If UpdateListOfQuerries(lstQuerries, Conn, True, intNewRecId) Then
            lstQuerries.Clear()
            lstQuerries.Add("Insert into Misc (Cat, Val,Val1,Val2, Details) values ('PrefixC','" & intNewRecId _
                             & "','" & Mid(txtCardNbr.Text, 1, txtCardNbr.Text.Length - 4) _
                            & "','" & cboBiz.Text & "','" & txtRemark.Text & "')")
            lstQuerries.Add("Update Misc set Status='OK' where RecId=" & intNewRecId)
            If UpdateListOfQuerries(lstQuerries, Conn) Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Dispose()
            Else
                ExecuteNonQuerry("delete GO_MiscWzDate where RecId=" & intNewRecId, Conn)
                MsgBox("Unablle to Add CreditCard")
            End If
        Else
            MsgBox("Unable to Add CcNbr")
        End If
    End Sub

    Private Function CheckInputValues() As Boolean
        Dim strCardTypeByNbr As String
        Dim arrExp As String()
        If Not CheckFormatTextBox(txtCardHolder, , 3) Then
            Return False
        ElseIf Not CheckFormatTextBox(txtCardNbr, True, 15, 16) Then
            Return False
        End If
        If Not CheckFormatComboBox(cboBiz, , 3, 3) Then
            Return False
        End If
        strCardTypeByNbr = GetCardTypeByCardNbr(txtCardNbr.Text)
        If strCardTypeByNbr = "XX" Then
            MsgBox("Invalid Card Nbr")
            Return False
        ElseIf strCardTypeByNbr <> cboCardType.Text Then
            MsgBox("Unmatched Card Nbr >< Card Type")
            Return False
        End If
        arrExp = txtExpDate.Text.Split("/")
        If arrExp.Length <> 2 Then
            MsgBox("Invalid ExpDate")
            Return False
        ElseIf Not IsNumeric(arrExp(0)) Then
            MsgBox("Invalid ExpDate")
            Return False
        ElseIf Not IsNumeric(arrExp(1)) Then
            MsgBox("Invalid ExpDate")
            Return False
        ElseIf CInt(arrExp(1) & arrExp(0)) < Format(Now, "yyMM") Then
            MsgBox("Invalid ExpDate")
            Return False
        End If
        'check dup xem da nhap chua
        Return True
    End Function

    Private Sub frmCcInfoEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cboBiz.SelectedIndex = 0
    End Sub
End Class