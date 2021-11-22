Public Class frmStartBalanceEdit

    Private Sub frmStartBalanceEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadComboDisplay(cboCustomers, "Select RecId as value, CustShortName as display from Customerlist" _
                 & " where Status<>'XX' order by CustShortName", Conn)
        cboCustomers.SelectedIndex = 0
    End Sub

    Private Sub lbkSave_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkSave.LinkClicked
        If Not CheckInputValues() Then
            Exit Sub
        End If
        Dim strQuerry As String = "Insert into ChotCongNo (CustId,CustShortName,VND_Avail,AsOf,FstUser) values(" _
                                  & cboCustomers.SelectedValue & ",'" & cboCustomers.Text & "'," & txtVND_Avail.Text _
                                  & ",'" & CreateFromDate(dtpAsOf.Value) & "','" & myStaff.SICode & "')"
        If Not ExecuteNonQuerry(strQuerry, Conn) Then
            MsgBox("Unable to inert new Start Balance!")
        Else
            DialogResult = Windows.Forms.DialogResult.OK
        End If
    End Sub
    Private Function CheckInputValues() As Boolean
        Dim intCount As Integer

        If Not IsNumeric(txtVND_Avail.Text) Then
            MsgBox("Invalid VND_Avail!")
            Return False
        End If

        intCount = ScalarToInt("ChotCongNo", "Count(RecId)", " where AsOf>'" & CreateFromDate(dtpAsOf.Value) & "'")
        If intCount > 0 Then
            MsgBox("New As Of must after last As Of for the same Customer!")
            Return False
        End If
        Return True
    End Function
End Class