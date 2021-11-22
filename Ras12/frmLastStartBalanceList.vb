Public Class frmStartBalanceList
    Private mblnRefreshSearch As Boolean
    Private Sub frmStartBalanceList_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadComboDisplay(cboCustomers, "Select RecId as value, CustShortName as display from Customerlist" _
                         & " where Status<>'XX' order by CustShortName", Conn)
        mblnRefreshSearch = True
        cboCustomers.SelectedIndex = -1
        '  Search()
    End Sub

    Public Function Search() As Boolean
        Dim strQuerry As String = "Select * from ChotCongNo where Status<>'XX'"

        If cboCustomers.SelectedIndex <> -1 Then
            strQuerry = strQuerry & " and CustId=" & cboCustomers.SelectedValue
        End If
        strQuerry = strQuerry & " order by CustShortName"
        LoadDataGridView(dgStartBalance, strQuerry, Conn)

    End Function

    Private Sub cboCustomers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCustomers.SelectedIndexChanged
        If mblnRefreshSearch Then
            Search()
        End If
    End Sub

    Private Sub lbkAdd_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkAdd.LinkClicked
        frmStartBalanceEdit.ShowDialog()
        Search()
    End Sub
End Class