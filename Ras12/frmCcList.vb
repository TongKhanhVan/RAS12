Public Class frmCcList
    
    Private Sub frmCcList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strCustShortName As String = "Select Val as value from misc where Cat='CcAccess' and Status='OK' and Val1='" _
                                         & myStaff.SICode & "' order by Val"
        LoadCombo(cboCustShortName, strCustShortName, Conn)
        If cboCustShortName.Items.Count > 0 Then
            cboCustShortName.SelectedIndex = 0
        End If
        cboStatus.SelectedIndex = 0
    End Sub
    Private Function Search() As Boolean
        Dim strQuerry As String
        Dim strCcNbr As String

        If chkHideCcNbr.Checked Then
            strCcNbr = ",m.Details as Last4Digit"
        Else
            strCcNbr = ",(m2.Val1+m.Details) as CardNbr"
        End If

        strQuerry = "Select m.RecId, m.Val as CustShortName,m.Val1 as CardHolder,m.Val2 as CardType" _
                    & strCcNbr & ",m.Description as ExpDate" _
                    & " ,m2.Val2 as Biz,m2.Description as Remark" _
                    & " from  Misc m" _
                    & " LEFT JOIN  Misc m2 on cast(m.RecId as varchar) =m2.Val" _
                    & " where m.Cat='CcNbr' and m.Status<>'XX' and m2.Cat='PrefixC'"

        If cboStatus.SelectedIndex > -1 Then
            strQuerry = strQuerry & " and m.Status='" & cboStatus.Text & "'"
        End If

        If cboCustShortName.SelectedIndex > -1 Then
            strQuerry = strQuerry & " and m.Val='" & cboCustShortName.Text & "'"
        End If
        strQuerry = strQuerry & " order by m.Val1"
        LoadDataGridView(dgCcList, strQuerry, Conn)
        Return True
    End Function
    
    Private Sub cboCustShortName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCustShortName.SelectedIndexChanged
        If cboCustShortName.SelectedIndex > -1 Then
            Search()
        End If
    End Sub

    Private Sub cboStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStatus.SelectedIndexChanged
        Search()
    End Sub

    Private Sub lbkNew_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkNew.LinkClicked
        Dim frmNew As New frmCcInfoEdit(cboCustShortName.Text)
        If frmNew.ShowDialog = Windows.Forms.DialogResult.OK Then
            Search()
        End If
    End Sub

    Private Sub lbkExpire_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkExpire.LinkClicked
        If dgCcList.CurrentRow IsNot Nothing Then
            Dim strQuerry As String = "Update Misc set Status='EX',LstUpdate=getdate(), LstUser='" _
                                      & myStaff.SICode & "' where RecId=" & dgCcList.CurrentRow.Cells("RecId").Value
            If ExecuteNonQuerry(strQuerry, Conn) Then
                Search()
            Else
                MsgBox("Unable to Expire CreditCard")
            End If
        End If
    End Sub

    Private Sub chkHideCcNbr_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideCcNbr.CheckedChanged
        Search()
    End Sub
End Class