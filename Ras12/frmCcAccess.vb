Public Class frmCcAccess
    Private mblnFirstLoad As Boolean
    Private Sub frmCcAccess_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Dim strUser As String = "Select CustShortName,CompanyName from cwt.dbo.Go_CompanyInfo1 where Status='OK' ORDER BY CompanyName,CustShortName"
        Dim strUser As String = "Select SiCode,SiName,Counter from tblUser where Status<>'XX' and Counter in ('CWT','N-A') ORDER BY SiName"
        LoadDataGridView(dgUsers, strUser, Conn)
        mblnFirstLoad = True
    End Sub
    Private Function Search() As Boolean

        If Not mblnFirstLoad Then Return False

        Dim strSelectCusst As String = "Select (case m.status when 'ok' then m.RecId else 0 end) as RecId " _
                                        & ",c.custshortname,c.CompanyName" _
                                        & ", cast((case m.status when 'ok' then 'True' else 'False' end) as bit) as Selected" _
                                        & " from CompanyInfo c" _
                                        & " left join Misc m on c.CustShortName=m.Val and m.status<>'XX' and m.Cat='CcAccess'" _
                                        & " and m.Val1='" & dgUsers.CurrentRow.Cells("SiCode").Value _
                                        & "' where c.Status<>'xx' and c.CustId>0" _
                                        & " order by Selected desc,CompanyName,c.CustShortName"
        LoadDataGridView(dgCompanies, strSelectCusst, Conn)
        Return True
    End Function

    Private Sub dgUsers_SelectionChanged(sender As Object, e As EventArgs) Handles dgUsers.SelectionChanged
        With dgUsers
            If .Rows.Count > 0 Then
                If .CurrentRow IsNot Nothing Then
                    Search()
                End If
            End If
            End With


    End Sub

    Private Sub LblSelect_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSave.LinkClicked
        Dim lstQuerries As New List(Of String)
        For Each objdgrCust As DataGridViewRow In dgCompanies.Rows
            With objdgrCust
                If .Cells("RecId").Value = 0 AndAlso .Cells("Selected").Value Then
                    lstQuerries.Add("Insert into Misc (Cat,Val,Val1,Status,FstUser) values ('CcAccess','" _
                                    & .Cells("CustShortName").Value & "','" & dgUsers.CurrentRow.Cells("SiCode").Value _
                                    & "','OK','" & myStaff.SICode & "')")

                ElseIf .Cells("RecId").Value > 0 AndAlso Not .Cells("Selected").Value Then
                    lstQuerries.Add("Update Misc  Set Status='XX',LstUpdate=Getdate(),LstUser='" _
                                    & myStaff.SICode & "' where RecId=" & .Cells("RecId").Value)
                End If
            End With
        Next
        If UpdateListOfQuerries(lstQuerries, Conn) Then
            Search()
        Else
            MsgBox("Unable to update")
        End If
    End Sub

    Private Sub lbkSelectAll_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkSelectAll.LinkClicked
        ChangeGridViewSelectedColumn(dgCompanies, True)
    End Sub

    Private Sub lbkUnselectAll_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkUnselectAll.LinkClicked
        ChangeGridViewSelectedColumn(dgCompanies, False)
    End Sub
    
End Class