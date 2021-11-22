Public Class frmLinkTourItems
    Private mobjSelectedRow As DataGridViewRow
    Private Sub frmLinkTourItems_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strQuerry As String
        With mobjSelectedRow
            strQuerry = "select * from DuToan_Item where Status<>'XX' and RelatedItem=0 and DuToanID=" _
                        & .Cells("DutoanId").Value & " and RecId <>" & .Cells("RecId").Value
            If .Cells("Service").Value = "TransViet  SVC Fee" Then
                strQuerry = strQuerry & " and Service <>'TransViet SVC Fee'"
            Else
                strQuerry = strQuerry & " and Service ='TransViet SVC Fee'"
            End If
        End With

        LoadDataGridView(grdItems, strQuerry, Conn)
    End Sub

    Public Sub New(objSelectedRow As DataGridViewRow)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mobjSelectedRow = objSelectedRow
    End Sub

    Private Sub lbkCancel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkCancel.LinkClicked
        Me.Dispose()
    End Sub

    Private Sub lbkOK_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkOK.LinkClicked
        With grdItems
            If .CurrentRow IsNot Nothing Then
                Dim intMainId As Integer
                Dim strRelatedItems As String = "(" & .CurrentRow.Cells("RecId").Value _
                                                & "," & mobjSelectedRow.Cells("RecId").Value & ")"

                If mobjSelectedRow.Cells("Service").Value = "TransViet SVC Fee" Then
                    intMainId = .CurrentRow.Cells("RecId").Value
                Else
                    intMainId = mobjSelectedRow.Cells("RecId").Value
                End If

                ExecuteNonQuerry("Update Dutoan_Item set RelatedItem=" & intMainId _
                                 & " where RecId in " & strRelatedItems, Conn)
            End If
        End With
    End Sub
End Class