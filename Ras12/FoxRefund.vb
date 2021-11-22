Public Class FoxRefund
    Private Sub FoxRefund_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        Dim strSQL As String
        strSQL = "Select CustID, CustShortName, DOI, TKNO, PaxName, Itinerary " & _
            " from tkT_1A t inner join customerlist c on t.custid=c.recid where  " & _
            " srv in ('R','V') and left(tkno,3)='738' and custid in (select CustID from cc_Setting)" & _
            " and tkno NOT In (select tkno from tkt where status <>'XX' and srv in ('R','V')) "
        Me.GridRef.DataSource = GetDataTable(strSQL)
        Me.GridRef.Columns("CustID").Width = 64
        Me.GridRef.Columns("DOI").Width = 64
        Me.GridRef.Columns("Itinerary").Width = 128
    End Sub
End Class