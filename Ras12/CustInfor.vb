Public Class CustInfor
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Sub LblShowFull_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblShowFull.LinkClicked
        LoadGridCust("CustShortName='" & Me.TxtSearch.Text & "'")
    End Sub

    Private Sub LblFindShort_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblFindShort.LinkClicked
        LoadGridCust("CustFullName like '%" & Me.TxtSearch.Text & "%'")
    End Sub
    Private Sub LoadGridCust(ByVal pDK As String)
        Dim strSQL As String
        strSQL = "Select RecID , CustShortName as ShortName, CustFullName as FullName from customerlist where status='OK' and " & pDK
        Me.GridCust.DataSource = GetDataTable(strSQL)
        Me.GridCust.Columns(0).Visible = False
        Me.GridCust.Columns(1).Width = 64
        Me.GridCust.Columns(2).Width = 256
        Me.LblViewPolicy.Enabled = False
    End Sub

    Private Sub GridCust_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCust.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblViewPolicy.Enabled = True
    End Sub

    Private Sub LblViewPolicy_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblViewPolicy.LinkClicked
        Dim strSQL As String
        Dim DKValidity As String = "and '" & Now.Date & "' between validFrom and ValidThru"
        Dim StrDK As String = " where status='OK' and SBU+AL+Channel+CustLevel in ("
        StrDK = StrDK & " select SBU+AL+Channel+CustLevel from Cust_Channel_Level where custid="
        StrDK = StrDK & Me.GridCust.CurrentRow.Cells("RecID").Value & DKValidity & " and sbu='" & MySession.Domain & "' and status='OK')"
        StrDK = StrDK & DKValidity

        strSQL = "Select AL, Channel, CustLevel, DType, VAL, Base, ValidFrom, ValidThru, FareSource "
        strSQL = strSQL & " from Cust_Discount " & StrDK
        Me.GridDiscount.DataSource = GetDataTable(strSQL)
        For c As Int16 = 0 To Me.GridDiscount.Columns.Count - 7
            Me.GridDiscount.Columns(c).Width = 36
        Next
        For c As Int16 = Me.GridDiscount.Columns.Count - 6 To Me.GridDiscount.Columns.Count - 1
            Me.GridDiscount.Columns(c).Width = 0
        Next
        Me.GridDiscount.Columns("VAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        strSQL = "Select AL,TRXtype, Base, FareFrom, FareThrough, SFTYpe, VAL, MinVal, MaxVal,Area,INF, "
        strSQL = strSQL & "ValidFrom, ValidThru, Refundable "
        strSQL = strSQL & " from ServiceFee " & StrDK
        Me.GridSF.DataSource = GetDataTable(strSQL)

        For c As Int16 = 0 To Me.GridSF.Columns.Count - 4
            Me.GridSF.Columns(c).Width = 38
        Next
        For c As Int16 = Me.GridSF.Columns.Count - 3 To Me.GridSF.Columns.Count - 1
            Me.GridSF.Columns(c).Width = 64
        Next
        Me.GridSF.Columns("FareFrom").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("FareThrough").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("VAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("MinVal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("MaxVal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    Private Sub CustInfor_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    Private Sub CustInfor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
    End Sub
End Class