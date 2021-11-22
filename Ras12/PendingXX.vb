Public Class PendingXX
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Sub PendingXX_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        If MySession.Domain = "GSA" Then Me.OptGSA.Checked = True
        If MySession.Domain = "TVS" Then Me.OptTVS.Checked = True
        LoadGridPendingXX()
        Me.TxtLyDo.Width = 262
        Me.LblChecked.Left = Me.LblInvalidTKNO.Left
        Me.LblClear.Left = Me.LblChecked.Left
    End Sub
    Private Sub LoadGridPendingXX()
        Dim strSQL As String, LstDay As Int16, FrmDate As String
        On Error GoTo exitHere
        LstDay = CInt(Me.TxtLstDay.Text)
        On Error GoTo 0
        FrmDate = Format(Now.Date.AddDays(LstDay * -1), "dd-MMM-yy")
        Me.LblInvalidTKNO.Visible = False
        Me.LblChecked.Visible = False
        strSQL = "Select t.RecID, TKNO, t.rcpNO AS TRX, t.SRV, CustType, DOI  from TKT "
        strSQL = strSQL & " t inner join RCP r on t.rcpid=r.recid where t.status='XX' and t.srv<>'V' and "
        strSQL = strSQL & " TKNO+t.SRV not in (select TKNO+SRV from tkt where status<>'XX')"
        strSQL = strSQL & " and sbu='" & IIf(Me.OptTVS.Checked, "TVS", "GSA") & "' and t.fstupdate >'"
        strSQL = strSQL & FrmDate & "' and TKNO+t.RCPNO not in (select DoWhat+F1 from ActionLog where TableName='CheckXX')"
        Me.GridPendingXX.DataSource = GetDataTable(strSQL)
        Me.GridPendingXX.Columns(3).Width = 32
        Me.GridPendingXX.Columns(0).Visible = False
        Me.GridPendingXX.Columns("CustType").Width = 32
        CheckQuyenTheoAccess()
exitHere:
        On Error GoTo 0
    End Sub

    Private Sub OptGSA_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptGSA.Click, OptTVS.Click
        LoadGridPendingXX()
    End Sub
    Private Sub CheckQuyenTheoAccess()

        For i As Int16 = 0 To Me.GridPendingXX.RowCount - 1
            If (Me.OptGSA.Checked And myStaff.AAccess <> "YY" AndAlso InStr(myStaff.AAccess, Me.GridPendingXX.Item("TRX", i).Value.ToString.Substring(0, 2)) = 0) Then
                Me.GridPendingXX.Rows(i).DefaultCellStyle.ForeColor = Color.DarkGray
            ElseIf Me.OptTVS.Checked And InStr(myStaff.CAccess, Me.GridPendingXX.Item("CustType", i).Value) = 0 Then
                Me.GridPendingXX.Rows(i).DefaultCellStyle.ForeColor = Color.DarkGray
            End If
        Next
    End Sub
    Private Sub GridPendingXX_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridPendingXX.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblInvalidTKNO.Visible = False
        Me.LblChecked.Visible = False
        Dim strSQL As String, tmpDomain As String = IIf(Me.OptTVS.Checked, "TVS", "GSA"), LstDay As Int16, FrmDate As String
        FrmDate = Format(Now.Date.AddDays(LstDay * -1), "dd-MMM-yy")
        strSQL = "Select TKNO, rcpNO AS TRX, SRV, Status, StatusAL  from TKT "
        strSQL = strSQL & " where TKNO='" & Me.GridPendingXX.CurrentRow.Cells("TKNO").Value & "'"
        Me.GridTKNO.DataSource = GetDataTable(strSQL)
        Me.GridTKNO.Columns(2).Width = 32
        Me.GridTKNO.Columns(3).Width = 32
        Me.GridTKNO.Columns(4).Width = 32
        If Me.GridPendingXX.CurrentRow.DefaultCellStyle.ForeColor <> Color.DarkGray Then
            If Me.GridPendingXX.CurrentRow.Cells("SRV").Value = "S" And Me.GridTKNO.RowCount = 1 Then
                If Me.GridPendingXX.CurrentRow.Cells("TKNO").Value.ToString.Contains("TV") Then
                    Me.LblInvalidTKNO.Visible = False
                Else
                    Me.LblInvalidTKNO.Visible = True

                End If
                Me.LblClear.Visible = False
                Me.LblChecked.Visible = Not Me.LblInvalidTKNO.Visible
            Else
                Me.LblClear.Visible = True
            End If
        End If
    End Sub
    Private Sub LblClear_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles _
        LblClear.LinkClicked, LblInvalidTKNO.LinkClicked, LblChecked.LinkClicked
        Dim lb As LinkLabel = CType(sender, LinkLabel)
        Dim Lydo As String = Me.TxtLyDo.Text
        If lb.Text = "Clear" Then
            If Lydo = "" Then Exit Sub
        Else
            Lydo = lb.Text
        End If
        cmd.CommandText = UpdateLogFile("CheckXX", Me.GridPendingXX.CurrentRow.Cells("TKNO").Value, Me.GridPendingXX.CurrentRow.Cells("TRX").Value, Lydo, "", "", "", "", "", "")
        cmd.ExecuteNonQuery()
        lb.Visible = False
    End Sub
End Class