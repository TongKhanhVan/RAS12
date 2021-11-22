Public Class unReportedTickets
    Dim strDKdate As String
    Dim strSQL As String
    Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand

    Private Sub Opt1S_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Opt1S.Click
        Dim LstRun As Date, tmpTVSI As Integer
        Dim conn_f1s As New SqlClient.SqlConnection
        conn_f1s.ConnectionString = CnStr_F1S
        Dim Cmd1s As SqlClient.SqlCommand = conn_f1s.CreateCommand
        conn_f1s.Open()
        cmd.CommandText = "select Details from [42.117.5.86].FT.dbo.MISC where RMK='" & MySession.City & "' and cat='M1S' and val='CompletionDate'"
        LstRun = cmd.ExecuteScalar
        If LstRun < Now.Date.AddDays(-1) Then
            MsgBox("Not All Data Has Been Downloaded")
        End If
        DropTableUnReportedTKT()

        cmd.CommandText = "select * into zUnreportedTKT from [42.117.5.86].ft.dbo.M1S_SrpTkts where location='" & myStaff.City & "' and rmk <>'REXC'"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "delete from zUnreportedTKT where not " & strDKdate & _
            "; delete from zUnreportedTKT where TKNO+SRV+convert(char(12), DOI) in (select TKNO+SRV+convert(char(12), DOI) from tkt where status <>'XX' or statusAL<>'XX')" & _
            "; delete from zUnreportedTKT where RelatedDoc+'R'+convert(char(12), DOI) in (select TKNO+SRV+convert(char(12), DOI) from tkt where status <>'XX')" & _
            "; delete from zUnreportedTKT where voided<>0 and srv='R'"
        cmd.ExecuteNonQuery()
        strSQL = "select DOI, OffcID, TKNO, SRV, RLOC, 0 as TVSI, RelatedDoc, '' as Counter from zUnreportedTKT"
        Me.GridUnRptTix.DataSource = GetDataTable(strSQL)
        For i As Int16 = 0 To Me.GridUnRptTix.RowCount - 1
            If Me.GridUnRptTix.Item("SRV", i).Value = "V" Then
                Cmd1s.CommandText = "Select TVSI from f1s.dbo.F1S_Inout where Cmd='WV" & _
                    Replace(Me.GridUnRptTix.Item("TKNO", i).Value, " ", "") & "' and Output like '%VOIDED%'"
            Else
                Cmd1s.CommandText = "Select TVSI from f1s.dbo.F1S_PNR_TKT_Agent where SRV='" & _
                    Me.GridUnRptTix.Item("SRV", i).Value & "' and DocNo='"
                If Me.GridUnRptTix.Item("SRV", i).Value = "R" Then
                    Cmd1s.CommandText = Cmd1s.CommandText & Replace(Me.GridUnRptTix.Item("RelatedDoc", i).Value, "", "") & "'"
                Else
                    Cmd1s.CommandText = Cmd1s.CommandText & Replace(Me.GridUnRptTix.Item("TKNO", i).Value, " ", "") & "'"
                End If
            End If
            tmpTVSI = Cmd1s.ExecuteScalar
            If tmpTVSI > 0 Then
                Me.GridUnRptTix.Item("TVSI", i).Value = tmpTVSI
                Cmd1s.CommandText = "select Office from f1s_TVSI where recid=" & tmpTVSI
                Me.GridUnRptTix.Item("Counter", i).Value = Cmd1s.ExecuteScalar
            End If
        Next
        conn_f1s.Close()
        conn_f1s.Dispose()
    End Sub

    Private Sub Opt1A_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Opt1A.Click
        Dim LstRun As Date = ScalarToDate("tvcs.dbo.MISC", "Details", "CAT ='RasChecker' and VAL ='CompletionDate' and RMK='" & MySession.City & "'")
        If LstRun < Now.Date.AddDays(-1) Then
            MsgBox("Not All Data Has Been Downloaded")
        End If
        Dim i As Integer
        Dim strSRV As String = "SRV"
        Dim arrQuerries(0 To 2) As String
        Dim strSrvFilter As String

        strSQL = String.Empty
        For i = 0 To strSRV.Length - 1
            strSrvFilter = " and Srv='" & strSRV.Chars(i) & "'"
            arrQuerries(i) = "Select * from Srp1a where " & strDKdate & strSrvFilter _
                & " and tkno NOT in " _
                & " (select tkno from Tkt where (status='OK' or statusal='OK') and " & strDKdate & strSrvFilter & ")"
        Next
        strSQL = Join(arrQuerries, " UNION ") & " order by DOI"
        Me.GridUnRptTix.DataSource = GetDataTable(strSQL)
    End Sub
    Private Sub DropTableUnReportedTKT()
        cmd.CommandText = "drop table zUnreportedTKT"
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub unReportedTickets_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DropTableUnReportedTKT()
        Me.Dispose()
    End Sub

    Private Sub unReportedTickets_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CmbLstNdays.Text = "8"
    End Sub

    Private Sub CmbLstNdays_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbLstNdays.SelectedIndexChanged
        Dim Frm As Date, Thru As Date
        Thru = Now.Date.AddDays(-1)
        Frm = Now.Date.AddDays(-1 * CInt(Me.CmbLstNdays.Text))
        strDKdate = " (DOI between '" & Format(Frm, "dd-MMM-yy") & "' and '" & Format(Thru, "dd-MMM-yy") & " 23:59')"
    End Sub

End Class