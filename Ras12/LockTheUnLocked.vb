Public Class LockTheUnLocked
    Private WhatAction As String
    Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Sub LoadGridALRPT()
        Try
            cmd.CommandText = "Drop table #LockUnLock"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        cmd.CommandText = "select * into #LockUnLock from Actionlog where tableName='RPTLISTAL'"
        cmd.ExecuteNonQuery()
        GridALRPT.DataSource = GetDataTable("Select top 16 recID, F1 as RPTNo, actionBy as ClosedBy, ActionDate as CloseOn " & _
                                            " from #LockUnLock where F1 not in " & _
                                            "(select F1 from #LockUnLock where DoWhat='U') and F2='OK' order by RecID desc")
        Me.LblUnlockALRPT.Visible = False
        Me.GridALRPT.Columns(0).Visible = False
        Me.GridALRPT.Columns(2).Width = 64
        Me.GridALRPT.Columns(3).Width = 128
    End Sub
    Private Sub LoadGridPending(ByVal pAction As String)
        WhatAction = pAction
        If pAction = "UNLOCK" Then
            Me.GridPendingTRX.DataSource = GetDataTable("select RecID, AL, RCPNO, Currency, TTLDue, RPTNO, CustShortName from rcp where rcpno='" & _
                                                    Me.TxtTRXNo.Text & "' and status='OK' and RPTNO<>'' and counter='" & MySession.Counter & "'")
        Else

            Try
                cmd.CommandText = "Drop table #LockUnLock"
                cmd.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            cmd.CommandText = "select * into #LockUnLock from Actionlog where tableName='RCP' and DoWhat='UnlockTRX'"
            cmd.ExecuteNonQuery()

            Me.GridPendingTRX.DataSource = GetDataTable("select F3,tableName, F1 as RCPID, F2 as TRXNO, ActionBy as UnLockedBy," & _
                                                        "ActionDate as UnLockedOn from #LockUnLock " & _
                                                        " where f1 in (select recID from RCP where RPTNo='') and F4='" & MySession.Counter & "'")
        End If
        Me.GridPendingTRX.Columns(0).Visible = False
        Me.GridPendingTRX.Columns(1).Visible = False
        Me.GridPendingTRX.Columns(2).Width = 64
        Me.GridPendingTRX.Columns(3).Width = 90
        Me.GridPendingTRX.Columns(4).Width = 80
        Me.GridPendingTRX.Columns(5).Width = 128
        Me.LblLock.Enabled = False
        Me.LblUnLockThis.Enabled = False
    End Sub
    Private Sub LockTheUnLocked_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        LoadGridPending("LOCK")
        If MySession.Domain = "GSA" Then LoadGridALRPT()
    End Sub

    Private Sub LblLock_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblLock.LinkClicked
        Me.LblLock.Enabled = False
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update RCP set rptNo='" & Me.GridPendingTRX.CurrentRow.Cells("F3").Value & "' where recid=" & _
            Me.GridPendingTRX.CurrentRow.Cells("RCPID").Value & _
            ";" & UpdateLogFile("RCP", "ReLockTRX", Me.GridPendingTRX.CurrentRow.Cells("RCPID").Value, _
            Me.GridPendingTRX.CurrentRow.Cells("TRXNO").Value, Me.GridPendingTRX.CurrentRow.Cells("F3").Value, "", "", "", "", "")
        cmd.ExecuteNonQuery()
        LoadGridPending("LOCK")
    End Sub

    Private Sub GridPendingTRX_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridPendingTRX.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If WhatAction = "LOCK" Then
            Me.LblLock.Enabled = True
        Else
            Me.LblUnLockThis.Enabled = True
        End If
    End Sub

    Private Sub TxtTRXNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTRXNo.LostFocus
        LoadGridPending("UNLOCK")
    End Sub

    Private Sub LblUnLockThis_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUnLockThis.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update RCP set rptNo='' where recid=" & Me.GridPendingTRX.CurrentRow.Cells("RECID").Value & _
            ";" & UpdateLogFile("RCP", "UnlockTRX", Me.GridPendingTRX.CurrentRow.Cells("RECID").Value, Me.TxtTRXNo.Text, Me.GridPendingTRX.CurrentRow.Cells("RCPNO").Value, MySession.Counter, "", "", "", "")
        cmd.ExecuteNonQuery()
        Me.TxtTRXNo.Text = ""
    End Sub

    Private Sub GridALRPT_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridALRPT.CellContentClick
        Me.LblUnlockALRPT.Visible = True
    End Sub

    Private Sub LblUnlockALRPT_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUnlockALRPT.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update actionlog set F2='XX', F3='" & myStaff.SICode & "', F4='Unlock@" & Now & "' where recid=" & _
            Me.GridALRPT.CurrentRow.Cells("RecID").Value & _
            "; Update TKT set RPTNO='' where rptno='" & Me.GridALRPT.CurrentRow.Cells("RPTNO").Value & "'"
        cmd.ExecuteNonQuery()
        LoadGridALRPT()
    End Sub
End Class