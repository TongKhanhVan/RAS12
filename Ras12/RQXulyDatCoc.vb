Public Class RQXulyDatCoc
    Private varWho As String
    Private strSQL As String
    Private fList As String
    Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal parWho As String)
        InitializeComponent()
        varWho = parWho
    End Sub
    Private Sub RQXulyDatCoc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        Me.LblProcess.Enabled = False
        If varWho = "SALES" Then
            Me.CmbRQType.Items.Add("C-Change")
            Me.CmbRQType.Items.Add("R-Refund")
            Me.CmbRQType.Items.Add("P-Penalty")
        ElseIf varWho = "FO" Then
            Me.CmbRQType.Items.Add("I-reInstate")
            Me.CmbRQType.Items.Add("D-Delete")
        ElseIf varWho = "BO" Then
            Me.LblProcess.Enabled = True

        End If
        If Me.CmbRQType.Items.Count > 0 Then
            Me.CmbRQType.Text = Me.CmbRQType.Items(0).ToString
        Else
            Me.CmbRQType.Enabled = False
        End If
        Me.LblFind.Enabled = Me.CmbRQType.Enabled
        Me.LblSendRQ.Enabled = Me.CmbRQType.Enabled
        LoadPendingRQ()
        fList = "RCPID, TKNO, FTKT, SRV, Qty, DOI, DOF, BkgClass, PaxName, DocType, PaxType, TourCode, PromoCode, FareType," & _
            "Fare, Charge, ChargeTV, NetToAL, RPTNo, RPTstatus, Itinerary, FareBasis, RCPNO, FstUser, FstUpdate, Currency, AL, StatusAL,"
    End Sub
    Private Sub LoadPendingRQ()
        Try
            cmd.CommandText = "Drop Table #DEPO"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        cmd.CommandText = "select * into #DEPO from ActionLog where tablename='RQDEPHANDLE'"
        cmd.ExecuteNonQuery()

        strSQL = "Select RecID, doWhat as Status, f2 as DocNo, F3 as RCPNO, F4 as DOI, F6 as Amount, F10 as Fee" & _
            ", F9 as NewTL, F7 as CustShortName, F8 as RQBy, f1 as TKTID, f5 as Itinerary " & _
            " from #DEPO where tablename='RQDEPHANDLE' "
        If Me.chckOKonly.Checked Then
            strSQL = strSQL & " and left(doWhat,2)='OK' "
        Else
            strSQL = strSQL & " and left(doWhat,2)='QQ' "
        End If
        If Me.CmbRQType.Items.Count > 0 Then
            strSQL = strSQL & " and substring(doWhat,4,1)='" & Me.CmbRQType.Text.Substring(0, 1) & "'"
        End If
        strSQL = strSQL & " order by recID desc"
        Me.GridPending.DataSource = GetDataTable(strSQL)
        Me.GridPending.Columns(0).Visible = False
        Me.GridPending.Columns("Status").Width = 50
        Me.GridPending.Columns("DOI").Width = 64
        Me.LblProcess.Visible = False
        Me.LblRefund.Visible = False
    End Sub
    Private Sub LblFind_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblFind.LinkClicked
        strSQL = "Select t.RecID, TKNO as DocNo, t.RCPNO, DOI, Itinerary, t.Currency as Curr, Fare as Amount, CustShortName, "
        strSQL = strSQL & " t.Status from TKT t inner join RCP r on t.rcpid=r.recid"
        strSQL = strSQL & " where left(tkno,3)='GRP' and replace(TKNO,' ','')='" & Replace(Me.TxtDocNo.Text.Trim, " ", "") & "' "
        Me.GridPreview.DataSource = GetDataTable(strSQL)
        Me.GridPreview.Columns(0).Visible = False
        Me.GridPreview.Columns("RecID").Visible = False
        Me.GridPreview.Columns("DocNo").Width = 75
        Me.GridPreview.Columns("Curr").Width = 40
        Me.GridPreview.Columns("DOI").Width = 50
        Me.GridPreview.Columns("DOi").DefaultCellStyle.Format = "dd-MMM-yy"
        Me.GridPreview.Columns("Amount").Width = 75
        Me.GridPreview.Columns("Amount").DefaultCellStyle.Format = "#,##0"
        Me.GridPreview.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.LblSendRQ.Enabled = False
        If Me.GridPreview.RowCount = 0 Then
            Me.TxtNote.Enabled = False
        Else
            Me.TxtNote.Enabled = True
        End If

    End Sub

    Private Sub LblSendRQ_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSendRQ.LinkClicked
        Dim DaRQ As Integer, rqType As String = Me.CmbRQType.Text.Substring(0, 1)
        Dim RecNo As Integer = Me.GridPreview.Item("RecID", 0).Value

        If Me.TxtNote.Text = "" Then
            MsgBox("Plz Specify the Reason", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        DaRQ = ScalarToInt("ActionLog", "RecID", "left(doWhat,2)='QQ' and tableName='RQDEPHANDLE' and F1='" & Me.GridPreview.Item("RecID", 0).Value & "'")

        If DaRQ <> 0 Then
            MsgBox("RQ for This Deposit Handler Has Been Previously Sent", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        cmd.CommandText = UpdateLogFile("RQDEPHANDLE", "QQ-" & Me.CmbRQType.Text.Substring(0, 1), Me.GridPreview.Item("RecID", 0).Value, _
            Me.GridPreview.Item("DocNo", 0).Value, Me.GridPreview.Item("RCPNO", 0).Value, Me.GridPreview.Item("DOI", 0).Value, _
            Me.GridPreview.Item("Itinerary", 0).Value, Me.GridPreview.Item("Curr", 0).Value & Me.GridPreview.Item("Amount", 0).Value.ToString, _
             Me.GridPreview.Item("CustShortName", 0).Value, myStaff.SICode & Now.ToString, Format(Me.TxtNewTL.Value, "dd-MMM-yy"), Me.TxtFee.Text)
        cmd.ExecuteNonQuery()

        MsgBox("RQST Has Been Sucessfully Sent", MsgBoxStyle.Information, msgTitle)
        Me.TxtDocNo.Text = ""
        Me.TxtNote.Text = ""
        Me.LblSendRQ.Enabled = False
        LoadPendingRQ()
    End Sub

    Private Sub TxtNote_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNote.LostFocus
        If Me.TxtNote.Text.Length > 0 Then Me.LblSendRQ.Enabled = True
    End Sub

    Private Sub CmbRQType_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbRQType.LostFocus
        If Me.CmbRQType.Text.Substring(0, 1) = "C" Then
            Me.LblNewFltDate.Visible = True
        Else
            Me.LblNewFltDate.Visible = False
        End If
        Me.TxtNewTL.Visible = Me.LblNewFltDate.Visible
    End Sub
    Private Sub LblChange_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblProcess.LinkClicked
        Dim Fee As Decimal = CDec(Me.GridPending.CurrentRow.Cells("fee").Value)
        Dim newTL As Date = CDate(Me.GridPending.CurrentRow.Cells("newTL").Value)
        Dim TKTID As Integer = CInt(Me.GridPending.CurrentRow.Cells("TKTID").Value)
        Dim RQType As String = Me.GridPending.CurrentRow.Cells("Status").Value.ToString.Substring(3, 1)
        If RQType = "D" Then ' Delete
            cmd.CommandText = " update TKT set Status='XX' where recid=" & TKTID
        ElseIf RQType = "I" Then ' reInstate
            cmd.CommandText = " update TKT set Status='OK' where recid=" & TKTID
        ElseIf InStr("CR", RQType) > 0 Then
            cmd.CommandText = "Insert TKT (" & fList & "Status, LstUser) select " & fList & "'XX','" & myStaff.SICode & _
                "' from TKT where recid=" & TKTID
            cmd.CommandText = cmd.CommandText & "; update TKT set Fare=Fare-" & Math.Abs(Fee) & ", DOI='" & newTL & "' where recid=" & TKTID
        ElseIf RQType = "P" Then ' Penalty
            cmd.CommandText = " update TKT set Fare=0 where recid=" & TKTID
        End If
        cmd.ExecuteNonQuery()
        RQ_Done("OK-" & RQType)
    End Sub
    Private Sub RQ_Done(ByVal pAction As String)
        cmd.CommandText = "update Actionlog set doWhat='" & pAction & "', F11='" & myStaff.SICode & "', F12='" & Format(Now, "dd-MMM-yy") & _
            "' where recid=" & Me.GridPending.CurrentRow.Cells(0).Value
        cmd.ExecuteNonQuery()
        MsgBox("Selected RQ Has Been Proceeded", MsgBoxStyle.Information, msgTitle)
        LoadPendingRQ()
    End Sub
    Private Sub GridPreview_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridPreview.CellContentClick
        Me.TxtNewTL.Text = Format(Me.GridPreview.CurrentRow.Cells("DOI").Value, "dd-MMM-yy")
    End Sub

    Private Sub chckOKonly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chckOKonly.CheckedChanged
        LoadPendingRQ()
    End Sub

    Private Sub GridPending_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridPending.CellContentClick
        If Me.GridPending.CurrentRow.Cells("Status").Value.ToString.Contains("-R") Then
            Me.LblRefund.Visible = True
        End If

        If Me.GridPending.CurrentRow.Cells("Status").Value.ToString.Contains("-I") Or _
            Me.GridPending.CurrentRow.Cells("Status").Value.ToString.Contains("-D") Then
            Me.LblProcess.Visible = True
        End If
        lbkReject.Visible = True
    End Sub

    Private Sub LblRefund_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblRefund.LinkClicked
        cmd.CommandText = "update Actionlog set dowhat='OK-R', actiondate=getdate(), actionby='" & myStaff.SICode & "' where recid=" & _
            Me.GridPending.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadPendingRQ()
    End Sub

    Private Sub lbkReject_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkReject.LinkClicked
        RQ_Done("XX-X")
        LoadPendingRQ()
    End Sub
End Class