Imports SharedFunctions.MySharedFunctions
Public Class GroupBooking
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private CharEntered As Boolean = False
    Private SelectedRQ As String = ""
    Private Sub Nummeric_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQtyADL.KeyDown, _
        txtQtyCHD.KeyDown, txtQtyINF.KeyDown, txtQtyFOC.KeyDown, txtTaxADL.KeyDown, txtTaxCHD.KeyDown, txtTaxINF.KeyDown, txtTaxFOC.KeyDown, _
        txtFareADL.KeyDown, txtFareCHD.KeyDown, txtFareINF.KeyDown, txtFareFOC.KeyDown, _
            txtRSPADL.KeyDown, txtRSPCHD.KeyDown, txtRSPINF.KeyDown, txtRSPFOC.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub
    Private Sub Nummeric_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQtyADL.KeyPress, _
        txtQtyCHD.KeyPress, txtQtyINF.KeyPress, txtQtyFOC.KeyPress, txtTaxADL.KeyPress, txtTaxCHD.KeyPress, txtTaxINF.KeyPress, txtTaxFOC.KeyPress, _
        txtFareADL.KeyPress, txtFareCHD.KeyPress, txtFareINF.KeyPress, txtFareFOC.KeyPress, _
            txtRSPADL.KeyPress, txtRSPCHD.KeyPress, txtRSPINF.KeyPress, txtRSPFOC.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPNR_Enter(sender As Object, e As EventArgs) Handles txtPNR.Enter
        If Me.txtPNR.Text = "PNR" Then Me.txtPNR.Text = ""
        Me.txtPNR.Top = 48
        Me.txtPNR.Height = 272
    End Sub

    Private Sub GroupBooking_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCmb_VAL(Me.CmbCust, "select RecID as VAL, CustShortName as DIS from customerlist where recID in " & _
                     "(select custID from cust_detail where status='OK' and Cat='Channel' and val in ('TA','TO'))" & _
                     " and recID>0 order by CustShortName")
        LoadCmb_MSC(Me.CmbAL, "select AL as VAL from airline where AL not in ('01','XX') order by AL")
        LoadCmb_MSC(Me.CmbAL_Mail, "select AL as VAL from airline where AL not in ('01','XX') order by AL")
        Me.CmbCurr.Text = "USD"
        Me.CmbStatus.Text = "OK"
        Me.CmbPackg.Text = "Series"
        LoadGridRQ("")
        ClearScreen()
    End Sub
    Private Sub ClearScreen()
        Me.CmbCurr.Text = "USD"
        Me.CmbStatus.Text = "OK"
        Me.txtEmail.Text = ""
        Me.txtPNR.Text = "PNR"
        Me.txtNote.Text = "Condition"
        Me.txtRBD.Text = ""
        Me.txtRTG.Text = ""
        Me.txtQtyADL.Text = "0"
        Me.txtQtyCHD.Text = "0"
        Me.txtQtyINF.Text = "0"
        Me.txtQtyFOC.Text = "0"
        Me.txtTaxADL.Text = "0"
        Me.txtTaxCHD.Text = "0"
        Me.txtTaxINF.Text = "0"
        Me.txtTaxFOC.Text = "0"
        Me.txtRSPADL.Text = "0"
        Me.txtRSPCHD.Text = "0"
        Me.txtRSPINF.Text = "0"
        Me.txtRSPFOC.Text = "0"
        Me.txtFareADL.Text = "0"
        Me.txtFareCHD.Text = "0"
        Me.txtFareINF.Text = "0"
        Me.txtFareFOC.Text = "0"
        Me.txtDepo1.Value = "01-Jan-2000"
        Me.txtDepo2.Value = "01-Jan-2000"
        Me.txtAdvName.Value = "01-Jan-2000"
        Me.txtTktg.Value = "01-Jan-2000"
        Me.txtDepo1.Checked = False
        Me.txtDepo2.Checked = False
        Me.txtAdvName.Checked = False
        Me.txtTktg.Checked = False

        Me.CmbPackg.Text = Me.CmbPackg.Items(0).ToString
    End Sub
    Private Sub LblClear_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblClear.LinkClicked
        ClearScreen()
    End Sub
    Private Function invalidInput() As Boolean
        If Me.txtDepo2.Checked And Me.txtDepo2.Value < Me.txtDepo1.Value Then Return True
        If Me.txtAdvName.Checked And Me.txtAdvName.Value < Me.txtDepo1.Value Then Return True
        If Me.txtTktg.Checked And Me.txtTktg.Value < Me.txtAdvName.Value Then Return True
        If Me.txtReturn.Value <= Me.txtDOF.Value Then Return True
        If Me.txtDOF.Value < Me.txtTktg.Value Then Return True
        If CInt(Me.txtQtyINF.Text) > CInt(Me.txtQtyCHD.Text) Or CInt(Me.txtQtyCHD.Text) > CInt(Me.txtQtyADL.Text) Then Return True
        If CInt(Me.txtFareINF.Text) > CInt(Me.txtFareCHD.Text) Or CInt(Me.txtFareCHD.Text) > CInt(Me.txtFareADL.Text) Then Return True
        If CInt(Me.txtTaxINF.Text) > CInt(Me.txtTaxCHD.Text) Or CInt(Me.txtTaxCHD.Text) > CInt(Me.txtTaxADL.Text) Then Return True
        If CInt(Me.txtFareINF.Text) > CInt(Me.txtRSPINF.Text) Or CInt(Me.txtFareCHD.Text) > CInt(Me.txtRSPCHD.Text) Then Return True
        If CInt(Me.txtFareFOC.Text) > CInt(Me.txtRSPFOC.Text) Or CInt(Me.txtFareADL.Text) > CInt(Me.txtRSPADL.Text) Then Return True
        'If Me.txtRLOC.Text.Trim.Length <> 6 Then Return True
        If Me.txtRTG.Text.Trim.Length > 83 Then Return True
        If Not CheckRTG(Me.txtRTG.Text.Trim) Then Return True
        Return False
    End Function
    Private Sub InsertNewPNR(pRQNo As String, pCustID As Integer)
        cmd.Parameters.Clear()
        cmd.CommandText = "insert ActionLog (TableName, doWhat, ActionBy, F1, F2, F3, F4, F5, F6, F7, F8,F9,F10, F11, F13) " & _
            "values ('GRPBKG','OK','" & myStaff.SICode & "', @F1, @F2, @F3, @F4, @F5, @F6, @F7, @F8,@F9,@F10,@F11,@F13) ;SELECT SCOPE_IDENTITY() AS [RecID]"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@F1", SqlDbType.VarChar).Value = pcustID
        cmd.Parameters.Add("@F2", SqlDbType.VarChar).Value = pRQNo
        cmd.Parameters.Add("@F3", SqlDbType.VarChar).Value = Me.txtRLOC.Text & "|" & Me.CmbAL.Text & "|" & _
            Format(Me.txtDOF.Value, "dd-MMM-yy") & "|" & Me.txtRTG.Text & "|" & Me.txtRBD.Text
        cmd.Parameters.Add("@F4", SqlDbType.VarChar).Value = Me.CmbStatus.Text & "|" & Me.txtNote.Text & "|" & Me.CmbCurr.Text & "|" & Me.txtRTG.Text
        cmd.Parameters.Add("@F5", SqlDbType.VarChar).Value = Me.txtQtyADL.Text & "|" & Me.txtQtyCHD.Text & "|" & Me.txtQtyINF.Text & "|" & Me.txtQtyFOC.Text
        cmd.Parameters.Add("@F6", SqlDbType.VarChar).Value = Me.txtFareADL.Text & "|" & Me.txtFareCHD.Text & "|" & Me.txtFareINF.Text & "|" & Me.txtFareFOC.Text
        cmd.Parameters.Add("@F7", SqlDbType.VarChar).Value = Me.txtTaxADL.Text & "|" & Me.txtTaxCHD.Text & "|" & Me.txtTaxINF.Text & "|" & Me.txtTaxFOC.Text
        cmd.Parameters.Add("@F8", SqlDbType.VarChar).Value = Me.txtRSPADL.Text & "|" & Me.txtRSPCHD.Text & "|" & Me.txtRSPINF.Text & "|" & Me.txtRSPFOC.Text
        cmd.Parameters.Add("@F9", SqlDbType.VarChar).Value = Me.txtPNR.Text
        cmd.Parameters.Add("@F10", SqlDbType.VarChar).Value = Format(Me.txtDepo1.Value, "dd-MMM-yy") & "|" & _
            Format(Me.txtDepo2.Value, "dd-MMM-yy") & "|" & _
            Format(Me.txtAdvName.Value, "dd-MMM-yy") & "|" & _
            Format(Me.txtTktg.Value, "dd-MMM-yy") & "|"
        cmd.Parameters.Add("@F11", SqlDbType.VarChar).Value = myStaff.Counter
        cmd.Parameters.Add("@F13", SqlDbType.VarChar).Value = Format(Me.txtReturn.Value, "dd-MMM-yy") & "|" & Me.CmbPackg.Text
        Dim recid As Integer = cmd.ExecuteScalar

        cmd.Parameters.Clear()
        cmd.Parameters.Add("@F2", SqlDbType.VarChar)
        cmd.Parameters.Add("@F3", SqlDbType.VarChar)
        cmd.Parameters.AddWithValue("@F1", recid)
        cmd.Parameters.AddWithValue("@F4", "??")
        cmd.Parameters.AddWithValue("@F5", myStaff.Counter)

        cmd.CommandText = "insert ActionLog (TableName, doWhat, actionBy, F1, F2, F3, F4, F5) " & _
            "values ('GRPBKGDL','OK','" & myStaff.SICode & "', @F1, @F2, @F3, @F4, @F5)"
        cmd.Parameters("@F2").Value = "ADVNAME"
        cmd.Parameters("@F3").Value = Format(Me.txtAdvName.Value, "dd-MMM-yy")
        cmd.ExecuteNonQuery()

        cmd.CommandText = "insert ActionLog (TableName, doWhat, actionBy, F1, F2, F3, F4, F5) " & _
            "values ('GRPBKGDL','OK','" & myStaff.SICode & "', @F1, @F2, @F3, @F4, @F5)"
        cmd.Parameters("@F2").Value = "TKTG"
        cmd.Parameters("@F3").Value = Format(Me.txtTktg.Value, "dd-MMM-yy")
        cmd.ExecuteNonQuery()

        cmd.CommandText = "insert ActionLog (TableName, doWhat, actionBy, F1, F2, F3, F4, F5) " & _
            "values ('GRPBKGDL','OK','" & myStaff.SICode & "', @F1, @F2, @F3, @F4, @F5)"
        cmd.Parameters("@F2").Value = "DEPO1"
        cmd.Parameters("@F3").Value = Format(Me.txtDepo1.Value, "dd-MMM-yy")
        cmd.ExecuteNonQuery()

        cmd.CommandText = "insert ActionLog (TableName, doWhat, actionBy, F1, F2, F3, F4, F5) " & _
            "values ('GRPBKGDL','OK','" & myStaff.SICode & "', @F1, @F2, @F3, @F4, @F5)"
        cmd.Parameters("@F2").Value = "DEPO2"
        cmd.Parameters("@F3").Value = Format(Me.txtDepo2.Value, "dd-MMM-yy")
        cmd.ExecuteNonQuery()
        LoadGridRQ("")

    End Sub
    Private Sub LblSave_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSave.LinkClicked
        Dim RQNo As String = SelectedRQ, CustID As Integer
        If invalidInput() Then
            MsgBox("Invalid Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        Dim RecID As Integer = ScalarToInt("ActionLog", "RecID", "tableName='GRPBKG' and doWhat='OK' and left(f3,6)='" & Me.txtRLOC.Text & "'")
        If RecID > 0 Then
            If MsgBox(Me.txtRLOC.Text & " Already Exists. Wanna Correct Input", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, msgTitle) = vbYes Then Exit Sub
            DeletePNR(RecID, "XX")
        End If
        If SelectedRQ <> "" Then
            If MsgBox("Wanna Add This RLOC to Selected Request?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle) = vbNo Then
                RQNo = Me.CmbCust.Text & Format(Now, "ddMMyyHHmm")
                CustID = Me.CmbCust.SelectedValue
            Else
                CustID = Me.GridRQ.CurrentRow.Cells(0).Value
            End If
        Else
            RQNo = Me.CmbCust.Text & Format(Now, "ddMMyyHHmm")
            CustID = Me.CmbCust.SelectedValue
        End If
        InsertNewPNR(RQNo, CustID)
    End Sub
    Private Sub LoadGridRQ(pRQNo As String)
        'Try
        '    cmd.CommandText = "Drop Table #GRPBKG"
        '    cmd.ExecuteNonQuery()
        'Catch ex As Exception
        'End Try
        'cmd.CommandText = "select * into #GRPBKG from ActionLog where tableName='GRPBKG' and doWhat='OK'"
        'cmd.ExecuteNonQuery()

        Dim strQry As String = "select Distinct F1 as CustID, F2 as RequestNo from GRPBKG where doWhat='OK' "
        If pRQNo <> "" Then
            strQry = strQry & " and F2='" & pRQNo & "'"
        ElseIf Me.ChkSelectCustOnly.Checked Then
            strQry = strQry & " and F1='" & Me.CmbCust.SelectedValue.ToString.Trim & "'"
        End If
        Me.GridRQ.DataSource = GetDataTable(strQry)
        Me.GridRQ.Columns(0).Width = 50
        Me.GridRQ.Columns(1).Width = 128
        Me.LblDeleteRQ.Visible = False
        Me.LblDone.Visible = False
        Me.GridRLOC.DataSource = Nothing
    End Sub

    Private Sub LoadGridRLOC(pRQ As String, pRLOC As String)
        Dim strDK As String = "F2='" & pRQ & "'"
        If pRLOC <> "" Then strDK = "left(f3,6)='" & pRLOC & "'"
        Me.GridRLOC.DataSource = GetDataTable("select RecID, F1, F2,F3, F4, F5, F6, F7, F8,F9, F10, '' as RLOC" _
                                              & ", '' as DOF, 0 as TTLPax, F13 " & _
                                              " from actionLog where tableName='GRPBKG' and doWhat='OK' and " & strDK)
        For c As Int16 = 1 To 10
            Me.GridRLOC.Columns(c).Visible = False
        Next
        For r As Int16 = 0 To Me.GridRLOC.RowCount - 1
            Me.GridRLOC.Item("RLOC", r).Value = Me.GridRLOC.Item("F3", r).Value.ToString.Split("|")(0)
            Me.GridRLOC.Item("DOF", r).Value = Me.GridRLOC.Item("F3", r).Value.ToString.Split("|")(2)
            Me.GridRLOC.Item("TTLPax", r).Value = CInt(Me.GridRLOC.Item("F5", r).Value.ToString.Split("|")(0)) + _
                CInt(Me.GridRLOC.Item("F5", r).Value.ToString.Split("|")(1)) + _
                CInt(Me.GridRLOC.Item("F5", r).Value.ToString.Split("|")(2)) + _
                CInt(Me.GridRLOC.Item("F5", r).Value.ToString.Split("|")(3))
        Next

        Me.LblEmail.Visible = False
        Me.lblEmailAL.Visible = False
        Me.LblCFM.Visible = False
        Me.LblDeletePNR.Visible = False
        Me.LblAssign.Visible = False
        Me.LblUC.Visible = False
        Me.GridRLOC.Columns("RecID").Width = 64
        Me.GridRLOC.Columns("RLOC").Width = 64
        Me.GridRLOC.Columns("TTLPAx").Width = 56
        Me.GridRLOC.Columns("DOF").Width = 64
        Me.GridRLOC.Columns("F13").Visible = False
    End Sub
    Private Sub GridRQ_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridRQ.CellContentClick
        Me.LblDeleteRQ.Visible = True
        Me.LblDone.Visible = True
        SelectedRQ = Me.GridRQ.CurrentRow.Cells(1).Value
        LoadGridRLOC(SelectedRQ, "")
    End Sub

    Private Sub LblDeleteRQ_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDeleteRQ.LinkClicked

        cmd.CommandText = "update ActionLog set doWhat='XX', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & _
            "' where tableName='GRPBKG' and doWhat='OK' and F2='" & SelectedRQ & "'" & _
                "; update ActionLog set dowhat='XX', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & _
                "' where tableName='GRPBKGDL' and F1 in (select RecID from actionlog where " & _
                " tableName='GRPBKG' and F2='" & SelectedRQ & "')"
        cmd.ExecuteNonQuery()
        LoadGridRQ("")
    End Sub

    Private Sub LblDone_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDone.LinkClicked
        cmd.CommandText = "update ActionLog set doWhat='RR', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & _
            "' where tableName='GRPBKG' and doWhat='OK' and F2='" & SelectedRQ & "'" & _
            "; update ActionLog set F4='OK', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & _
                "' where tableName='GRPBKGDL' and F1 in (select RecID from actionlog where " & _
                " tableName='GRPBKG' and F2='" & SelectedRQ & "')"
        cmd.ExecuteNonQuery()
        LoadGridRQ("")
    End Sub
    Private Sub DeletePNR(pRecID As Integer, pStatus As String)
        cmd.CommandText = "update ActionLog set doWhat='" & pStatus & "', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & _
            "' where RecID=" & pRecID & _
            "; update ActionLog set dowhat='XX', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & _
            "' where tableName='GRPBKGDL' and F1='" & pRecID & "'"
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub LblDeletePNR_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDeletePNR.LinkClicked
        DeletePNR(Me.GridRLOC.CurrentRow.Cells("RecID").Value, "XX")
        LoadGridRLOC(SelectedRQ, "")
    End Sub

    Private Sub LblCFM_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCFM.LinkClicked
        If Me.GridRLOC.CurrentRow.Cells("f4").Value.ToString.Substring(0, 2) = "OK" Then Exit Sub
        cmd.CommandText = "update ActionLog set F4='OK" & Me.GridRLOC.CurrentRow.Cells("f4").Value.ToString.Substring(2) & _
            "', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & "'" & _
            " where RecID=" & Me.GridRLOC.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridRLOC(SelectedRQ, "")
    End Sub

    Private Sub GridDL_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridDL.CellContentClick
        Me.LblDone_DL.Visible = True
        Me.LblPostpone.Visible = True
    End Sub

    Private Sub GridRLOC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridRLOC.CellContentClick
        Me.LblEmail.Visible = True
        Me.lblEmailAL.Visible = True
        Me.LblCFM.Visible = True
        Me.LblDeletePNR.Visible = True
        Me.LblAssign.Visible = True
        Me.LblUC.Visible = True

        Me.CmbCust.SelectedValue = Me.GridRLOC.CurrentRow.Cells("F1").Value
        Me.txtRLOC.Text = Me.GridRLOC.CurrentRow.Cells("F3").Value.ToString.Split("|")(0)
        Me.CmbAL.Text = Me.GridRLOC.CurrentRow.Cells("F3").Value.ToString.Split("|")(1)
        Me.txtDOF.Text = Me.GridRLOC.CurrentRow.Cells("F3").Value.ToString.Split("|")(2)
        Me.txtReturn.Text = Me.GridRLOC.CurrentRow.Cells("F13").Value.ToString.Split("|")(0)
        Me.txtRBD.Text = Me.GridRLOC.CurrentRow.Cells("F3").Value.ToString.Split("|")(3)
        Me.txtRBD.Text = Me.GridRLOC.CurrentRow.Cells("F3").Value.ToString.Split("|")(4)
        Me.CmbStatus.Text = Me.GridRLOC.CurrentRow.Cells("F4").Value.ToString.Split("|")(0)
        Me.txtNote.Text = Me.GridRLOC.CurrentRow.Cells("F4").Value.ToString.Split("|")(1)
        Me.CmbCurr.Text = Me.GridRLOC.CurrentRow.Cells("F4").Value.ToString.Split("|")(2)
        Me.txtRTG.Text = Me.GridRLOC.CurrentRow.Cells("F4").Value.ToString.Split("|")(3)
        Me.txtQtyADL.Text = Me.GridRLOC.CurrentRow.Cells("F5").Value.ToString.Split("|")(0)
        Me.txtQtyCHD.Text = Me.GridRLOC.CurrentRow.Cells("F5").Value.ToString.Split("|")(1)
        Me.txtQtyINF.Text = Me.GridRLOC.CurrentRow.Cells("F5").Value.ToString.Split("|")(2)
        Me.txtQtyFOC.Text = Me.GridRLOC.CurrentRow.Cells("F5").Value.ToString.Split("|")(3)

        Me.txtFareADL.Text = Me.GridRLOC.CurrentRow.Cells("F6").Value.ToString.Split("|")(0)
        Me.txtFareCHD.Text = Me.GridRLOC.CurrentRow.Cells("F6").Value.ToString.Split("|")(1)
        Me.txtFareINF.Text = Me.GridRLOC.CurrentRow.Cells("F6").Value.ToString.Split("|")(2)
        Me.txtFareFOC.Text = Me.GridRLOC.CurrentRow.Cells("F6").Value.ToString.Split("|")(3)

        Me.txtTaxADL.Text = Me.GridRLOC.CurrentRow.Cells("F7").Value.ToString.Split("|")(0)
        Me.txtTaxCHD.Text = Me.GridRLOC.CurrentRow.Cells("F7").Value.ToString.Split("|")(1)
        Me.txtTaxINF.Text = Me.GridRLOC.CurrentRow.Cells("F7").Value.ToString.Split("|")(2)
        Me.txtTaxFOC.Text = Me.GridRLOC.CurrentRow.Cells("F7").Value.ToString.Split("|")(3)


        Me.txtRSPADL.Text = Me.GridRLOC.CurrentRow.Cells("F8").Value.ToString.Split("|")(0)
        Me.txtRSPCHD.Text = Me.GridRLOC.CurrentRow.Cells("F8").Value.ToString.Split("|")(1)
        Me.txtRSPINF.Text = Me.GridRLOC.CurrentRow.Cells("F8").Value.ToString.Split("|")(2)
        Me.txtRSPFOC.Text = Me.GridRLOC.CurrentRow.Cells("F8").Value.ToString.Split("|")(3)
        Me.txtPNR.Text = Me.GridRLOC.CurrentRow.Cells("F9").Value

        Me.txtDepo1.Text = Me.GridRLOC.CurrentRow.Cells("F10").Value.ToString.Split("|")(0)
        Me.txtDepo2.Text = Me.GridRLOC.CurrentRow.Cells("F10").Value.ToString.Split("|")(1)
        Me.txtAdvName.Text = Me.GridRLOC.CurrentRow.Cells("F10").Value.ToString.Split("|")(2)
        Me.txtTktg.Text = Me.GridRLOC.CurrentRow.Cells("F10").Value.ToString.Split("|")(3)

    End Sub

    Private Sub chkAllDeadline_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllRQST.CheckedChanged
        LoadGridDeadline()
    End Sub
    Private Sub LoadGridDeadline()
        Dim NextTwoDays As String = Format(Now.AddDays(2), "dd-MMM-yy")
        Me.LblDone_DL.Visible = False
        Me.LblPostpone.Visible = False
        Try
            Dim strQry As String = "select d.RecID, d.F1 as RLOCID,'' as RequestNo, '' as RLOC" _
                                    & ",g.F3 as AL, d.F2 as Task, d.F3 as DeadLine, d.F5 as RMK" _
                                    & " from ActionLog d " _
                                    & " left join ActionLog g on d.f1=g.recid" _
                                    & " where d.tableName='GRPBKGDL' and d.doWhat='OK' and d.F4='??'" _
                                    & " and d.f3 <>'01-jan-00'  and d.F5 ='" & myStaff.SelectedDomain _
                                    & "' and  g.doWhat='OK'"

            If Not Me.chkAllRQST.Checked Then
                strQry = strQry & " and d.F1='" & Me.GridRLOC.CurrentRow.Cells("RecID").Value.ToString & "'"
            End If
            If Me.ChkPastAndComing.Checked Then
                strQry = strQry & " and cast(d.f3 as datetime)<'" & NextTwoDays & "'"
            End If
            strQry = strQry & " order by cast(d.f3 as datetime)"
            Me.GridDL.DataSource = GetDataTable(strQry)
            Me.GridDL.Columns("RecId").Visible = False
            Me.GridDL.Columns("RlocId").Visible = False
            Me.GridDL.Columns("Task").Width = 100
            Me.GridDL.Columns("DeadLine").Width = 56
            Me.GridDL.Columns("RMK").Width = 128
            Me.GridDL.Columns("RLOC").Width = 64
            Me.GridDL.Columns("AL").Width = 32
            For r As Int16 = 0 To Me.GridDL.RowCount - 1
                If Me.GridDL.Item("AL", r).Value <> "" Then
                    Me.GridDL.Item("AL", r).Value = Me.GridDL.Item("AL", r).Value.ToString.Split("|")(1)
                End If
                If Me.chkAllRQST.Checked Then
                    If Me.GridDL.Item("RLOCID", r).Value <> 0 Then
                        Me.GridDL.Item("RLOC", r).Value = ScalarToString("ActionLog", "f3", "RecID=" & Me.GridDL.Item("RLOCID", r).Value).ToString.Substring(0, 6)
                        Me.GridDL.Item("RequestNo", r).Value = ScalarToString("ActionLog", "f2", "RecID=" & Me.GridDL.Item("RLOCID", r).Value)
                    End If
                Else
                    Me.GridDL.Item("RLOC", r).Value = Me.GridRLOC.CurrentRow.Cells("RLOC").Value
                    Me.GridDL.Item("RequestNo", r).Value = Me.GridRQ.CurrentRow.Cells("RequestNo").Value
                End If
                If Me.GridDL.Item("Deadline", r).Value < Now Then
                    Me.GridDL.Rows(r).DefaultCellStyle.ForeColor = Color.Red
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub txtFareADL_LostFocus(sender As Object, e As EventArgs) Handles txtFareADL.LostFocus, txtFareCHD.LostFocus, _
        txtFareINF.LostFocus, txtFareFOC.LostFocus, txtTaxADL.LostFocus, txtTaxCHD.LostFocus, txtTaxINF.LostFocus, txtTaxFOC.LostFocus
        Dim txt As TextBox = CType(sender, TextBox)
        Dim aa As Decimal = CDec(txt.Text)
        If Me.CmbCurr.Text = "VND" Then
            txt.Text = Format(aa, "#,##0")
        Else
            txt.Text = Format(aa, "#,##0.00")
        End If
    End Sub

    Private Sub LblDone_DL_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDone_DL.LinkClicked
        cmd.CommandText = "update ActionLog set F4='OK', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & "'" & _
            " where RecID=" & Me.GridDL.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridDeadline()
    End Sub

    Private Sub LblChangeDL_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblPostpone.LinkClicked
        If Me.txtTask.Text = "" Then Exit Sub
        cmd.CommandText = "update ActionLog set F3='" & Format(Me.txtMyDeadline.Value, "dd-MMM-yy") & "',F5='" & Me.txtTask.Text & _
            "', F12='" & myStaff.SICode & "|" & Format(Now, "dd-MMM-yy HH:mm") & "'" & _
            " where RecID=" & Me.GridDL.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridDeadline()
    End Sub
    Private Sub LblAddDL_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAddDL.LinkClicked
        cmd.CommandText = "insert ActionLog (TableName, doWhat, actionBy, F1, F2, F3, F4, F5) " & _
            "values ('GRPBKGDL','OK','" & myStaff.SICode & "', @F1, @F2, @F3, @F4, @F5)"
        cmd.Parameters.Clear()
        If Me.chkAllRQST.Checked Then
            cmd.Parameters.Add("@F1", SqlDbType.VarChar).Value = 0
        Else
            cmd.Parameters.Add("@F1", SqlDbType.VarChar).Value = Me.GridRLOC.CurrentRow.Cells("recID").Value.ToString
        End If
        cmd.Parameters.Add("@F2", SqlDbType.VarChar).Value = "My Task"
        cmd.Parameters.Add("@F3", SqlDbType.VarChar).Value = Format(Me.txtMyDeadline.Value, "dd-MMM-yy")
        cmd.Parameters.Add("@F4", SqlDbType.VarChar).Value = "??"
        cmd.Parameters.Add("@F5", SqlDbType.VarChar).Value = Me.txtTask.Text
        cmd.ExecuteNonQuery()
        LoadGridDeadline()
    End Sub

    Private Sub txtNote_Enter(sender As Object, e As EventArgs) Handles txtNote.Enter
        If Me.txtNote.Text = "Condition" Then Me.txtNote.Text = ""
    End Sub

    
    Private Sub txtDepo1_Enter(sender As Object, e As EventArgs) Handles txtDepo1.Enter
        If Me.txtDepo1.Value = "01-Jan-2000" Then Me.txtDepo1.Value = Now.Date.AddDays(2)
    End Sub

    Private Sub txtDepo2_Enter(sender As Object, e As EventArgs) Handles txtDepo2.Enter
        If Me.txtDepo2.Value = "01-Jan-2000" Then Me.txtDepo2.Value = Now.Date.AddDays(4)
    End Sub

    Private Sub txtAdvName_Enter(sender As Object, e As EventArgs) Handles txtAdvName.Enter
        If Me.txtAdvName.Value = "01-Jan-2000" Then Me.txtAdvName.Value = Me.txtDOF.Value.AddDays(-8)
    End Sub

    Private Sub txtTktg_Enter(sender As Object, e As EventArgs) Handles txtTktg.Enter
        If Me.txtTktg.Value = "01-Jan-2000" Then Me.txtTktg.Value = Me.txtDOF.Value.AddDays(-2)
    End Sub

    Private Sub LblEmail_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblEmail.LinkClicked
        Dim txtSubj As String = Me.GridRQ.CurrentRow.Cells("RequestNo").Value & ": " & _
            Me.GridRLOC.CurrentRow.Cells("F3").Value.ToString.Split("|")(3) & _
            "/" & Me.GridRLOC.CurrentRow.Cells("DOF").Value
        Dim txtCC As String = "groupticketing.sgn@transviet.com ;khanh.nguyenngoc@transviet.com"
        Dim txtBody As String = "Dear " & Me.CmbCust.Text & ","
        txtBody = txtBody & "%0d" & "Please find below details for your group booking request" & "%0d" & "%0d" & _
            Me.GridRLOC.CurrentRow.Cells("F9").Value.ToString.Replace(vbLf, "%0d")
        Dim txtEmail As String = ""
        Dim dTbl As DataTable = GetDataTable("select distinct VAL from Cust_detail where cat='EMAIL' and status='OK' and CustID=" & _
                                             Me.CmbCust.SelectedValue)
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            txtEmail = txtEmail & ";" & dTbl.Rows(i)("Val")
        Next
        If txtEmail.Length > 2 Then txtEmail = txtEmail.Substring(1)
        Process.Start(String.Format("mailto:{0}?subject={1}&cc={2}&body={3}", txtEmail, txtSubj, txtCC, txtBody))
    End Sub


    Private Sub ChkSelectCustOnly_Click(sender As Object, e As EventArgs) Handles ChkSelectCustOnly.Click
        LoadGridRQ("")
    End Sub

    Private Sub txtEmailAL_Enter(sender As Object, e As EventArgs) Handles txtEmailAL.Enter
        Me.txtEmailAL.Text = ScalarToString("MISC", "VAL1", "CAT='ALMAILGRPBKG' and val='" & Me.CmbAL_Mail.Text & "'")
    End Sub
    Private Sub LblUpdateALMail_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdateALMail.LinkClicked
        If Not Me.txtEmailAL.Text.Contains("@") Then Exit Sub
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim CurrRec As Integer = ScalarToInt("MISC", "RecID", "CAT='ALMAILGRPBKG' and val='" & Me.CmbAL_Mail.Text & "'")
        cmd.Parameters.Add("@email", SqlDbType.VarChar).Value = Me.txtEmailAL.Text
        If CurrRec = 0 Then
            cmd.CommandText = "insert MISC (CAt, val, val1) values ('ALMAILGRPBKG','" & Me.CmbAL_Mail.Text & "',@email)"
        Else
            cmd.CommandText = "update MISC set val1=@email where recid=" & CurrRec
        End If
        cmd.ExecuteNonQuery()
        Me.txtEmailAL.Text = ""
    End Sub

    Private Sub lblEmailAL_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblEmailAL.LinkClicked
        Dim txtSubj As String = Me.GridRQ.CurrentRow.Cells("RequestNo").Value & ": " & _
            Me.GridRLOC.CurrentRow.Cells("F3").Value.ToString.Split("|")(3) & _
            "/" & Me.GridRLOC.CurrentRow.Cells("DOF").Value
        Dim txtCC As String = "groupticketing.sgn@transviet.com ;khanh.nguyenngoc@transviet.com"
        Dim txtBody As String = "Dear  ,"
        txtBody = txtBody & "%0d" & "Please " & "%0d" & "%0d" & _
            Me.GridRLOC.CurrentRow.Cells("F9").Value.ToString.Replace(vbLf, "%0d")
        Dim txtEmail As String = ScalarToString("MISC", "VAL1", "CAT='ALMAILGRPBKG' and val='" & Me.CmbAL.Text & "'")
        Process.Start(String.Format("mailto:{0}?subject={1}&cc={2}&body={3}", txtEmail, txtSubj, txtCC, txtBody))

    End Sub

    Private Sub GridDL_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridDL.CellContentDoubleClick
        SelectedRQ = Me.GridDL.CurrentRow.Cells("RequestNo").Value
        LoadGridRQ(SelectedRQ)
        LoadGridRLOC(SelectedRQ, "")
        Me.TabControl1.SelectTab("TabPage1")
    End Sub
    Private Sub ChkPastAndComing_Click(sender As Object, e As EventArgs) Handles ChkPastAndComing.Click
        LoadGridDeadline()
    End Sub

    Private Sub LblFind_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblFind.LinkClicked
        LoadGridRLOC("", Me.txtRloc2Find.Text)
    End Sub

    Private Sub LblAssign_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAssign.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        DeletePNR(Me.GridRLOC.CurrentRow.Cells("RecID").Value, "XX")
        InsertNewPNR(Me.CmbCust.Text & Format(Now, "ddMMyyHHmm"), Me.CmbCust.SelectedValue)
    End Sub

    Private Sub txtPNR_LostFocus(sender As Object, e As EventArgs) Handles txtPNR.LostFocus
        Me.txtPNR.Top = 154
        Me.txtPNR.Height = 144
    End Sub

    Private Sub LblUC_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUC.LinkClicked
        DeletePNR(Me.GridRLOC.CurrentRow.Cells("RecID").Value, "UC")
        LoadGridRLOC(SelectedRQ, "")
    End Sub
End Class