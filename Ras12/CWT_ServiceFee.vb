Public Class CWT_ServiceFee
    Private RegList As String
    Private MyCust As New objCustomer
    Private mblnFirstLoadCompleted As Boolean
    Private Sub CWT_ServiceFee_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub CWT_ServiceFee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim defaultRegCountry As String
        Dim CList As String
        MyCust.GenCustList()
        Me.BackColor = pubVarBackColor
        LoadCmb_VAL(Me.CmbCust, MyCust.List_CS & "  Union " & MyCust.List_LC)
        Me.CmbService.Text = Me.CmbService.Items(0).ToString
        Me.CmbTRX.Text = Me.CmbTRX.Items(0).ToString
        Me.CmbCurr.Text = Me.CmbCurr.Items(0).ToString
        CList = ScalarToString("MISC", "VAL", " Cat='REGCountry'")
        defaultRegCountry = ScalarToString("MISC", "VAL", " Cat='DefaultREGCountry'")
        For i As Int16 = 0 To UBound(CList.Split(" "))
            If InStr(defaultRegCountry, CList.Split(" ")(i)) Then
                Me.LstReg.Items.Add(CList.Split(" ")(i), True)
            Else
                Me.LstReg.Items.Add(CList.Split(" ")(i))
            End If
        Next
        'loadGridSF
        mblnFirstLoadCompleted = True
    End Sub

    Private Sub CmbCust_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCust.SelectedIndexChanged
        Me.LblUpdate.Visible = True
        loadGridSF()
    End Sub

    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        Dim aa As Decimal, cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim DKValid As String = " and ((ValidFrom between '" & Format(Me.TxtValidFrom.Value, "dd-MMM-yy") & "' and '" & Format(Me.TxtValidThru.Value, "dd-MMM-yy") & " 23:59') "
        DKValid = DKValid & " or ((ValidThru between '" & Format(Me.TxtValidFrom.Value, "dd-MMM-yy") & "' and '" & Format(Me.TxtValidThru.Value, "dd-MMM-yy") & " 23:59')))"
        Dim HS As Decimal = 1
        If Me.CmbAL.Text = "" Or Me.CmbRtg.Text = "" Or Me.cmbSegTKT.Text = "" Or Me.CmbCabin.Text = "" Then Exit Sub
        RegList = ""
        If Me.CmbRtg.Text.Substring(0, 3) = "REG" Then
            RegList = CountryOfRegional()
            If RegList = "" Then
                MsgBox("No Country Has Been Selected. Plz check your input", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If

        End If
        If ScalarToInt("CWT_SF", "RecID", "Custid=" & Me.CmbCust.SelectedValue & " and status='OK' and SVCType='" & Me.CmbService.Text & _
            "' and trx='" & Me.CmbTRX.Text & "' and RtgType='" & Me.CmbRtg.Text & "' and ALType='" & Me.CmbAL.Text & _
            "' and tkt ='" & Me.CmbTKT.Text & "' and SegTKT='" & Me.cmbSegTKT.Text.Substring(0, 1) & "' and OWRT='" & Me.CmbOWRT.Text & _
            "' and cabin='" & Me.CmbCabin.Text & _
            "' and FltTime='" & Me.TxtFltTime.Text & "' and WzAIR=" & IIf(Me.chkWzAIR.Checked, 1, 0) & DKValid) > 0 Then
            MsgBox("Duplicate Record found. Plz check your input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        Try
            If Not Me.ChkVAT.Checked Then HS = 1.1
            aa = CDec(Me.TxtAmt.Text)
            aa = CDec(Me.TxtPersonal.Text)
            aa = CDec(Me.TxtFltTime.Text)
            If (Me.CmbCurr.Text = "VND" And CDec(Me.TxtAmt.Text) <> 0 And CDec(Me.TxtAmt.Text) < 16000) Or _
                (Me.CmbCurr.Text = "USD" And CDec(Me.TxtAmt.Text) > 256) Then
                MsgBox("Illogic Currency/Amount. Plz check your input", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
            cmd.CommandText = "insert Cwt_SF (CustID, TRX, SVCType, Curr, RtgType, ALType, Amount, Countries, wzAIR, Base, ValidFrom, ValidThru, " & _
                "FstUser, TKT, Cabin, FltTime, SegTKT, AmtPersonal, OWRT) values (@CustID, @TRX, @SVCType, @Curr, @RtgType, @ALType, @Amount, @Countries, @wzAIR, @Base, @ValidFrom," & _
                "@ValidThru, @FstUser,@TKT, @Cabin, @FltTime,@SegTKT, @AmtPersonal,@OWRT)"
            cmd.Parameters.Clear()
            cmd.Parameters.Add("@CustID", SqlDbType.Int).Value = Me.CmbCust.SelectedValue
            cmd.Parameters.Add("@TRX", SqlDbType.VarChar).Value = Me.CmbTRX.Text
            cmd.Parameters.Add("@TKT", SqlDbType.VarChar).Value = Me.CmbTKT.Text
            cmd.Parameters.Add("@SVCType", SqlDbType.VarChar).Value = Me.CmbService.Text
            cmd.Parameters.Add("@Cabin", SqlDbType.VarChar).Value = Me.CmbCabin.Text
            cmd.Parameters.Add("@SegTKT", SqlDbType.VarChar).Value = Me.cmbSegTKT.Text.Substring(0, 1)
            cmd.Parameters.Add("@Curr", SqlDbType.VarChar).Value = Me.CmbCurr.Text
            cmd.Parameters.Add("@RtgType", SqlDbType.VarChar).Value = Me.CmbRtg.Text
            cmd.Parameters.Add("@ALType", SqlDbType.VarChar).Value = Me.CmbAL.Text
            cmd.Parameters.Add("@OWRT", SqlDbType.VarChar).Value = Me.CmbOWRT.Text
            cmd.Parameters.Add("@Amount", SqlDbType.Decimal).Value = CDec(Me.TxtAmt.Text) * HS
            cmd.Parameters.Add("@AmtPersonal", SqlDbType.Decimal).Value = CDec(Me.TxtPersonal.Text) * HS
            cmd.Parameters.Add("@FltTime", SqlDbType.Decimal).Value = CDec(Me.TxtFltTime.Text)
            cmd.Parameters.Add("@Countries", SqlDbType.VarChar).Value = RegList
            cmd.Parameters.Add("@wzAIR", SqlDbType.Bit).Value = IIf(Me.chkWzAIR.Checked, -1, 0)
            cmd.Parameters.Add("@Base", SqlDbType.VarChar).Value = Me.CmbBase.Text
            cmd.Parameters.Add("@ValidFrom", SqlDbType.DateTime).Value = Me.TxtValidFrom.Value.Date
            cmd.Parameters.Add("@ValidThru", SqlDbType.DateTime).Value = CDate(Format(Me.TxtValidThru.Value, "dd-MMM-yy") & " 23:59")
            cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
            cmd.ExecuteNonQuery()
            Me.LblUpdate.Visible = False
        Catch ex As Exception

        End Try
        loadGridSF()
    End Sub
    Private Function CountryOfRegional() As String
        Dim KQ As String = ""
        For i As Int16 = 0 To Me.LstReg.Items.Count - 1
            If Me.LstReg.GetItemChecked(i) Then
                KQ = KQ & "_" & Me.LstReg.Items(i).ToString
            End If
        Next
        Return KQ.Substring(1)
    End Function

    Private Sub GridSF_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridSF.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblDelete.Visible = True
    End Sub

    Private Sub LblDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = ChangeStatus_ByID("CWT_SF", "XX", Me.GridSF.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        loadGridSF()
    End Sub

    Private Sub LblCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCheckAll.LinkClicked
        Dim vCheck As Boolean
        If Me.LblCheckAll.Text = "Check ALL" Then
            Me.LblCheckAll.Text = "UnCheck ALL"
            vCheck = True
        Else
            vCheck = False
            Me.LblCheckAll.Text = "Check ALL"
        End If
        For i As Int16 = 0 To Me.LstReg.Items.Count - 1
            Me.LstReg.SetItemChecked(i, vCheck)
        Next
    End Sub
    Private Sub loadGridSF()
        If Not mblnFirstLoadCompleted Then
            Exit Sub
        End If
        'Try
        Me.GridSF.DataSource = GetDataTable("Select Recid,CustID, SVCType, WzAIR, TRX, Status" _
                                            & ", ValidFrom, ValidThru, TKT, ALType, RtgType" _
                                            & ", Curr,Amount, Base, Cabin, FltTime, SegTKT, AmtPersonal, OWRT" _
                                            & ", Countries, FstUser, FstUpdate, LstUser, LstUpdate" _
                                            & " from cwt_SF where custid=" & Me.CmbCust.SelectedValue & " and status='OK'")
        'Me.GridSF.Columns(0).Visible = False
        'Me.GridSF.Columns(11).Visible = False
        'Me.GridSF.Columns(12).Visible = False
        Me.GridSF.Columns("CustId").Width = 40
        Me.GridSF.Columns("SvcType").Width = 50
        Me.GridSF.Columns("TRX").Width = 30
        Me.GridSF.Columns("Curr").Width = 40
        Me.GridSF.Columns("Status").Width = 40
        Me.GridSF.Columns("FstUser").Width = 50
        Me.GridSF.Columns("LstUser").Width = 50
        Me.GridSF.Columns("ValidFrom").Width = 65
        Me.GridSF.Columns("ValidThru").Width = 70
        Me.GridSF.Columns("TKT").Width = 30
        Me.GridSF.Columns("ALType").Width = 40
        Me.GridSF.Columns("RtgType").Width = 45
        Me.GridSF.Columns("Cabin").Width = 45
        Me.GridSF.Columns("FltTime").Width = 45
        Me.GridSF.Columns("WzAir").Width = 45
        Me.GridSF.Columns("Base").Width = 45
        Me.GridSF.Columns("SegTkt").Width = 50
        Me.GridSF.Columns("OWRT").Width = 40
        'Catch ex As Exception
        'End Try
    End Sub

    Private Sub CmbTRX_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbTRX.SelectedIndexChanged
        Me.LblUpdate.Visible = True
    End Sub

    Private Sub CmbService_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbService.SelectedIndexChanged
        Me.chkWzAIR.Checked = False
        If Me.CmbService.Text = "AIR" Then
            Me.chkWzAIR.Visible = False
        Else
            Me.chkWzAIR.Visible = True
        End If
    End Sub

    Private Sub CmbCurr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCurr.SelectedIndexChanged
        If Me.CmbCurr.Text = "PCT" Then
            Me.CmbBase.Visible = True
        Else
            Me.CmbBase.Visible = False
        End If
        Me.Label12.Visible = Me.CmbBase.Visible
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles TxtValidThru.ValueChanged
        If Me.TxtValidThru.Value < Me.TxtValidFrom.Value Then Me.TxtValidThru.Focus()
    End Sub

    Private Sub CmbRtg_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbRtg.SelectedIndexChanged
        If Me.CmbRtg.Text.Contains("REG") Then
            Me.LstReg.Enabled = True
        Else
            Me.LstReg.Enabled = False
        End If
    End Sub
End Class