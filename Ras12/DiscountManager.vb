Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.MySharedFunctionsWzConn
Public Class DiscountManager
    Private CharEntered As Boolean = False
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private strSQL As String
    Private Sub DiscountManager_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        cmd.Dispose()
        If Conn.State = ConnectionState.Open Then Conn.Close()
        Me.Dispose()
    End Sub
    Private Sub DiscountManager_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If HasNewerVersion_R12(Application.ProductVersion) Or SysDateIsWrong(Conn) Then
            Me.Close()
            Me.Dispose()
            End
        End If
        Me.BackColor = pubVarBackColor
        Me.CmbNoOfVCR.Text = "01"
        'LoadGridVCR("")
        Me.CmbCurr.Text = "VND"
        Me.CmbType.Text = "SELL"
        cboFilter.SelectedIndex = 0
    End Sub

    Private Sub LblUpdateVCRAMT_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateVCRAMT.LinkClicked
        Dim i As Int16 = 1, VCRNo As String, StartNo As Int32, LstVCRofThisCode As String
        Dim VCRPrefix As String = Me.txtPrjCode.Text
        Dim MyAns As Int16 = MsgBox("Plz Make Sure VcrCode and NoOfVCR Are Correct. Need Another Thought?", MsgBoxStyle.Information Or MsgBoxStyle.YesNo, msgTitle)
        If MyAns = vbYes Then Exit Sub

        If Me.TxtAmt.Text = 0 Or Me.CmbType.Text = "" Or Me.CmbCurr.Text = "" Or Me.txtPrjCode.Text = "' then" Then
            MsgBox("Some Inputs Cant Be Empty", MsgBoxStyle.Information, msgTitle)
            Exit Sub
        ElseIf Me.txtPrjCode.Text.Length <> 3 And Me.txtPrjCode.Text.Length <> 4 And Me.txtPrjCode.Text <> "TVCR100-" Then
            MsgBox("Invalid Project Code", MsgBoxStyle.Information, msgTitle)
            Exit Sub
        ElseIf Me.txtPrjCode.Text.Length = 4 AndAlso InStr("SFD", Me.txtPrjCode.Text.Substring(0, 1)) = 0 Then
            MsgBox("4-Charactor Code Should Begin With S or F or D", MsgBoxStyle.Information, msgTitle)
            Exit Sub
        ElseIf cboTerritory.Text = "" Then
            MsgBox("Invalid Territory")
            Exit Sub
        ElseIf cboSVC.Text = "" Then
            MsgBox("Invalid SVC")
            Exit Sub
        ElseIf Not IsNumeric(txtMAX_OB.Text) Or txtMAX_OB.Text > 1 Then
            MsgBox("Invalid MAX_OB")
            Exit Sub
        ElseIf Not IsNumeric(txtKickBack.Text) Or txtKickBack.Text > 100 Or txtKickBack.Text < 0 Then
            MsgBox("Invalid Kickback")
            Exit Sub
        End If


        If CmbNoValue.Text <> "" Then VCRPrefix = VCRPrefix & Me.CmbNoValue.Text
        cmd.CommandText = "select top 1 VCRNo from DiscountVCR where VCRno like '" & VCRPrefix _
            & "%' and Isnumeric(substring(VCRNo," & VCRPrefix.Length + 1 & ",13))=1 order by VCRNo Desc"
        LstVCRofThisCode = cmd.ExecuteScalar
        If String.IsNullOrEmpty(LstVCRofThisCode) Then
            StartNo = 0
        Else
            StartNo = LstVCRofThisCode.Substring(VCRPrefix.Length)
        End If
        For i = 1 To CInt(Me.CmbNoOfVCR.Text)
            VCRNo = VCRPrefix & Format(i + StartNo, "00000")
            cmd.CommandText = "insert DiscountVCR (Amount, Curr, ValidThru, Condi, VCRNo, Status" _
                & ",Territory,SVC,MAX_OB,KickBack, FstUser, City) values (" & _
                CDec(Me.TxtAmt.Text) & ",'" & Me.CmbCurr.Text & "','" & Me.TxtValidThru.Value.Date & "','" & Me.TxtCondi.Text & "','" & VCRNo & _
                "','" & IIf(Me.CmbType.Text = "FOC", "OK", "QQ") & "','" _
                & cboTerritory.Text & "','" & cboSVC.Text & "'," & txtMAX_OB.Text _
                & "," & txtKickBack.Text & ",'" & myStaff.SICode & "','" & myStaff.City & "')"
            cmd.ExecuteNonQuery()
        Next
        LoadGridVCR("")
    End Sub
    Private Sub TxtAmt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAmt.KeyDown, TxtBlockAmt.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub
    Private Sub TxtAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAmt.KeyPress, TxtBlockAmt.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub GridVCR_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridVCR.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblDelete.Visible = (GridVCR.CurrentRow.Cells("Status").Value = "OK")
        Me.LblBlock.Visible = (GridVCR.CurrentRow.Cells("Status").Value = "OK")
        Me.TxtBlockAmt.Text = Me.GridVCR.CurrentRow.Cells("Amount").Value
    End Sub
    Private Sub LoadGridVCR(pDK As String)
        strSQL = "select RecID,Status, ValidThru, Amount, VCRno, Curr, Condi from DiscountVCR "
        Select Case cboFilter.Text
            Case "OK only"
                strSQL = strSQL & " where status ='OK' and validthru>='" & Format(Now, "dd-MMM-yy") & "'"
            Case "XX only"
                strSQL = strSQL & " where status ='XX'"
            Case "BK only"
                strSQL = strSQL & " where status ='BK'"
            Case "Expired"
                strSQL = strSQL & " where status ='OK' and validthru<'" & Format(Now, "dd-MMM-yy") & "'"
        End Select
        strSQL = strSQL & pDK
        Me.GridVCR.DataSource = GetDataTable(strSQL)
        Me.GridVCR.Columns("RecID").Width = 50
        Me.GridVCR.Columns("ValidThru").Width = 60
        Me.GridVCR.Columns("Curr").Width = 40
        Me.GridVCR.Columns("Condi").Width = 350

        Me.GridVCR.Columns("RecID").Visible = False
        Me.GridVCR.Columns("Status").Visible = False
        Me.LblDelete.Visible = False
        Me.LblBlock.Visible = False
        Me.LblUpdateVCRAMT.Visible = False
    End Sub
    Private Sub TxtCondi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCondi.TextChanged
        Me.LblUpdateVCRAMT.Visible = True
    End Sub
    Private Sub ChkXX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadGridVCR("")
    End Sub

    Private Sub LblDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        If Me.TxtRMK.Text = "" Then Exit Sub
        cmd.CommandText = "Update DiscountVCR set Status='XX', RMK='" & Me.TxtRMK.Text.Replace("--", "") & _
            "', LstUser='" & myStaff.SICode & "', LstUpdate=getdate() where RecID=" & Me.GridVCR.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridVCR("")
    End Sub

    Private Sub LblBlock_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblBlock.LinkClicked
        If Me.TxtRMK.Text = "" Then Exit Sub
        Dim AvailAmt As Decimal
        For i As Int16 = 0 To Me.GridVCR.RowCount - 1
            If Me.GridVCR.Item("S", i).Value Then
                If CDec(Me.TxtBlockAmt.Text) > Me.GridVCR.Item("Amount", i).Value Then
                    MsgBox("Current Value is Smaller that Value You Wanna Block.", MsgBoxStyle.Critical, msgTitle)
                Else
                    AvailAmt = Me.GridVCR.Item("Amount", i).Value - CDec(Me.TxtBlockAmt.Text)
                    cmd.CommandText = "Update DiscountVCR set Status='BK', RMK='" & Me.TxtRMK.Text.Replace("--", "") & "|" & _
                        Me.txtTenkhach.Text.Replace("--", "") & "', LstUser='" & myStaff.SICode & "', LstUpdate='" & Now.Date & "'"
                    If AvailAmt > 0 Then
                        cmd.CommandText = cmd.CommandText & ", Amount=" & CDec(Me.TxtBlockAmt.Text) & _
                            " where RecID=" & Me.GridVCR.Item("RecID", i).Value & _
                            "; insert DiscountVCR (Amount, Curr, ValidThru, Condi, VCRNo, Status, FstUser, city) " & _
                            " select " & AvailAmt & ", Curr, ValidThru, Condi, VCRNo, 'OK','" & myStaff.SICode & _
                            "'','" & myStaff.City & "' from DiscountVCR "
                    End If
                    cmd.CommandText = cmd.CommandText & " where RecID=" & Me.GridVCR.Item("RecID", i).Value
                    cmd.ExecuteNonQuery()
                End If
            End If
        Next
        LoadGridVCR("")
    End Sub

    Private Sub LbLFilter_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LbLFilter.LinkClicked
        LoadGridVCR(" and VCRNO like '" & Me.txtPrjCode.Text & "%' ")
    End Sub

    Private Sub cboFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFilter.SelectedIndexChanged
        LoadGridVCR("")
    End Sub
End Class