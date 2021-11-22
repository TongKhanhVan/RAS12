Imports SharedFunctions.MySharedFunctions
Public Class ServiceFee
    Private CharEntered As Boolean = False
    Private ActivePage As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Sub ServiceFee_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    Private Sub ServiceFee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        LoadCmb_MSC(Me.CmbChannel, "Channel")
        LoadCmb_MSC(Me.CmbLevel, "Level")
        LoadCmb_MSC(Me.CmbDiscBase, "DiscBase")
        LoadCmb_MSC(Me.CmbFType, "FareType")
        ActivePage = "SF"
        Me.CmbLevel.Items.Add("--")
        Me.CmbChannel.Text = Me.CmbChannel.Items(0).ToString
        Me.CmbSBU.Text = Me.CmbSBU.Items(0).ToString
        Me.CmbBase.Text = Me.CmbBase.Items(0).ToString
        Me.CmbTRXtype.Text = Me.CmbTRXtype.Items(0).ToString
        Me.CmbSFType.Text = Me.CmbSFType.Items(0).ToString
        Me.CmbFType.Text = Me.CmbFType.Items(0).ToString
        Me.CmbLevel.Text = "--"
        Me.CmbArea.Text = Me.CmbArea.Items(0).ToString
        Me.CmbISI.Text = Me.CmbISI.Items(0).ToString
        CheckRightForALLForm(Me)
    End Sub

    Private Sub CmbBiz_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbSBU.LostFocus
        GenALList()
        If Me.CmbSBU.Text <> "TVS" Then
            Me.CmbDiscBase.Text = Me.CmbDiscBase.Items(1).ToString
            Me.CmbDiscBase.Enabled = False
        Else
            Me.CmbDiscBase.Enabled = True
        End If
    End Sub
    Private Sub GenALList()

        If Me.CmbSBU.Text = "GSA" Then
            LoadCmb_MSC(Me.CmbAL, myStaff.GSA)
        ElseIf Me.CmbSBU.Text = "TVS" Then
            LoadCmb_MSC(Me.CmbAL, myStaff.ALList)
        ElseIf Me.CmbSBU.Text = "EDU" Then
            LoadCmb_MSC(Me.CmbAL, "select 'XX' as VAL")
        End If
    End Sub

    Private Sub TxtfareFrm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        TxtfareFrm.KeyDown, TxtFareThru.KeyDown, TxtVAL.KeyDown, TxtMin.KeyDown, TxtMax.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub

    Private Sub TxtfareFrm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles _
        TxtfareFrm.KeyPress, TxtFareThru.KeyPress, TxtVAL.KeyPress, TxtMin.KeyPress, TxtMax.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub
    Private Function TrungSF() As Boolean
        Dim StrDK As String, RecID As Integer
        Dim KQ As Boolean = False
        StrDK = " SBU+AL+Channel+CustLevel+TRXType+G_FIT+FID_fier+ISI+Area='"
        StrDK = StrDK & Me.CmbSBU.Text & Me.CmbAL.Text & Me.CmbChannel.Text & Me.CmbLevel.Text
        StrDK = StrDK & Me.CmbTRXtype.Text & Me.CmbFType.Text & Me.CmbFSource.Text
        StrDK = StrDK & Me.CmbISI.Text & Me.CmbArea.Text & "' and status<>'XX' and ("
        StrDK = StrDK & " ('" & Me.TxtValidFrom.Value & "' between ValidFrom and ValidThru) or "
        StrDK = StrDK & " ('" & Me.TxtValidThru.Value & "' between ValidFrom and ValidThru)) and "
        StrDK = StrDK & " ((" & CDec(Me.TxtfareFrm.Text) & " between FareFrom and FareThrough) or"
        StrDK = StrDK & " (" & CDec(Me.TxtFareThru.Text) & " between FareFrom and FareThrough) )"
        RecID = ScalarToInt("servicefee", "RecID", StrDK)
        If RecID > 0 Then KQ = True
        Return KQ
    End Function
    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateSF.LinkClicked
        Dim MyAns As Int16
        MyAns = MsgBox("Plz Check All Inputs Before Continue. Go?", MsgBoxStyle.Information Or MsgBoxStyle.YesNo, msgTitle)
        If MyAns = vbNo Then Exit Sub
        If Me.CmbAL.Text = "" Or CDec(Me.TxtFareThru.Text) = 0 Or Me.CmbLevel.Text = "--" Or _
            (Me.CmbSFType.Text = "PCT" And (CDec(Me.TxtMin.Text) = 0 Or CDec(Me.TxtMax.Text) = 0)) Or _
            CDec(Me.TxtFareThru.Text) < CDec(Me.TxtfareFrm.Text) Or CDec(Me.TxtMax.Text) < CDec(Me.TxtMin.Text) Or _
            Me.TxtValidFrom.Value > Me.TxtValidThru.Value Or Me.TxtValidThru.Value < Now.Date Then
            MsgBox("Invalid Input. ", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        ElseIf CDec(Me.TxtVAL.Text) = 0 Then
            MyAns = MsgBox("Zero Service Fee. Plz Check Your Input. Continue?", MsgBoxStyle.Information Or MsgBoxStyle.YesNo, msgTitle)
            If MyAns = vbNo Then Exit Sub
        End If
        If TrungSF() Then
            MsgBox("Overlap Found. Plz Check", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        cmd.CommandText = "insert ServiceFee (SBU, AL, ValidFrom, ValidThru, Channel, CustLevel, TRXType, Base, FareFrom, " & _
            "fareThrough, VAL, SFType, MinVal, MaxVal, Refundable, Area, INF, G_FIT, FID_fier, ISI, DestCountries, FstUser) values ('" & _
            Me.CmbSBU.Text & "','" & Me.CmbAL.Text & "','" & Me.TxtValidFrom.Value.Date & "','" & Me.TxtValidThru.Value.Date & _
            " 23:59','" & Me.CmbChannel.Text & "','" & Me.CmbLevel.Text & "','" & Me.CmbTRXtype.Text & "','" & Me.CmbBase.Text & "','" & _
            CDec(Me.TxtfareFrm.Text) & "','" & CDec(Me.TxtFareThru.Text) & "','" & CDec(Me.TxtVAL.Text) & "','" & Me.CmbSFType.Text & _
            "','" & CDec(Me.TxtMin.Text) & "','" & CDec(Me.TxtMax.Text) & "','" & IIf(Me.ChkNonRefundable.Checked, 0, 1) & "','" & _
            Me.CmbArea.Text & "','" & IIf(Me.ChkINF.Checked, 1, 0) & "','" & Me.CmbFType.Text & "','" & Me.CmbFSource.Text & "','" & _
            Me.CmbISI.Text & "','" & IIf(Me.TxtDest.Text.Substring(0, 2) = "--", "", Me.TxtDest.Text) & "','" & myStaff.SICode & "')"
        cmd.ExecuteNonQuery()
        LoadGridSF("")
    End Sub
    Private Sub LoadGridSF(ByVal pFilter As String)
        Dim DKstatus As String = "", strSQL As String
        Me.LckLblApproveSF.Visible = False
        Me.LblDelete.Visible = False
        strSQL = "Select * from ServiceFee where sbu='" & Me.CmbSBU.Text & "' and al='" & Me.CmbAL.Text & "' "
        If Me.ChkActiveOnly.Checked Then
            DKstatus = " and status <>'XX' and validthru>getdate()"
        End If
        If Me.ChkQOnly.Checked Then
            DKstatus = " and status ='QQ'"
        End If
        strSQL = strSQL & DKstatus
        Me.GridSF.DataSource = GetDataTable(strSQL)
        Me.GridSF.Columns(0).Visible = False
        Me.GridSF.Columns("AL").Width = 32
        Me.GridSF.Columns("SBU").Width = 32
        Me.GridSF.Columns("Base").Width = 32
        Me.GridSF.Columns("TRXType").Width = 32
        Me.GridSF.Columns("Channel").Width = 32
        Me.GridSF.Columns("CustLevel").Width = 32
        Me.GridSF.Columns("Area").Width = 32
        Me.GridSF.Columns("INF").Width = 32
        Me.GridSF.Columns("MinVal").Width = 32
        Me.GridSF.Columns("MaxVal").Width = 32
        Me.GridSF.Columns("SFType").Width = 32
        Me.GridSF.Columns("ISI").Width = 32
        Me.GridSF.Columns("G_FIT").Width = 32
        Me.GridSF.Columns("FID_fier").Width = 32
        Me.GridSF.Columns("VAL").Width = 32
        Me.GridSF.Columns("FareFrom").Width = 56
        Me.GridSF.Columns("FareThrough").Width = 56
        Me.GridSF.Columns("ValidFrom").Width = 75
        Me.GridSF.Columns("ValidThru").Width = 75
        Me.GridSF.Columns("Status").Width = 32
        Me.GridSF.Columns("FstUser").Width = 32
        Me.GridSF.Columns("LstUser").Width = 32
        Me.GridSF.Columns("ApproveBy").Width = 32
        Me.GridSF.Columns("MinVal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("MaxVal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("VAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("FareFrom").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("FareThrough").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        For i As Int16 = 0 To Me.GridSF.Rows.Count - 1
            ChangeGridCurrentColor(Me.GridSF, i)
        Next

        strSQL = "Select * from Cust_Discount where SBU+AL='" & Me.CmbSBU.Text + Me.CmbAL.Text & "'"
        strSQL = strSQL & DKstatus
        Me.GridDisc.DataSource = GetDataTable(strSQL)
        Me.GridDisc.Columns(0).Visible = False
        Me.GridDisc.Columns("AL").Width = 32
        Me.GridDisc.Columns("SBU").Width = 32
        Me.GridDisc.Columns("Channel").Width = 32
        Me.GridDisc.Columns("CustLevel").Width = 32
        Me.GridDisc.Columns("fareSource").Width = 32
        Me.GridDisc.Columns("Base").Width = 32
        Me.GridDisc.Columns("DType").Width = 32
        Me.GridDisc.Columns("VAL").Width = 32
        Me.GridDisc.Columns("ValidFrom").Width = 75
        Me.GridDisc.Columns("ValidThru").Width = 75
        Me.GridDisc.Columns("Status").Width = 32
        Me.GridDisc.Columns("FstUser").Width = 32
        Me.GridDisc.Columns("LstUser").Width = 32
        Me.GridDisc.Columns("ApproveBy").Width = 32
        Me.GridDisc.Columns("VAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        For i As Int16 = 0 To Me.GridDisc.Rows.Count - 1
            ChangeGridCurrentColor(Me.GridDisc, i)
        Next
    End Sub
    Private Sub GridSF_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridSF.CellContentClick
        If e.RowIndex < 0 Then
            Me.LckLblApproveSF.Visible = False
            Me.LblDelete.Visible = False
            Me.LblChangeValidity.Visible = False
            Exit Sub
        End If
        If Me.GridSF.CurrentRow.Cells("Status").Value = "QQ" Then Me.LckLblApproveSF.Visible = True
        If Me.GridSF.CurrentRow.Cells("Status").Value <> "XX" Then
            Me.LblChangeValidity.Visible = True
            Me.LblDelete.Visible = True
        End If
    End Sub

    Private Sub LblApprove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblApproveSF.LinkClicked
        ChangeStatus("OK")
        Me.LckLblApproveSF.Visible = False
    End Sub
    Private Sub ChangeGridCurrentColor(ByVal myGrid As DataGridView, ByVal varR As Integer)
        Dim varColor As Color
        If myGrid.Item("Status", varR).Value = "QQ" Then
            varColor = Color.Gray
        ElseIf myGrid.Item("Status", varR).Value = "XX" Then
            varColor = Color.Red
        Else
            varColor = Color.Black
        End If
        For c As Int16 = 0 To myGrid.Columns.Count - 1
            myGrid.Item(c, varR).Style.ForeColor = varColor
        Next
    End Sub
    Private Sub LblDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        ChangeStatus("XX")
        Me.LblDelete.Visible = False
    End Sub
    Private Sub ChangeStatus(ByVal pStatus As String)
        Dim tblName As String = IIf(ActivePage = "SF", "ServiceFee", "Cust_Discount")
        Dim currStatus As String, RecID As Integer, strSQL As String
        If ActivePage = "SF" Then
            RecID = Me.GridSF.CurrentRow.Cells("recID").Value
            currStatus = Me.GridSF.CurrentRow.Cells("Status").Value
        Else
            RecID = Me.GridDisc.CurrentRow.Cells("recID").Value
            currStatus = Me.GridDisc.CurrentRow.Cells("status").Value
        End If
        If pStatus = "XX" And currStatus = "QQ" Then
            strSQL = "delete from " & tblName
        Else
            strSQL = "update " & tblName & " set status='" & pStatus & "'"
            If pStatus = "OK" Then
                strSQL = strSQL & ", ApproveBy='" & mystaff.SICOde & "', ApproveOn=getdate()"
            Else
                strSQL = strSQL & ", LstUser='" & myStaff.SICode & "', LstUpdate=getdate()"
            End If
        End If
        strSQL = strSQL & " where recid=" & RecID
        cmd.CommandText = strSQL
        cmd.ExecuteNonQuery()

        If pStatus = "XX" And currStatus = "QQ" Then
            LoadGridSF("")
        Else
            If ActivePage = "SF" Then
                Me.GridSF.CurrentRow.Cells("Status").Value = pStatus
                ChangeGridCurrentColor(Me.GridSF, Me.GridSF.CurrentRow.Index)
            Else
                Me.GridDisc.CurrentRow.Cells("Status").Value = pStatus
                ChangeGridCurrentColor(Me.GridDisc, Me.GridDisc.CurrentRow.Index)
            End If
        End If
    End Sub
    Private Sub CmbSFType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbSFType.SelectedIndexChanged
        If Me.CmbSFType.Text = "PCT" Then
            Me.LblPCT.Text = "%"
            Me.CmbBase.Text = "TKT"
            Me.TxtMin.Enabled = True
        Else
            Me.TxtMin.Enabled = False
            Me.LblPCT.Text = "$"
        End If
        Me.TxtMax.Enabled = Me.TxtMin.Enabled
        'Me.CmbBase.Enabled = Not Me.TxtMin.Enabled
    End Sub
    Private Sub CmbTRXtype_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbTRXtype.LostFocus
        If Me.CmbTRXtype.Text = "S" Then
            Me.CmbSFType.Enabled = True
        Else
            Me.CmbSFType.Enabled = False
            Me.CmbSFType.Text = "VAL"
            Me.CmbBase.Text = "TKT"
            Me.TxtfareFrm.Text = 0
            Me.TxtFareThru.Text = 1
            Me.ChkNonRefundable.Checked = True
        End If
        'Me.CmbBase.Enabled = Me.CmbSFType.Enabled
        Me.ChkNonRefundable.Enabled = Me.CmbSFType.Enabled
        Me.TxtfareFrm.Enabled = Me.CmbSFType.Enabled
        Me.TxtFareThru.Enabled = Me.CmbSFType.Enabled
    End Sub


    Private Sub ChkQOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkQOnly.Click
        LoadGridSF("")
    End Sub

    Private Sub ChkXXOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkActiveOnly.Click
        LoadGridSF("")
    End Sub

    Private Sub TabDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabDiscount.Enter
        ActivePage = "DS"
        Me.LckLblApproveSF.Visible = False
        Me.LblDelete.Visible = False
        Me.LblFSource.Text = "Fare Source"
    End Sub
    Private Sub TabSF_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabSF.Enter
        ActivePage = "SF"
        Me.LckLblApproveSF.Visible = False
        Me.LblDelete.Visible = False
        Me.LblFSource.Text = "Fare Identifier"
    End Sub

    Private Sub LblUpdateDisc_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateDisc.LinkClicked

        If Me.CmbLevel.Text = "--" Or _
            (InStr("COM_COP", Me.CmbDiscBase.Text) = 0 And Me.CmbSFType.Text = "PCT" And CDec(Me.TxtVAL.Text) > 9) Then
            MsgBox("Invalid Input. Plz Check", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        cmd.CommandText = "insert Cust_Discount (SBU, al, COC, channel, CustLevel,  DType, VAL, ValidFrom, ValidThru, Base, " & _
            "FareSource, FstUser) values ('" & Me.CmbSBU.Text & "','" & Me.CmbAL.Text & "','" & Me.TxtCOC.Text & "','" & _
            Me.CmbChannel.Text & "','" & Me.CmbLevel.Text & "','" & Me.CmbSFType.Text & "','" & Me.TxtVAL.Text & "','" & _
            Me.TxtValidFrom.Value.Date & "','" & Me.TxtValidThru.Value.Date & " 23:59','" & Me.CmbDiscBase.Text & "','" & _
            Me.CmbFSource.Text & "','" & myStaff.SICode & "')"
        cmd.ExecuteNonQuery()
        LoadGridSF("")
    End Sub
    Private Sub GridDisc_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridDisc.CellContentClick
        If e.RowIndex < 0 Then
            Me.LckLblApproveSF.Visible = False
            Me.LblDelete.Visible = False
            Me.LblChangeValidity.Visible = False
            Exit Sub
        End If
        If Me.GridDisc.CurrentRow.Cells("Status").Value = "QQ" Then Me.LckLblApproveSF.Visible = True
        If Me.GridDisc.CurrentRow.Cells("Status").Value <> "XX" Then
            Me.LblChangeValidity.Visible = True
            Me.LblDelete.Visible = True
        End If

    End Sub
    Private Sub CmbChannel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbChannel.SelectedIndexChanged
        Dim tmpTbl As String = IIf(ActivePage = "SF", "ServiceFee", "Cust_Discount")
        Dim tmpSQL As String, strSQL As String
        Dim strDK As String = Me.CmbSBU.Text + Me.CmbAL.Text + Me.CmbChannel.Text
        strSQL = "select distinct CustLevel as VAL from ServiceFee where status <>'XX' and SBU+AL+Channel='"
        strSQL = strSQL & strDK & "' UNION "
        strSQL = strSQL & "select distinct CustLevel as VAL from " & tmpTbl & " where status <>'XX' and SBU+AL+Channel='"
        strSQL = strSQL & strDK & "' "

        tmpSQL = strSQL
        strSQL = strSQL & " UNION "
        strSQL = strSQL & "select  VAL from MISC where CAT='Level' and val not in ("
        strSQL = strSQL & tmpSQL & ")"
        LoadCmb_MSC(Me.CmbLevel, strSQL)
        If Me.CmbChannel.Text = "WK" Then
            Me.CmbLevel.Items.Add("WKI")
            Me.CmbLevel.Text = "WKI"
            Me.CmbLevel.Enabled = False
        Else
            Me.CmbLevel.Enabled = True
        End If
    End Sub
    Private Sub LblShowAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblShowAll.LinkClicked
        LoadGridSF("")
    End Sub

    Private Sub LblFilter_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblFilter.LinkClicked
        MsgBox("Chua xong")
    End Sub

    Private Sub LblChangeValidity_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblChangeValidity.LinkClicked
        Dim tblName As String = IIf(ActivePage = "SF", "ServiceFee", "Cust_Discount")
        Dim RecID As Integer, CurrStatus As String
        If ActivePage = "SF" Then
            RecID = Me.GridSF.CurrentRow.Cells("recID").Value
            CurrStatus = Me.GridSF.CurrentRow.Cells("Status").Value
        Else
            RecID = Me.GridDisc.CurrentRow.Cells("recID").Value
            CurrStatus = Me.GridDisc.CurrentRow.Cells("status").Value
        End If
        cmd.CommandText = "update " & tblName & " set status='QQ', validThru='" & Me.TxtValidThru.Value.Date & _
            " 23:59' where recid=" & RecID
        cmd.ExecuteNonQuery()
        If CurrStatus = "OK" Then
            If ActivePage = "SF" Then
                cmd.CommandText = UpdateLogFile(tblName, "CHG VALDTY", "RecID:" & RecID.ToString, _
                    "OldValThru:" & Me.GridSF.CurrentRow.Cells("ValidThru").Value.ToString, _
                    "NewValThru:" & Me.TxtValidThru.Value.Date.ToString, _
                    "OldApprBy:" & Me.GridSF.CurrentRow.Cells("ApproveBy").Value, _
                    "OldApprOn:" & Me.GridSF.CurrentRow.Cells("ApproveOn").Value.ToString, _
                    "", "", "")
            Else
                cmd.CommandText = UpdateLogFile(tblName, "CHG VALDTY", "RecID:" & RecID.ToString, _
                    "OldValThru:" & Me.GridDisc.CurrentRow.Cells("ValidThru").Value.ToString, _
                    "NewValThru:" & Me.TxtValidThru.Value.Date.ToString, _
                    "OldApprBy:" & Me.GridDisc.CurrentRow.Cells("ApproveBy").Value, _
                    "OldApprOn:" & Me.GridDisc.CurrentRow.Cells("ApproveOn").Value.ToString, _
                    "", "", "")
            End If
            cmd.ExecuteNonQuery()
        End If
        LoadGridSF("")
    End Sub

    Private Sub CmbArea_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbArea.SelectedIndexChanged
        Me.TxtDest.Text = "--COMMA SEPARATED COUNTRY CODE--"
        Me.TxtDest.ForeColor = Color.DarkGray
        If Me.CmbArea.Text = "DOM" Then
            Me.TxtDest.Enabled = False
        Else
            Me.TxtDest.Enabled = True
        End If
    End Sub

    Private Sub TxtDest_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDest.Enter
        If Me.TxtDest.Text.Substring(0, 2) = "--" Then
            Me.TxtDest.Text = ""
            Me.TxtDest.ForeColor = Color.Black
        End If
    End Sub
    Private Sub CmbAL_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbAL.SelectedIndexChanged
        If ActivePage = "SF" Then
            LoadCmbFareIdentifier()
        End If
    End Sub
    Private Sub LoadCmbFareIdentifier()
        LoadCmb_MSC(Me.CmbFSource, "select VAL from misc where cat in ('SFFareSource','SFFareSource" & Me.CmbAL.Text & "')")
        Me.CmbFSource.Text = Me.CmbFSource.Items(0).ToString
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If ActivePage = "SF" Then
            LoadCmbFareIdentifier()
        Else
            LoadCmb_MSC(Me.CmbFSource, "FareSource")
        End If
    End Sub

    Private Sub LblViewXpireDate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblViewXpireDate.LinkClicked
        Me.GridExpireDate.Visible = True
        Dim tblName As String = IIf(ActivePage = "SF", "ServiceFee", "Cust_Discount")
        Dim frmExtend As New frmMassExtendServiceFee(CmbSBU.Text, tblName)
        frmExtend.ShowDialog()
        'LoadGridExprireDate()
    End Sub
    Private Sub LoadGridExprireDate()

        Dim strSQL As String
        Dim tblName As String = IIf(ActivePage = "SF", "ServiceFee", "Cust_Discount")

        strSQL = "Select distinct AL, Channel, CustLevel, ValidThru from " & tblName & " where sbu='" & Me.CmbSBU.Text & "'"
        strSQL = strSQL & " and status='OK' and ValidThru between getdate() and '" & Now.AddMonths(3) & "'"
        Me.GridExpireDate.DataSource = GetDataTable(strSQL)
        Me.GridExpireDate.Left = 597
        Me.GridExpireDate.Height = 256
        Me.GridExpireDate.Columns(0).Width = 32
        Me.GridExpireDate.Columns(1).Width = 32
        Me.GridExpireDate.Columns(2).Width = 32
    End Sub
    Private Sub GridExpireDate_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridExpireDate.CellContentClick
        Me.GridExpireDate.Visible = False
        If e.RowIndex < 0 Then Exit Sub
        Me.CmbAL.Text = Me.GridExpireDate.CurrentRow.Cells(0).Value
        Me.CmbChannel.Text = Me.GridExpireDate.CurrentRow.Cells(1).Value
        Me.CmbLevel.Text = Me.GridExpireDate.CurrentRow.Cells(2).Value
        Me.TxtValidFrom.Text = Me.GridExpireDate.CurrentRow.Cells(3).Value
    End Sub

End Class
