Public Class CustChannelLevel
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand

    Private Sub CustChannelLevel_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub CustChannelLevel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadCmb_MSC(Me.CmbChannel, "Channel")
        LoadCmb_MSC(Me.CmbLevel, "Level")
        Me.BackColor = pubVarBackColor
        Me.CmbSBU.Text = Me.CmbSBU.Items(0).ToString
        Me.CmbChannel.Text = Me.CmbChannel.Items(0).ToString
        Me.CmbLevel.Text = Me.CmbLevel.Items(0).ToString
        CheckRightForALLForm(Me)
    End Sub

    Private Sub CmbSBU_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbSBU.LostFocus
        Me.CmbAL.Items.Clear()
        If Me.CmbSBU.Text = "EDU" Then
            Me.CmbAL.Items.Add("XX")
        ElseIf Me.CmbSBU.Text = "GSA" Then
            LoadCmb_MSC(Me.CmbAL, myStaff.GSA)
        ElseIf Me.CmbSBU.Text = "TVS" Then
            LoadCmb_MSC(Me.CmbAL, myStaff.ALList)
        End If
        Me.CmbAL.Text = Me.CmbAL.Items(0).ToString
    End Sub
    Private Sub LoadListAvail()
        Dim strSQL As String
        strSQL = "select CustID, l.CustShortName from cust_detail c inner join Customerlist l on l.recid=c.custid"
        strSQL = strSQL & " where c.status='OK' and cat='Channel' and val='" & Me.CmbChannel.Text & "' and  "
        strSQL = strSQL & " l.recid in (select custID from cust_Detail where cat='AL' and status='OK' and "
        strSQL = strSQL & " (VAL ='YY' or VAL like '%" & Me.CmbAL.Text & "%' or VAL like '%TS%'))"
        strSQL = strSQL & " and custid not in "
        strSQL = strSQL & "(select custid from cust_Channel_level where status<>'XX' and "
        strSQL = strSQL & " ValidFrom ='" & Me.TxtValidFrm.Value.Date & "' and ValidThru='"
        strSQL = strSQL & Me.TxtValidThru.Value.Date & " 23:59' and  "
        strSQL = strSQL & " (AL ='YY' or AL like '%" & Me.CmbAL.Text & "%' or AL like '%TS%')) "
        strSQL = strSQL & " order by CustShortName"
        Me.ChkListAvail.DataSource = GetDataTable(strSQL)
        Me.ChkListAvail.DisplayMember = "CustShortName"
        Me.ChkListAvail.ValueMember = "CustID"
        checkUnCheckALL(False)
    End Sub
    Private Sub checkUnCheckAll(ByVal pStatus As Boolean)

        For i As Int16 = 0 To Me.ChkListAvail.Items.Count - 1
            Me.ChkListAvail.SetItemChecked(i, pStatus)
        Next
    End Sub
    Private Sub LoadGridCurrent()
        Dim strSQL As String
        strSQL = "select c.recid, CustID, CustShortName, c.Status, ValidFrom, ValidThru from cust_Channel_level c "
        strSQL = strSQL & "inner join Customerlist l on l.recid=c.custid"
        strSQL = strSQL & " where SBU+al+Channel+CustLevel='"
        strSQL = strSQL & Me.CmbSBU.Text & Me.CmbAL.Text & Me.CmbChannel.Text & Me.CmbLevel.Text & "'"
        If Me.ChkQOnly.Checked Then
            strSQL = strSQL & " and c.status='QQ'"
        End If
        If Me.ChkActiveOnly.Checked Then
            strSQL = strSQL & " and c.status<>'XX' and '" & Now.Date & "' between c.validfrom and c.validthru "
        End If
        Me.GridCurrent.DataSource = GetDataTable(strSQL)
        Me.GridCurrent.Columns("CustID").Visible = False
        Me.GridCurrent.Columns("RecID").Visible = False
        Me.GridCurrent.Columns("Status").Width = 32
        Me.GridCurrent.Columns("CustShortName").Width = 128
        Me.GridCurrent.Columns("ValidFrom").Width = 75
        Me.GridCurrent.Columns("ValidThru").Width = 128
        For i As Int16 = 0 To Me.GridCurrent.Rows.Count - 1
            ChangeGridCurrentColor(i)
        Next
        Me.LckLblApprove.Visible = False
        Me.LblDeactive.Visible = False
        Me.LblChangeHL.Visible = False
    End Sub
    Private Sub ChangeGridCurrentColor(ByVal varR As Integer)
        Dim varColor As Color
        If Me.GridCurrent.Item("Status", varR).Value = "QQ" Then
            varColor = Color.Gray
        ElseIf Me.GridCurrent.Item("Status", varR).Value = "XX" Then
            varColor = Color.Red
        Else
            varColor = Color.Black
        End If
        For c As Int16 = 0 To Me.GridCurrent.Columns.Count - 1
            Me.GridCurrent.Item(c, varR).Style.ForeColor = varColor
        Next
    End Sub

    Private Sub CmbAL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        CmbAL.LostFocus, CmbChannel.LostFocus, CmbLevel.LostFocus
        Me.LblMemberOf.Text = "Members Of " & Me.CmbSBU.Text & "/" & Me.CmbAL.Text & "/" & Me.CmbChannel.Text & "/" & Me.CmbLevel.Text
    End Sub

    Private Sub LblUnAssigned_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUnAssigned.LinkClicked
        LoadListAvail()
    End Sub

    Private Sub LblMemberOf_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblMemberOf.LinkClicked
        LoadGridCurrent()
    End Sub
    Private Function CheckOverLap(ByVal pCustID As Integer, ByVal pCurrID As Integer) As String
        Dim KQ As String = "0", tmpSQL As String, LstValidThru As Date
        tmpSQL = " status <>'XX' and custID=" & pCustID & " and recID <>" & pCurrID & " and sbu='"
        tmpSQL = tmpSQL & Me.CmbSBU.Text & "' and al='" & Me.CmbAL.Text & "' order by recid desc"
        LstValidThru = ScalarToDate("cust_Channel_level", "top 1 ValidThru", tmpSQL)
        If LstValidThru.Date.AddDays(1) = Me.TxtValidFrm.Value.Date Then
            KQ = "0"
        ElseIf LstValidThru = #12:00:00 AM# Then
            KQ = "0"
        ElseIf LstValidThru.Date.AddDays(1) < Me.TxtValidFrm.Value.Date Then
            KQ = "1Current Policy for Custid=" & pCustID.ToString & " is Valid Thru " & Format(LstValidThru, "dd-MMM-yy")
        ElseIf LstValidThru.Date > Me.TxtValidFrm.Value.Date Then
            KQ = "2Current Policy for Custid=" & pCustID.ToString & " is Valid Thru " & Format(LstValidThru, "dd-MMM-yy")
        End If
        Return KQ
    End Function
    Private Sub LblToRight_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblToRight.LinkClicked
        Dim OverLap As String, MyAns As Int16, CustID As Integer, strSQL As String = ""
        For i As Int16 = 0 To Me.ChkListAvail.Items.Count - 1
            If Me.ChkListAvail.GetItemChecked(i) Then
                CustID = Me.ChkListAvail.Items(i).item("CustID")
                OverLap = CheckOverLap(CustID, 0)
                If OverLap.Substring(0, 1) = "1" Then
                    MyAns = MsgBox("Overlap: " & OverLap.Substring(1) & ". Wanna Continue?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, msgTitle)
                    If MyAns = vbNo Then GoTo ResumeHere
                ElseIf OverLap.Substring(0, 1) = "2" Then
                    MsgBox("Overlap: " & OverLap.Substring(1) & ". Check Your Input or Change The Current One Before Continue", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, msgTitle)
                    GoTo ResumeHere
                End If
                strSQL = strSQL & "; insert cust_Channel_Level (CustID, Channel, AL, CustLevel, ValidFrom, ValidThru, " & _
                    "SBU, FstUser) values ('" & CustID & "','" & Me.CmbChannel.Text & "','" & Me.CmbAL.Text & "','" & _
                    Me.CmbLevel.Text & "','" & Me.TxtValidFrm.Value.Date & "','" & Me.TxtValidThru.Value.Date & " 23:59','" & _
                    Me.CmbSBU.Text & "','" & myStaff.SICode & "')"
ResumeHere:
            End If
        Next
        Try
            cmd.CommandText = strSQL.Substring(1)
            cmd.ExecuteNonQuery()
            LoadGridCurrent()
            LoadListAvail()
        Catch ex As Exception
            MsgBox("Error Updating Database", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub
    Private Sub GridCurrent_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCurrent.CellContentClick
        If Me.GridCurrent.CurrentRow.Cells("Status").Value = "QQ" Then Me.LckLblApprove.Visible = True
        If Me.GridCurrent.CurrentRow.Cells("Status").Value <> "XX" Then Me.LblDeactive.Visible = True
        If Me.GridCurrent.CurrentRow.Cells("Status").Value = "OK" Then Me.LblChangeHL.Visible = True
    End Sub
    Private Sub ChangeStatus(ByVal pStatus As String)
        If Me.GridCurrent.CurrentRow.Cells("Status").Value = "QQ" AndAlso pStatus = "XX" Then
            cmd.CommandText = "delete from Cust_channel_Level where recid=" & Me.GridCurrent.CurrentRow.Cells("recID").Value
            cmd.ExecuteNonQuery()
            LoadGridCurrent()
        Else
            cmd.CommandText = ChangeStatus_ByID("Cust_channel_Level", pStatus, Me.GridCurrent.CurrentRow.Cells("recID").Value)
            cmd.ExecuteNonQuery()
            Me.GridCurrent.CurrentRow.Cells("Status").Value = pStatus
            ChangeGridCurrentColor(Me.GridCurrent.CurrentRow.Index)
        End If
    End Sub
    Private Sub LblApprove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblApprove.LinkClicked
        Dim MyAns As Int16, strSQL As String = ""
        MyAns = MsgBox("Be Careful. This Will Change Status of All QQs to OK. Abort This Action?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
        If MyAns = vbYes Then Exit Sub
        For i As Int16 = 0 To Me.GridCurrent.RowCount - 1
            If Me.GridCurrent.Item("Status", i).Value = "QQ" Then
                strSQL = strSQL & ";" & ChangeStatus_ByID("Cust_channel_Level", "OK", Me.GridCurrent.Item("recID", i).Value)
            End If
        Next
        cmd.CommandText = strSQL.Substring(1)
        cmd.ExecuteNonQuery()
        LoadGridCurrent()
        Me.LckLblApprove.Visible = False
    End Sub

    Private Sub LblDeactive_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDeactive.LinkClicked
        ChangeStatus("XX")
        Me.LblDeactive.Visible = False
    End Sub
    Private Sub ChkQOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkQOnly.Click
        Me.ChkActiveOnly.Checked = Not Me.ChkQOnly.Checked
        LoadGridCurrent()
    End Sub

    Private Sub ChkActiveOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkActiveOnly.Click
        LoadGridCurrent()
    End Sub

    Private Sub TabPage2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Enter
        LoadGridPolicy()
    End Sub
    Private Sub LoadGridPolicy()

        Dim strSQL As String
        strSQL = "select TRXType, Base, FareFrom, FareThrough, SFType, VAL, MinVal, MaxVal, Refundable,  "
        strSQL = strSQL & "ValidFrom, ValidThru from ServiceFee"
        strSQL = strSQL & " where SBU+al+Channel+CustLevel='"
        strSQL = strSQL & Me.CmbSBU.Text & Me.CmbAL.Text & Me.CmbChannel.Text & Me.CmbLevel.Text & "'"
        If Me.CmbFilter.Text <> "" AndAlso Me.CmbFilter.Text.Substring(0, 1) = "4" Then
            strSQL = strSQL & " and status ='XX'  "
        Else
            strSQL = strSQL & " and status <>'XX'  "
            If Me.CmbFilter.Text = "" OrElse Me.CmbFilter.Text.Substring(0, 1) = "1" Then
                strSQL = strSQL & " and '" & Now.Date & "' between ValidFrom and ValidThru"
            ElseIf Me.CmbFilter.Text.Substring(0, 1) = "2" Then
                strSQL = strSQL & " and ValidFrom > '" & Now.Date & "' "
            ElseIf Me.CmbFilter.Text.Substring(0, 1) = "3" Then
                strSQL = strSQL & " and ValidThru <'" & Now.Date & "' "
            End If
        End If
        Me.GridSF.DataSource = GetDataTable(strSQL)
        Me.GridSF.Columns("SFtype").Width = 36
        Me.GridSF.Columns("TRXtype").Width = 56
        Me.GridSF.Columns("Base").Width = 36
        Me.GridSF.Columns("FareFrom").Width = 56
        Me.GridSF.Columns("FareThrough").Width = 56
        Me.GridSF.Columns("MinVAL").Width = 56
        Me.GridSF.Columns("MaxVal").Width = 56
        Me.GridSF.Columns("VAL").Width = 32
        Me.GridSF.Columns("refundable").Width = 75
        Me.GridSF.Columns("ValidFrom").Width = 75
        Me.GridSF.Columns("VAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("FareFrom").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("FareThrough").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("MinVAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSF.Columns("MaxVAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        strSQL = "select COC, DType, VAL, ValidFrom, ValidThru  "
        strSQL = strSQL & " from Cust_Discount"
        strSQL = strSQL & " where SBU+al+Channel+CustLevel='"
        strSQL = strSQL & Me.CmbSBU.Text & Me.CmbAL.Text & Me.CmbChannel.Text & Me.CmbLevel.Text & "'"
        If Me.CmbFilter.Text <> "" AndAlso Me.CmbFilter.Text.Substring(0, 1) = "4" Then
            strSQL = strSQL & " and status ='XX'  "
        Else
            strSQL = strSQL & " and status <>'XX'  "
            If Me.CmbFilter.Text = "" OrElse Me.CmbFilter.Text.Substring(0, 1) = "1" Then
                strSQL = strSQL & " and '" & Now.Date & "' between ValidFrom and ValidThru"
            ElseIf Me.CmbFilter.Text.Substring(0, 1) = "2" Then
                strSQL = strSQL & " and ValidFrom > '" & Now.Date & "' "
            ElseIf Me.CmbFilter.Text.Substring(0, 1) = "3" Then
                strSQL = strSQL & " and ValidThru <'" & Now.Date & "' "
            End If
        End If
        Me.GridDiscount.DataSource = GetDataTable(strSQL)
        Me.GridDiscount.Columns("Dtype").Width = 56
        Me.GridDiscount.Columns("ValidFrom").Width = 75
        Me.GridDiscount.Columns("VAL").Width = 56
        Me.GridDiscount.Columns("VAL").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    Private Sub LblChangeHL_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblChangeHL.LinkClicked
        Dim OverLap As String, MyAns As Int16
        OverLap = CheckOverLap(Me.GridCurrent.CurrentRow.Cells("CustID").Value, Me.GridCurrent.CurrentRow.Cells("RecID").Value)
        If OverLap.Substring(0, 1) = "1" Then
            MyAns = MsgBox("Overlap: " & OverLap.Substring(1) & ". Wanna Continue?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, msgTitle)
            If MyAns = vbNo Then GoTo ExitHere
        ElseIf OverLap.Substring(0, 1) = "2" Then
            MsgBox("Overlap: " & OverLap.Substring(1) & ". Check Your Input or Change The Current One Before Continue", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, msgTitle)
            GoTo ExitHere
        End If

        ChangeStatus("XX")
        cmd.CommandText = "insert cust_Channel_Level (CustID, Channel, AL, CustLevel, SBU, FstUser, ValidFrom, ValidThru) " & _
                " select CustID, Channel,AL, CustLevel, SBU, '" & myStaff.SICode & "','" & Me.TxtValidFrm.Value.Date & "','" & _
                Me.TxtValidThru.Value.Date & " 23:59' from cust_Channel_Level where recid=" & Me.GridCurrent.CurrentRow.Cells("recID").Value
        cmd.ExecuteNonQuery()
        Me.TxtValidThru.Focus()
ExitHere:
        LoadGridCurrent()
    End Sub

    Private Sub CmbChannel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbChannel.SelectedIndexChanged
        Dim strDK As String, StrSQL As String, tmpSQL As String
        strDK = Me.CmbSBU.Text + Me.CmbAL.Text + Me.CmbChannel.Text
        strSQL = "select distinct CustLevel as VAL from ServiceFee where status <>'XX' and SBU+AL+Channel='"
        strSQL = strSQL & strDK & "' UNION "
        strSQL = strSQL & "select distinct CustLevel as VAL from Cust_Discount where status <>'XX' and SBU+AL+Channel='"
        StrSQL = StrSQL & strDK & "' "

        tmpSQL = strSQL
        strSQL = strSQL & " UNION "
        strSQL = strSQL & "select  VAL from MISC where CAT='Level' and val not in ("
        strSQL = strSQL & tmpSQL & ")"
        If Me.CmbChannel.Text = "WK" Then
            Me.CmbLevel.Items.Add("WKI")
            Me.CmbLevel.Enabled = False
            Me.CmbLevel.Text = "WKI"
        Else
            LoadCmb_MSC(Me.CmbLevel, strSQL)
            Me.CmbLevel.Enabled = True
        End If
    End Sub

    Private Sub LblCheckUnCheckALL_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCheckUnCheckALL.LinkClicked
        If Me.LblCheckUnCheckALL.Tag = "CHK" Then
            Me.LblCheckUnCheckALL.Tag = "UNCHK"
            Me.LblCheckUnCheckALL.Text = "UnCheckAll"
            checkUnCheckAll(True)
        Else
            Me.LblCheckUnCheckALL.Tag = "CHK"
            Me.LblCheckUnCheckALL.Text = "CheckAll"
            checkUnCheckAll(False)
        End If
    End Sub

    Private Sub CmbFilter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbFilter.SelectedIndexChanged
        LoadGridPolicy()
    End Sub
    Private Sub ChkBulkEnxtend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkBulkEnxtend.Click
        Me.ChkBulkEnxtend.Enabled = False
        Me.LblExtend.Visible = True
        Me.Label3.Text = "OLD ValThru"
        Me.Label4.Text = "NEW ValThru"
        Me.Label3.Left = 307
        Me.Label4.Left = 464
    End Sub

    Private Sub LblExtend_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblExtend.LinkClicked
        Dim myAns As Int16, strDK As String
        myAns = MsgBox("This is Bulk Update. Plz Check All the Input B4 Click NO" & vbCrLf & "Need Second Thought?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, msgTitle)
        If myAns = vbYes Then Exit Sub
        strDK = " status='OK' and SBU+AL+Channel+CustLevel='"
        strDK = strDK & Me.CmbSBU.Text + Me.CmbAL.Text + Me.CmbChannel.Text + Me.CmbLevel.Text & "'"
        strDK = strDK & " and validThru ='" & Me.TxtValidFrm.Value.Date & " 23:59'"

        cmd.CommandText = "insert cust_Channel_Level (CustID, Channel, AL, CustLevel, ValidFrom, ValidThru, Status, " & _
            "SBU, FstUser) Select CustID, channel, AL, CustLevel, ValidFrom, '" & Me.TxtValidThru.Value.Date & _
            " 23:59','QQ', SBU,'" & myStaff.SICode & "' from cust_channel_level where " & strDK & ";" & _
            ChangeStatus_ByDK("cust_Channel_Level", "XX", " where " & strDK)
        cmd.ExecuteNonQuery()

        Me.LblExtend.Visible = False
        Me.Label3.Text = "Valid From"
        Me.Label4.Text = "Valid Thru"
        Me.Label3.Left = 323
        Me.Label4.Left = 477
    End Sub
End Class