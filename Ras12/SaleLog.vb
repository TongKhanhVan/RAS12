Imports Microsoft.Office.Interop.Outlook
Public Class SaleLog
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private MyRTB As RichTextBox
    Private Sub SaleLog_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub SaleLog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ChannelList As String = myStaff.CAccess
        ChannelList = ChannelList.Replace("('", "")
        ChannelList = ChannelList.Replace("')", "")
        ChannelList = ChannelList.Replace("','", "_")
        LoadGridCust()
        Me.txtRptTimeFrm.Value = Now.Date.AddDays(-7)
        Me.txtRptTimeThru.Value = Now.Date
        If myStaff.SupOf <> "" Then
            Me.ChkSelectCust.Checked = False
            LoadGridDoneCall()
        Else
            Me.ChkSelectCust.Checked = True
            Me.ChkSelectCust.Visible = False
        End If
        LoadGridPending()
        For i As Int16 = 0 To ChannelList.Split("_").Length - 1
            If ChannelList.Split("_")(i) <> "WK" Then Me.CmbCustType.Items.Add(ChannelList.Split("_")(i))
        Next
    End Sub
    Private Sub LoadGridPending()
        Dim strQry As String = "select RecID, Subject, FrmTime, PreCall from GSAlog where status='OK'"
        If myStaff.SupOf = "" Then
            strQry = strQry & "and fstUser='" & myStaff.SICode & "'"
        End If
        Me.GridPendingCall.DataSource = GetDataTable(strQry)
        Me.GridPendingCall.Columns(0).Visible = False
        Me.GridPendingCall.Columns(1).Width = 200
        Me.GridPendingCall.Columns(3).Visible = False
        Me.LblDone.Visible = False
        Me.LblCancel.Visible = False
        Me.LblChange.Visible = False
        Me.LblUpdatePrepar.Visible = False
        Me.TxtPreView.Text = ""
        Me.TxtPost.Text = ""
    End Sub
    Private Function CreatePointmentItem(pSubj As String, pLocation As String, pStartDate As Date, pEndDate As Date, Optional pBody As String = "", Optional pRequireAttendees As String = "") As Boolean
        Dim outApp As Application
        Dim outAppoint As AppointmentItem
        If pStartDate >= pEndDate Or pSubj = "" Then
            MessageBox.Show("Empty Subject or Illogic Date/Time.", msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        outApp = CreateObject("Outlook.Application")
        outAppoint = outApp.CreateItem(OlItemType.olAppointmentItem)
        With outAppoint
            .Subject = pSubj
            .Location = pLocation & " Office"
            .Start = pStartDate
            .End = pEndDate
            .Body = pBody
            .ReminderSet = True
            .Save()
        End With
        Return True
    End Function

    Private Sub LblMakeAppt_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblMakeAppt.LinkClicked
        Dim startTime As Date = Me.txtStart.Value
        If CreatePointmentItem(Me.TxtSubj.Text, Me.GridCust.CurrentRow.Cells(1).Value, startTime, Me.txtEnd.Value) Then
            Me.LblMakeAppt.Visible = False
            cmd.CommandText = "insert GSALog (CustID, CustShortName, Subject, FrmTime, ToTime, PreCall, OverPhone, fstUser) values" & _
                "(@CustID, @CustShortName, @Subj, @FrmTime, @ToTime, @PreCall, @OverPhone, @fstUser)"
            cmd.Parameters.Clear()
            cmd.Parameters.Add("@CustID", SqlDbType.Int).Value = Me.GridCust.CurrentRow.Cells("RecID").Value
            cmd.Parameters.Add("@CustShortName", SqlDbType.VarChar).Value = Me.GridCust.CurrentRow.Cells("CustShortName").Value
            cmd.Parameters.Add("@FrmTime", SqlDbType.DateTime).Value = startTime
            cmd.Parameters.Add("@ToTime", SqlDbType.DateTime).Value = Me.txtEnd.Value
            cmd.Parameters.Add("@Subj", SqlDbType.NVarChar).Value = Me.TxtSubj.Text
            cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
            cmd.Parameters.Add("@OverPhone", SqlDbType.Bit).Value = IIf(Me.ChkByPhone.Checked, 1, 0)
            cmd.Parameters.Add("@PreCall", SqlDbType.NVarChar).Value = Me.TxtPre.Rtf
            cmd.ExecuteNonQuery()
            MsgBox("Appointment Has Been Added To Your Calendar", MsgBoxStyle.Information, msgTitle)
            Me.TxtSubj.Text = ""
            Me.TxtSubj.Enabled = False
        End If
    End Sub

    Private Sub OptTA_Click(sender As Object, e As EventArgs) Handles OptCust.Click, OptLead.Click
        LoadGridCust()
    End Sub
    Private Sub LoadGridCust()
        Dim strSQL As String = "select RecID, CustShortName, CustAddress from Customerlist where recid >0 " & _
            "and recid in (select CustID from cust_Detail where status+CAT='OKChannel' and VAL='" & Me.CmbCustType.Text & "')"
        If Me.OptCust.Checked Then
            strSQL = strSQL & " and Status='OK' "
            If myStaff.SupOf = "" AndAlso InStr("TA_TO_CA", Me.CmbCustType.Text) > 0 Then
                strSQL = strSQL & " and RecID in (select CustID from cust_detail where status+CAT='OKPIVOT' and VAL='" & myStaff.SICode & "') "
            End If
        ElseIf Me.OptLead.Checked Then
            strSQL = strSQL & " and Status='QQ'"
        End If
        Me.GridCust.DataSource = GetDataTable(strSQL)
        Me.GridCust.Columns(0).Visible = False
        Me.GridCust.Columns(2).Width = 256
        Me.LblMakeAppt.Visible = False
    End Sub

    Private Sub GridCust_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridCust.CellClick
        LoadGridDoneCall()
        LoadGridPending()
        Me.TxtSubj.Enabled = True
        Dim dteLstOne As String = Format(DateAdd("m", -1, Now), "dd-MMM-yy")
        Dim dteLstTree As String = Format(DateAdd("m", -3, Now), "dd-MMM-yy")
        Dim dteLstSix As String = Format(DateAdd("m", -6, Now), "dd-MMM-yy")
        Dim strDK As String = "CustID=" & Me.GridCust.CurrentRow.Cells("RecID").Value & " and status='RR' and FrmTime>'"
        Dim LstOne As Int16 = ScalarToInt("GSALog", "isnull(count(*),0)", strDK & dteLstOne & "'")
        Dim LstTree As Int16 = ScalarToInt("GSALog", "isnull(count(*),0)", strDK & dteLstOne & "'")
        Dim LstSix As Int16 = ScalarToInt("GSALog", "isnull(count(*),0)", strDK & dteLstOne & "'")
        Dim StartCursor As Int16, ArrAgenda(10) As String
        Dim newFontStyle As System.Drawing.FontStyle = FontStyle.Bold
        Me.LblSum.Text = LstOne.ToString & " Calls/Last 1M. " & LstTree & " Calls/Last 3M. " & LstSix & " Calls/Last 6M"
        Me.TxtPre.Text = ""
        ArrAgenda(0) = "1. Review performance and Customer infor"
        ArrAgenda(1) = "2. Define goals of the visit"
        ArrAgenda(2) = "3. Define steps to achieve set goals"
        ArrAgenda(3) = "4. Research for additional information"
        ArrAgenda(4) = "5. Get presentation and collaterals ready"
        ArrAgenda(5) = "6. Define roles of each participants"
        ArrAgenda(6) = "7. Send preCall docs to supervisor for advice"
        ArrAgenda(7) = "8. Make appointment"
        ArrAgenda(8) = "9. Practice"

        For i As Int16 = 0 To 8
            Me.TxtPre.Text = Me.TxtPre.Text & ArrAgenda(i) & vbCrLf & " " & vbCrLf & " " & vbCrLf
        Next
        For i As Int16 = 0 To 24 Step 3
            StartCursor = Me.TxtPre.GetFirstCharIndexFromLine(i)
            Me.TxtPre.SelectionStart = StartCursor
            Me.TxtPre.SelectionLength = Me.TxtPre.Lines(i).Length
            Me.TxtPre.SelectionColor = Color.CornflowerBlue
            Me.TxtPre.SelectionFont = New Drawing.Font(Me.TxtPre.SelectionFont.FontFamily, Me.TxtPre.SelectionFont.Size, newFontStyle)
        Next
    End Sub
    Private Sub LoadGridDoneCall()
        Dim strQry As String = "select FstUser, CustShortName,Subject, frmTime, AfterCall, RecID from GSAlog where status='RR'"
        If myStaff.SupOf = "" Or Me.ChkSelectCust.Checked Then
            strQry = strQry & " and custid=" & Me.GridCust.CurrentRow.Cells("recID").Value
        End If
        strQry = strQry & " and frmtime between '" & Format(Me.txtRptTimeFrm.Value, "dd-MMM-yy") & "' and '" & Format(Me.txtRptTimeThru.Value, "dd-MMM-yy") & " 23:59'"
        Me.GridDoneCall.DataSource = GetDataTable(strQry)
        Me.GridDoneCall.Columns(4).Visible = False
        Me.GridDoneCall.Columns(5).Visible = False
        Me.GridDoneCall.Columns(0).Width = 45
        Me.GridDoneCall.Columns(2).Width = 200
        Me.txtPostView.Text = ""
        Me.LckLblDelete.Visible = False
    End Sub

    Private Sub GridDoneCall_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridDoneCall.CellClick
        Me.txtPostView.Text = Me.GridDoneCall.CurrentRow.Cells(4).Value
        On Error Resume Next
        Me.txtPostView.Rtf = Me.GridDoneCall.CurrentRow.Cells(4).Value
        On Error GoTo 0
        Me.LckLblDelete.Visible = True
    End Sub

    Private Sub GridPendingCall_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridPendingCall.CellClick
        Me.LblUpdatePrepar.Visible = True
        Me.LblChange.Visible = True
        Me.TxtPreView.Text = Me.GridPendingCall.CurrentRow.Cells(3).Value
        On Error Resume Next
        Me.TxtPreView.Rtf = Me.GridPendingCall.CurrentRow.Cells(3).Value
        On Error GoTo 0
    End Sub
    Private Sub LblDone_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDone.LinkClicked
        UpdatePending("RR")
        LoadGridPending()
    End Sub
    Private Sub LblUpdatePrepar_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdatePrepar.LinkClicked
        cmd.CommandText = "update GSAlog set preCall=@PreCall where RecID=@RecID"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@PreCall", SqlDbType.NVarChar).Value = Me.TxtPreView.Text
        cmd.Parameters.Add("@RecID", SqlDbType.Int).Value = Me.GridPendingCall.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridPending()
    End Sub
    Private Sub LblCancel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCancel.LinkClicked
        DeletePointmentItem("DELETE", Me.GridPendingCall.CurrentRow.Cells("Subject").Value, Me.GridPendingCall.CurrentRow.Cells("FrmTime").Value)
        UpdatePending("XX")
        LoadGridPending()
    End Sub
    Private Sub UpdatePending(pStatus As String)
        If Me.TxtPost.Text = "" Then Exit Sub
        cmd.CommandText = "update GSAlog set AfterCall=@AfterCall, LstUser=@LstUser, LstUpdate=getdate(), Status=@Status where RecID=@RecID"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@AfterCall", SqlDbType.NVarChar).Value = Me.TxtPost.Rtf
        cmd.Parameters.Add("@LstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@Status", SqlDbType.VarChar).Value = pStatus
        cmd.Parameters.Add("@RecID", SqlDbType.Int).Value = Me.GridPendingCall.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridPending()
    End Sub
    Private Sub DeletePointmentItem(pWhat As String, oldSubject As String, oldStartDate As String, Optional newStartDate As Date = Nothing, Optional newEndDate As Date = Nothing)
        Dim outApp As Application = CreateObject("Outlook.Application")
        Dim outCalendar As MAPIFolder = outApp.GetNamespace("Mapi").GetDefaultFolder(OlDefaultFolders.olFolderCalendar)
        Dim outItems As Items = outCalendar.Items
        outItems.Sort("Start")
        If oldSubject = "" Or oldStartDate = "" Then
            MessageBox.Show("Cac gia tri truyen vao khong hop le", msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim strConditon As String = "", strSubject As String = "", strStart As String = ""
        strSubject = String.Format("[Subject] = '{0}'", oldSubject)
        strConditon = String.Format("[Start] = '{0}'", oldStartDate)
        strConditon = strConditon & " AND " & strSubject
        Dim outAppoint As AppointmentItem = outItems.Find(strConditon)
        If outAppoint Is Nothing Then Exit Sub
        With outAppoint
            If pWhat = "DELETE" Then
                .Delete()
            Else
                .Start = newStartDate
                .End = newEndDate
                .Save()
            End If
        End With
        outAppoint = Nothing
        outCalendar = Nothing
        outApp = Nothing
    End Sub
    Private Sub LblChange_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblChange.LinkClicked
        Dim startTime As Date = Me.txtNewStart.Value
        If Me.txtNewStart.Value > Me.TxtNewEnd.Value Or Me.txtNewStart.Value < Now.AddHours(4) Then Exit Sub
        DeletePointmentItem("CHANGE", Me.GridPendingCall.CurrentRow.Cells("Subject").Value, Me.GridPendingCall.CurrentRow.Cells("FrmTime").Value, startTime, Me.TxtNewEnd.Value)
        cmd.CommandText = "update gsalog set frmTime='" & Me.txtNewStart.Value & "', toTime='" & Me.TxtNewEnd.Value & "' where recID=" & Me.GridPendingCall.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridPending()
    End Sub

    Private Sub LblRefresh_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblRefresh.LinkClicked
        LoadGridDoneCall()
    End Sub

    Private Sub TxtSubj_TextChanged(sender As Object, e As EventArgs) Handles TxtSubj.TextChanged
        If Me.TxtSubj.Text <> "" Then Me.LblMakeAppt.Visible = True
    End Sub

    Private Sub TxtPost_TextChanged(sender As Object, e As EventArgs) Handles TxtPost.TextChanged
        Me.LblCancel.Visible = True
        Me.LblDone.Visible = True
    End Sub

    Private Sub TxtNewEnd_Enter(sender As Object, e As EventArgs) Handles TxtNewEnd.Enter
        Me.TxtNewEnd.Value = Me.txtNewStart.Value.AddHours(1)
    End Sub

    Private Sub txtEnd_Enter(sender As Object, e As EventArgs) Handles txtEnd.Enter
        Me.txtEnd.Value = Me.txtStart.Value.AddHours(1)
    End Sub

    Private Sub CmbCustType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCustType.SelectedIndexChanged
        LoadGridCust()
    End Sub
    Private Sub TxtPre_Enter(sender As Object, e As EventArgs) Handles TxtPre.Enter, TxtPost.Enter, TxtPreView.Enter
        Dim rtb As RichTextBox = CType(sender, RichTextBox)
        MyRTB = rtb
    End Sub
    Private Sub TB_Color_Click(sender As Object, e As EventArgs) Handles TB_Color.Click
        If ColorDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            MyRTB.SelectionColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub TB_Bold_Click(sender As Object, e As EventArgs) Handles TB_Bold.Click
        Dim currentFont As System.Drawing.Font = MyRTB.SelectionFont
        Dim newFontStyle As System.Drawing.FontStyle
        If MyRTB.SelectionFont.Bold = True Then
            newFontStyle = currentFont.Style - Drawing.FontStyle.Bold
        Else
            newFontStyle = currentFont.Style + Drawing.FontStyle.Bold
        End If
        MyRTB.SelectionFont = New Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle)
    End Sub
    Private Sub TB_Underline_Click(sender As Object, e As EventArgs) Handles TB_Underline.Click
        Dim currentFont As System.Drawing.Font = MyRTB.SelectionFont
        Dim newFontStyle As System.Drawing.FontStyle
        If MyRTB.SelectionFont.Underline = True Then
            newFontStyle = currentFont.Style - Drawing.FontStyle.Underline
        Else
            newFontStyle = currentFont.Style + Drawing.FontStyle.Underline
        End If
        MyRTB.SelectionFont = New Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle)
    End Sub
    Private Sub TB_Bullet_Click(sender As Object, e As EventArgs) Handles TB_Bullet.Click
        MyRTB.SelectionBullet = Not TxtPre.SelectionBullet
    End Sub
    Private Sub LblDelete_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LckLblDelete.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("GSALog", "XX", Me.GridDoneCall.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        Me.LckLblDelete.Visible = False
    End Sub

End Class