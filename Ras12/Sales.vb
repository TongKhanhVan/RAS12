Public Class Sales
    Private DoWhat As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private XlsProcess As Process
    Public Sub New()
        InitializeComponent()
    End Sub
    Public Sub New(ByVal parDowhat As String)
        DoWhat = parDowhat
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub Sales_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            XlsProcess.Kill()
        Catch ex As Exception
        End Try
        AutoUploadPPD2TransVietVN()
        Me.Dispose()
    End Sub

    Private Sub Sales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        initPanels()
        If InStr("FB_RBD", DoWhat) > 0 Then
            Me.PnlFB_RBD.Visible = True
            LoadData_S("**")
        ElseIf DoWhat = "RPT" Then
            Me.StatusFeedback.Text = "Plz Close All Open Workbooks Before Clicking RunReport"
            LoadCmb_MSC(Me.CmbCounter, "Select AL as VAL from AIrline where mauso <>'' and Mauso <>'000'")
            Dim strSQL As String = "select RecID as VAL, CustShortName as DIS from Customerlist where recid >0 and Status='OK' "
            strSQL = strSQL & " and RecID in (select CustID from cust_detail where status+CAT='OKChannel' and VAL in ('TA','TO','CA')) "
            If myStaff.SupOf = "" Then
                strSQL = strSQL & " and RecID in (select CustID from cust_detail where status+CAT='OKPIVOT' and VAL='" & myStaff.SICode & "') "
            End If
            LoadCmb_VAL(Me.CmbCust, strSQL)
            Me.TxtFrm.Value = Now.Date.AddYears(-1)
            Me.GrpReport.Visible = True
            Me.CmbTime.Text = "This MTD"
        End If
    End Sub
    Private Sub LoadData_S(ByVal parAL As String)
        Dim strSQL As String = "select RecID, AL, Raw, ADJ from FB_RBD_QT where cat='" & DoWhat & "'" & _
            " and AL='" & parAL & "'"
        Me.GridFB_RBD.DataSource = GetDataTable(strSQL)
        For i As Int16 = 0 To 2
            Me.GridFB_RBD.Columns(i).ReadOnly = True
        Next
    End Sub
    Private Sub initPanels()
        Me.PnlFB_RBD.Top = 0
        Me.PnlFB_RBD.Left = 0
        Me.GrpReport.Top = 0
        Me.GrpReport.Left = 0
        Me.PnlFB_RBD.Height = 444
        Me.PnlFB_RBD.Width = 808
    End Sub
    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        For r As Int16 = 0 To Me.GridFB_RBD.RowCount - 1
            cmd.CommandText = "Update FB_RBD_QT set ADJ='" & Me.GridFB_RBD.Item("ADJ", r).Value & "'," & _
                " FstUser='" & myStaff.SICode & "', FstUpdate =getdate() where RecID=" & Me.GridFB_RBD.Item("RecID", r).Value
            cmd.ExecuteNonQuery()
        Next
        Me.CmdSave.Enabled = False
    End Sub
    Private Sub CmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdFilter.Click
        If InStr(myStaff.AAccess, "YY") + InStr(myStaff.AAccess, Me.txtAL.Text) > 0 Then
            LoadData_S(Me.txtAL.Text)
        Else
            MsgBox("Invalid Input", MsgBoxStyle.Critical, msgTitle)
        End If
    End Sub

    Private Sub ChkAL_CheckedChanged(sender As Object, e As EventArgs) Handles ChkAL.CheckedChanged
        Me.CmbCounter.Enabled = Me.ChkAL.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles ChkCust.CheckedChanged
        Me.CmbCust.Enabled = Me.ChkCust.Checked
    End Sub

    Private Sub LblCancel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCancel.LinkClicked
        Me.Close()
    End Sub

    Private Sub LblRun_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblRun.LinkClicked
        If Me.txtThru.Value < Me.TxtFrm.Value Then Exit Sub
        Me.LblRun.Enabled = False
        Dim dteFrom As Date, dteThru As Date, BackMonth As Int16
        dteThru = Now.AddDays(-1)
        If myStaff.City = "SGN" Then
            cmd.CommandText = "select count(*) " & DKDataConvertMktg_RAS & " and DOI <'" & Format(dteThru, "dd-MMM-yy") & "'"
            Dim NewlyAddedRec As Int16 = cmd.ExecuteScalar
            cmd.CommandText = "select count(*) " & DKDataConvertMktg_BSP
            NewlyAddedRec = NewlyAddedRec + cmd.ExecuteScalar
            If NewlyAddedRec > 4 Then
                Me.StatusFeedback.ForeColor = Color.Blue
                Me.StatusFeedback.Text = "Processing Newly Added Data 4 More Accuracy. Be Patient..."
                Me.Refresh()
                GoDataConvert()
                Me.StatusFeedback.ForeColor = Color.Fuchsia
                Me.StatusFeedback.Text = "Querying Data for Report..."
                Me.Refresh()
            End If
        End If
        Dim tmpTableName As String = "##" & myStaff.SICode
        Dim Agt As String = IIf(Me.ChkCust.Checked, Me.CmbCust.Text, "ALL")
        Dim AL As String = IIf(Me.ChkAL.Checked, Me.CmbCounter.Text, "ALL")
        On Error Resume Next
        cmd.CommandText = "drop table " & tmpTableName
        cmd.ExecuteNonQuery()
        SaveSetting("R12", "GSA", "tmpTable", tmpTableName)
        SaveSetting("R12", "GSA", "SVR", MySession.ServerIP)
        XlsProcess.Kill()

        On Error GoTo 0
        dteFrom = DefineDateTime("FRM")
        dteThru = DefineDateTime("THR")
        TaoDataToTmpTable(0, dteFrom, dteThru, AL, Agt, tmpTableName)
        If Me.ChkCompare.Checked And Me.CmbTime.Text <> "Custome" Then
            If Me.CmbTime.Text <> "Custome" Then
                BackMonth = DefineBackMonth(Me.CmbTime.Text.Substring(5, 1))
                InsertDataSoSanh(1, dteFrom, dteThru, BackMonth, AL, Agt, tmpTableName)
                If Me.CmbTime.Text.Substring(5, 1) <> "Y" Then InsertDataSoSanh(2, dteFrom, dteThru, -12, AL, Agt, tmpTableName)
            End If
        End If
        Me.StatusFeedback.ForeColor = Color.Blue
        Me.StatusFeedback.Text = "Done."
        Me.Refresh()
        XlsProcess = Process.Start("X:\Xls from SQL\R12_SalesPerformance.xlt")
        Me.StatusFeedback.Text = ""
        Me.LblRun.Enabled = True
    End Sub
    Private Function DefineBackMonth(pLoai As String) As Int16
        If pLoai = "M" Then
            Return -1
        ElseIf pLoai = "Q" Then
            Return -3
        ElseIf pLoai = "Y" Then
            Return -12
        End If
    End Function
    Private Sub InsertDataSoSanh(pStep As Int16, pFrm As Date, pThru As Date, pBackMonth As Int16, pAL As String, pAgt As String, pTblName As String)
        Dim vFrm As Date, vThru As Date
        vFrm = pFrm.AddMonths(pBackMonth)
        vThru = pThru.AddMonths(pBackMonth)
        TaoDataToTmpTable(pStep, vFrm, vThru, pAL, pAgt, pTblName)
    End Sub
    Private Sub TaoDataToTmpTable(pStep As Int16, pFrom As Date, pThru As Date, pAL As String, pAgt As String, pTblName As String)
        If pStep = 0 Then
            cmd.CommandText = "select *,'Ky BCao' as KyBC into " & pTblName & " from forexcelreport (@frm,@Thru,@AL,@Agt)"
        ElseIf pStep = 1 Then
            cmd.CommandText = "Insert into " & pTblName & " select *,'Ky Trc' from forexcelreport (@frm,@Thru,@AL,@Agt)"
        Else
            cmd.CommandText = "Insert into " & pTblName & " select *,'NamTrc' from forexcelreport (@frm,@Thru,@AL,@Agt)"
        End If
        cmd.Parameters.Clear()
        cmd.CommandTimeout = 256
        cmd.Parameters.Add("@frm", SqlDbType.DateTime).Value = pFrom
        cmd.Parameters.Add("@Thru", SqlDbType.DateTime).Value = pThru
        cmd.Parameters.Add("@AL", SqlDbType.VarChar).Value = pAL
        cmd.Parameters.Add("@Agt", SqlDbType.VarChar).Value = pAgt
        cmd.ExecuteNonQuery()
    End Sub
    Private Function defineFstMonthOfQ(pMonth As Int16) As Int16
        If pMonth < 4 Then Return 1
        If pMonth < 7 Then Return 4
        If pMonth < 10 Then Return 7
        Return 10
    End Function
    Private Function DefineDateTime(pLoai As String) As Date
        Dim KQ_Frm As Date, KQ_Thru As Date
        Dim FstMonthOfQ As Int16, Loai As String = Me.CmbTime.Text.Substring(5, 1), BackMonth As Int16
        FstMonthOfQ = defineFstMonthOfQ(Now.Month)
        If Me.CmbTime.Text = "Custome" Then
            KQ_Frm = Me.TxtFrm.Value.Date
            KQ_Thru = Me.txtThru.Value
        Else
            KQ_Thru = Now.Date
            If Loai = "M" Then
                KQ_Frm = DateSerial(Now.Year, Now.Month, 1)
            ElseIf Loai = "Q" Then
                KQ_Frm = DateSerial(Now.Year, FstMonthOfQ, 1)
            ElseIf Loai = "Y" Then
                KQ_Frm = DateSerial(Now.Year, 1, 1)
            End If

            If Me.CmbTime.Text.Contains("Last") Then
                KQ_Thru = KQ_Frm.AddDays(-1)
                BackMonth = DefineBackMonth(Loai)
                KQ_Frm = KQ_Frm.AddMonths(BackMonth)
            End If
        End If
        Return IIf(pLoai = "FRM", KQ_Frm, KQ_Thru)
    End Function
    Private Sub CmbTime_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTime.SelectedIndexChanged
        If Me.CmbTime.Text = "Custome" Then
            Me.TxtFrm.Enabled = True
            Me.ChkCompare.Checked = False
        Else
            Me.TxtFrm.Enabled = False
        End If
        Me.txtThru.Enabled = Me.TxtFrm.Enabled
        Me.ChkCompare.Enabled = Not Me.TxtFrm.Enabled
    End Sub
End Class


