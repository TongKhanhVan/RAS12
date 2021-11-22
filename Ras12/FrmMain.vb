Imports SharedFunctions.MySharedFunctions
Imports System.IO

Public Class frmMain
    'route change 0.0.0.0 mask 0.0.0.0 172.16.1.252
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private tblName As String
    Private MyCust As New objCustomer
    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If String.IsNullOrEmpty(myStaff.SICode) Then GoTo ResumeHere
        If Conn.State <> ConnectionState.Open Then Conn.Open()
        If myStaff.City = "SGN" Then
            cmd.CommandTimeout = 512
            If myStaff.SICode = "SYS" Then
                cmd.CommandText = "Exec CleanReportData"
                cmd.ExecuteNonQuery()
            ElseIf myStaff.Counter = "CWT" Then
                cmd.CommandText = "Exec CleanReportData; Exec updateBooker"
                cmd.ExecuteNonQuery()
            ElseIf myStaff.UGroup.Contains("S") Then
                cmd.CommandText = "Exec CleanReportData; Exec UpdateQT_FB"
                cmd.ExecuteNonQuery()
            End If
        ElseIf myStaff.City = "HAN" Then
            'cmd.ExecuteNonQuery()
        End If
        If myStaff.SICode = "SYS" OrElse myStaff.UGroup.Contains("S") Or myStaff.Counter = "CWT" Then
            GoDataConvert()
        End If
        LogInOut("")
        Conn.Close()
ResumeHere:
        Me.Dispose()
        End
    End Sub
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = "dd-MMM-yy"
        System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.LongTimePattern = "HH:mm"
        Conn_Web.ConnectionString = CnStr_TVW
        DDAN = Application.StartupPath

        ClearLogFile(30)

        MySession.ServerIP = "42.117.5.86"
        MySession.ServerIP = GetServerName(DDAN & "\RAS12config.txt")
        If HasNewerVersion_R12(Application.ProductVersion) Then
            End
        End If
        CnStr = "server=XXXX;uid=rasusers;pwd=Healthy@FoodR12;database=RAS12"
        CnStr = CnStr.Replace("XXXX", MySession.ServerIP)
        Conn.ConnectionString = CnStr
        AddHandler Conn.StateChange, AddressOf Conn_StateChange
        Conn.Open()
        SetAllMenuStatus(False)
        myStaff.App = "RAS"
        myStaff.CnStr = CnStr
        MyCust.CnStr = myStaff.CnStr
        MyAL.CnStr = CnStr

        InitPanel()
        GenComboValueMain()
        Me.CmbCity.Text = IIf(ScalarToString("MISC", "VAL", "cat='POS'") = "0", "SGN", "HAN")

        Me.txtLogInSIcode.Text = GetSetting("car", "dangkiem", "Type")
        If Me.txtLogInSIcode.Text <> "" Then Me.txtLogInPSW.Focus()
        Me.txtLogInPSW.Text = EnCode(GetSetting("COS", "Preference", "WSP"), "115")
        pubVarBackColor = GetColorForToday()
        Me.Text = Me.Text & " [" & MySession.ServerIP & "]"

        pstrVnDomCities = GetColumnValuesAsString("CityCode", "Airport", "where Country='VN'", "_")
    End Sub
    Private Sub Conn_StateChange(ByVal sender As Object, ByVal e As System.Data.StateChangeEventArgs)
        If e.CurrentState = ConnectionState.Open Then
            Me.statusConnSate.Text = "Connected"
            Me.statusConnSate.ForeColor = Color.Blue
        Else
            Me.statusConnSate.Text = "Disconnected"
            Me.statusConnSate.ForeColor = Color.Red
        End If
    End Sub
    Private Sub BarLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarLogIn.Click
        HideFrame()
        If Me.BarLogIn.Text = "Log In" Then
            Me.PnlLogIn.Height = 80
            Me.PnlLogIn.Visible = True
        Else
            Me.txtLogInPSW.Text = ""
            myStaff.LogOut()
            LogInOut("")
        End If
    End Sub
    Private Sub BarChangePassword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarChangePassword.Click
        HideFrame()
        Me.PnlChangePSW.Visible = True
    End Sub
    Private Sub CmdLogInCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdLogInCancel.Click
        LogInOut("")
    End Sub
    Private Sub CmdLogInOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdLogInOK.Click
        Me.txtLogInSIcode.Text = Me.txtLogInSIcode.Text.Replace("'", "")
        If Me.txtLogInSIcode.Text.Length < 3 Then Exit Sub
        Dim tmpSIcode As String = "", tmpPSW As String = "", tmpNewPSW = GenDefaultPSW(Me.txtLogInSIcode.Text, "R12")
        Dim PhaidoiPSW As Boolean = False
        myStaff.SICode = Me.txtLogInSIcode.Text.Replace("--", "")

        myStaff.SelectedDomain = CmbBiz.Items(CmbBiz.FindStringExact(CmbBiz.Text))
        If CmbBiz.Text = "" Then
            myStaff.SelectedDomain = CmbBiz.Items(CmbBiz.FindStringExact(Mid(myStaff.DAccess, myStaff.DAccess.Length - 2)))
        End If

        If myStaff.UID = 0 Then
            myStaff.SICode = ""
            MsgBox("Invalid Sign-In Code ", MsgBoxStyle.Critical, msgTitle)
            Me.PnlLogIn.Visible = True
            Exit Sub
        End If

        PhaidoiPSW = False
        If DateDiff(DateInterval.Day, myStaff.CreateDate, Now) > 2 Then
            If Me.txtLogInPSW.Text = tmpNewPSW Or Me.txtLogInPSW.Text = "Abcd1234" Or Me.txtLogInPSW.Text = "" Then
                PhaidoiPSW = True
            End If
        End If

        MyAL.Counter = myStaff.Counter
        If myStaff.DAccess.Length < 8 Then
            MySession.Domain = Strings.Right(myStaff.DAccess, 3)
            Me.CmbBiz.Text = MySession.Domain
        Else
            MySession.Domain = Me.CmbBiz.Text
        End If
        MySession.City = myStaff.City
        MySession.Location = myStaff.Location
        If Me.CmbCounter.Visible AndAlso Me.PnlLogIn.Height > 128 Then
            MySession.Counter = Me.CmbCounter.Text
        Else
            If Me.CmbBiz.Text = "TVS" Then
                MySession.Counter = myStaff.Counter
            Else
                MySession.Counter = Me.CmbBiz.Text
            End If
        End If
        If myStaff.PSW <> HashToFixedLen(Me.txtLogInPSW.Text) Then
            myStaff.SICode = ""
            MsgBox("Invalid SI Code or Password", MsgBoxStyle.Critical, msgTitle)
            Me.PnlLogIn.Visible = True
        ElseIf MySession.Domain = "" OrElse InStr("EDU_TVS_GSA", MySession.Domain) = 0 OrElse _
            MySession.Counter = "" OrElse MySession.City = "" OrElse InStr("HAN_SGN", MySession.City) = 0 OrElse _
            MySession.Location = "" Then
            myStaff.SICode = ""
            MsgBox("Invalid Config of Domain or Counter or City or Loacation", MsgBoxStyle.Critical, msgTitle)
            Me.PnlLogIn.Visible = True
        ElseIf PhaidoiPSW Then
            Me.BarChangePassword.Enabled = True
            Me.PnlLogIn.Visible = False
            MsgBox("You Have To Change Default Password Before Continue", MsgBoxStyle.Critical, msgTitle)
        ElseIf InStr(myStaff.DAccess, "YY") + InStr(myStaff.DAccess, MySession.Domain) + InStr(myStaff.City, MySession.City) = 0 Then
            MsgBox("Your Are not Assigned To This Counter/City", MsgBoxStyle.Critical, msgTitle)
            myStaff.LogOut()
            Me.PnlLogIn.Visible = True
        Else
            SaveSetting("car", "dangkiem", "Type", myStaff.SICode)
            cmd.CommandText = "Update tblUser set status='ON' where SICode='" & myStaff.SICode & "'"
            cmd.ExecuteNonQuery()
            LogInOut(myStaff.SICode)
            Me.statusUser.Text = "   User=" & myStaff.SICode & ". Domain=" & MySession.Domain
            HideMenu4HAN()
            MyCust.GenCustList()
            If myStaff.SICode = "SYS" Then
                Me.CmbCity.Enabled = True
                Me.CmbCounter.Enabled = True
                Me.CmbLocation.Enabled = True
            End If
            If myStaff.SupOf <> "" Then
                Me.CmdReportPrint.Enabled = True
            End If
            
        End If

        Me.StatusVersion.Text = "Version: " & Application.ProductVersion
    End Sub
    Private Sub HideMenu4HAN()
        If MySession.POSCode = "3" Then
            Me.PadCorpSales.Visible = False
            Me.PadMICE.Visible = False
            Me.BarSalesCall.Visible = False
            Me.BarAPGInv_KTT.Visible = False
            Me.BarInvChecker.Visible = False
            Me.BarIATAMaintain.Visible = False
            Me.BarGoToTSP.Visible = False
            Me.BarExportToTSP.Visible = False
        Else
            Me.BarUpdateForExQuay.Visible = False
        End If
    End Sub
    Private Sub CmdChangePSWOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdChangePSWOK.Click
        Dim tmpPSW As String = ""
        tmpPSW = ScalarToString("tblUser", "PSW", String.Format("Status<>'XX' and SICode='{0}'", myStaff.SICode))
        If tmpPSW <> HashToFixedLen(Me.txtChangePSWOLD.Text) Or _
            Me.txtChangePSWNEW.Text <> Me.txtChangePSWNew2.Text Then
            MsgBox("Invalid Old Password or New Passwords Are not Identical", , msgTitle)
        Else
            cmd.CommandText = "update tblUser set psw=@PSW where SICode=@SIcode"
            cmd.Parameters.Clear()
            cmd.Parameters.Add("@PSW", SqlDbType.VarChar).Value = HashToFixedLen(Me.txtChangePSWNew2.Text)
            cmd.Parameters.Add("@SIcode", SqlDbType.VarChar).Value = myStaff.SICode
            cmd.ExecuteNonQuery()
            Me.PnlChangePSW.Visible = False
        End If
    End Sub
    Private Sub CmdChangePSWCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        CmdChangePSWCancel.Click, CmdReportClose.Click
        HideFrame()
    End Sub

    Private Sub BarUserManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarUserManager.Click, BarUserMngerSales.Click, BarFOUserManager.Click
        Dim Mnu As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim f As New UserManagement(Mnu.Tag)
        f.ShowDialog()
    End Sub
    Private Sub BarFOIssue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarFOIssue.Click
        Dim f As Form
        For Each f In My.Application.OpenForms
            If f.GetType.Name = "FOissueTKT" Then
                f.Activate()
                Exit Sub
            End If
        Next
        f = New FOissueTKT()
        f.ShowDialog()
    End Sub
    Private Sub BarFORefund_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarFOEditMoney.Click, BarFOEditFOP.Click, BarReprintTRX.Click, BarFOPrintDeleteRCP.Click, BarFODataCorrection.Click
        Dim f As Form, RCP As String, b As ToolStripItem = CType(sender, ToolStripItem)
        For Each f In My.Application.OpenForms
            If f.GetType.Name = "FOissueTKT" Then
                f.Activate()
                Exit Sub
            End If
        Next
        RCP = InputBox("Enter Transaction Confirmation Number:", msgTitle)
        If RCP = "" Then Exit Sub
        RCP = RCP.ToUpper
        f = New FOissueTKT(b.Tag & "_" & RCP)
        f.ShowDialog()
    End Sub
    Private Sub DoReport()
        GenRPTName()
        HideFrame()
        Me.PnlReport.Visible = True
    End Sub
    Private Sub GenRPTName()
        Dim WhoIs As String = Me.CmdReportClose.Tag
        Dim strSQL As String = String.Format("select VAL, val as ReportName from MISC where Status='OK' and cat like '%{0}%-RPTName'", WhoIs)
        If WhoIs = "C" Then
            If InStr(myStaff.AAccess, "YY") = 0 Then
                strSQL = String.Format("{0} and (substring(val,4,2)='YY' or substring(val,4,2) in {1})", strSQL, myStaff.AAccess)
            End If
            Me.CmdReportPrint.Enabled = True
            Me.LckCmdRPTexport.Visible = False
        ElseIf WhoIs = "A" Then
            Me.LckCmdRPTexport.Visible = True
            Me.CmdReportPrint.Enabled = False
        ElseIf WhoIs = "S" Then
            Me.LckCmdRPTexport.Visible = False
            Me.CmdReportPrint.Enabled = False
        End If

        Me.GridRPTname.DataSource = GetDataTable(strSQL)
        Me.GridRPTname.Columns(0).Visible = False
        Me.GridRPTname.Columns(0).Width = 0
        Me.GridRPTname.Columns(1).Width = 200
        For i As Int16 = 0 To Me.GridRPTname.RowCount - 1
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("CC_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("CS_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("CT_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("SR_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("RP_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("YY_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("TS_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("AR_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("CR_", "")
            Me.GridRPTname.Item(1, i).Value = Me.GridRPTname.Item(1, i).Value.Replace("PP_", "")
        Next
    End Sub
    Private Sub BarFOReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarFOReport.Click
        Me.CmdReportClose.Tag = "C"
        DoReport()
    End Sub
    Private Sub CmdReportRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdReportRun.Click, LckCmdRPTexport.Click
        Dim fName As String, varFrm As Date, varTo As Date
        Dim cmd As Button = CType(sender, Button)
        Dim RPTNO As String, varPreview As String = "V"
        If InStr(cmd.Name.ToUpper, "EXPORT") > 0 Then varPreview = "F"
        If Me.GridRPTname.CurrentCell.Value = "" Or Me.CmbReportCounter.Text = "" Then
            MsgBox("Please Select Report Name or Airline/Counter", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        fName = Me.GridRPTname.CurrentRow.Cells(0).Value
        varFrm = Me.txtReportFrom.Text
        varTo = Me.txtReportTo.Text
        If InStr("GSA_EDU", MySession.Domain) > 0 Then
            RPTNO = Me.CmbReportCounter.Text
        Else
            RPTNO = "TS"
        End If
        InHoaDon(DDAN, fName, varPreview, RPTNO, varFrm, varTo, Me.CmbRPTcust.SelectedValue _
                 , Me.CmbReportCounter.Text, MySession.Domain, Me.CmdReportClose.Tag)
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarCorporateAccount.Click, BarCustList.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New AddCustomer(b.Tag)
        f.ShowDialog()
    End Sub
    Private Sub statusConnSate_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles statusConnSate.DoubleClick
        If Me.statusConnSate.Text = "Disconnected" Then
            Try
                Conn.Open()
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub BarViewAcc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarBOViewAcc.Click, BarViewTRXList.Click, BarFOView.Click, BarTRXListing_CS.Click, BarTRXListing_KTT.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New View(b.Tag)
        f.ShowDialog()
    End Sub

    Private Sub CloseBC_AL()
        Dim fName As String, MyAns As Integer, RPTNO As String
        Dim varFrm As Date, varTo As Date, LstRPT As String
        If Me.GridRPTname.CurrentCell.Value = "" Then
            MsgBox("Please Select Report Name", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        MyAns = MsgBox("Do You Want to Authorize Report and Print It?", MsgBoxStyle.YesNo Or MsgBoxStyle.Critical Or MsgBoxStyle.DefaultButton2, msgTitle)
        If MyAns = vbNo Then Exit Sub

        varTo = Me.txtReportTo.Text & " 23:59"
        varFrm = Me.txtReportFrom.Text & " 00:01"
        MyAns = 0
        MyAns = ScalarToInt("TKT", "count(RPTNO)", "DOI <'" & varTo & "' and RPTNo = '' and AL='" & Me.CmbReportCounter.Text & "'")
        fName = Me.GridRPTname.CurrentRow.Cells(0).Value
        RPTNO = Me.CmbReportCounter.Text
        RPTNO = RPTNO & Format(varTo.Month, "00") & Format(varTo, "yy")
        RPTNO = RPTNO & "_" & Format(varFrm, "dd") & "-" & Format(varTo, "dd")
        If MyAns = 0 Then
            MyAns = MsgBox("All Tickets Have Been Reported. Wanna Reprint It Only?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
            If MyAns = vbNo Then
                Me.CmdReportPrint.Visible = False
            Else
                InHoaDon(DDAN, fName, "P", RPTNO, Me.txtReportFrom.Text, Me.txtReportTo.Text, 0, Me.CmbReportCounter.Text, MySession.Domain)
            End If
            Exit Sub
        End If
        Try
            cmd.CommandText = "Drop table #RPTNO"
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try
        cmd.CommandText = "select * into #RPTNO from ActionLog where tableName='RPTLISTAL' "
        cmd.ExecuteNonQuery()
        LstRPT = ScalarToString("#RPTNO", " F1", " left(F1,7) ='" & RPTNO.Substring(0, 7) & "' " & _
            "and F1 not in (select F1 from #RPTNO where DoWhat='U' " & _
            "and left(F1,7) ='" & RPTNO.Substring(0, 7) & "') and F2='OK' order by ActionDate desc")
        If LstRPT <> "" Then
            MyAns = CInt(LstRPT.Substring(10, 2))
            If MyAns > varFrm.Day Then
                MsgBox("Last Report Was Upto " & MyAns & Format(varFrm, "MMM") & ". Plz Check Your Input", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        End If
        cmd.CommandText = String.Format("update TKT set RPTNO='{0}' where doi <'{1}' and RPTNO+AL='{2}'", RPTNO, varTo, Me.CmbReportCounter.Text)
        cmd.ExecuteNonQuery()
        cmd.CommandText = UpdateLogFile("RPTLISTAL", "C", RPTNO, "OK", "", "", "", "", "", "")
        cmd.ExecuteNonQuery()

        InHoaDon(DDAN, fName, "P", RPTNO, Me.txtReportFrom.Text, Me.txtReportTo.Text, 0, Me.CmbReportCounter.Text, MySession.Domain, Me.CmbReportCounter.Text)
    End Sub
    Private Sub CloseBC_Daily()
        Dim fName As String, MyAns As Integer, RPTNO As String, NgayChuaDongBC As Date
        Dim varFrm As Date, varTo As Date
        If Me.txtReportTo.Value.Date <> Me.txtReportFrom.Value.Date Or Me.txtReportFrom.Value.Date > Now.Date Then
            MsgBox("Invalid Date Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        varTo = Me.txtReportTo.Text & " 23:59"
        varFrm = Me.txtReportFrom.Text & " 00:01"
        If Me.GridRPTname.CurrentCell.Value = "" Then
            MsgBox("Please Select Report Name", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        NgayChuaDongBC = ScalarToDate("RCP", " top 1 FstUpdate", "fstupdate >'" & CutOverDateCloseRPT & " 23:59' and " & _
                                    " fstupdate <'" & varFrm & "' and RPTNO=''" & _
                                    " and AL='" & Me.CmbReportCounter.Text & "' and status not in ('XX','NA')" & _
                                    " and City+Location+Counter='" & MySession.City + MySession.Location + MySession.Counter & "'")
        If NgayChuaDongBC > DateSerial(2016, 5, 1) Then
            MsgBox("Sales Report For " & Format(NgayChuaDongBC, "dd-MMM-yy") & " Has not Been Closed. Plz Close It Before Running This Report", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        MyAns = MsgBox("Do You Want to Authorize Report and Print It?", MsgBoxStyle.YesNo Or MsgBoxStyle.Critical Or MsgBoxStyle.DefaultButton2, msgTitle)
        If MyAns = vbNo Then Exit Sub
        If InStr("GSA_EDU", MySession.Domain) > 0 Then
            RPTNO = "KT_" & Me.CmbReportCounter.Text & Format(varTo.Month, "00") & Format(varTo, "yy")
        Else
            RPTNO = "KT_TS" & Format(varTo.Month, "00") & Format(varTo, "yy")
        End If
        RPTNO = RPTNO & "_"
        RPTNO = RPTNO & Format(varTo, "dd")
        fName = ScalarToString("actionlog", " F1", "tablename='RPTLIST' and dowhat='C' and F1='" & RPTNO & "' and F2='OK' and f3='" & MySession.Counter & "'")

        cmd.CommandText = "update RCP set rptno=@RPTNO where RPTNO='' and FstUpdate between @varFrm and @varTo and status='OK' and al=@AL" & _
            " and City+Counter+Location=@CityCounterLoc"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@RPTNO", SqlDbType.VarChar).Value = RPTNO
        cmd.Parameters.Add("@varFrm", SqlDbType.VarChar).Value = varFrm
        cmd.Parameters.Add("@VarTo", SqlDbType.VarChar).Value = varTo
        cmd.Parameters.Add("@AL", SqlDbType.VarChar).Value = Me.CmbReportCounter.Text
        cmd.Parameters.Add("@CityCounterLoc", SqlDbType.VarChar).Value = MySession.City + MySession.Counter + myStaff.Location
        cmd.ExecuteNonQuery()

        cmd.CommandText = UpdateLogFile("RPTLIST", "C", RPTNO, "OK", MySession.Counter, varFrm.ToShortDateString, varTo.ToShortDateString, "", "", "")
        cmd.ExecuteNonQuery()
        fName = Me.GridRPTname.CurrentRow.Cells(0).Value
        InHoaDon(DDAN, fName, "P", RPTNO.Substring(3, 2), Me.txtReportFrom.Text, Me.txtReportTo.Text, 0, Me.CmbReportCounter.Text, MySession.Domain, Me.CmdReportClose.Tag)
    End Sub
    Private Sub CmdReportPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdReportPrint.Click
        If Me.CmdReportPrint.Tag = "SR" Then
            CloseBC_AL()
        Else
            CloseBC_Daily()
        End If
    End Sub

    Private Sub GenCustList4RPT(ByVal pLoaiKhach As String)
        If pLoaiKhach = "CR" Then
            LoadCmb_VAL(Me.CmbRPTcust, MyCust.List_CR)
        ElseIf pLoaiKhach = "PP" Then
            LoadCmb_VAL(Me.CmbRPTcust, MyCust.List_PP)
        ElseIf pLoaiKhach = "CS" Then
            LoadCmb_VAL(Me.CmbRPTcust, MyCust.List_CS)
        ElseIf pLoaiKhach = "CT" Then
            LoadCmb_VAL(Me.CmbRPTcust, MyCust.List_CT)
        ElseIf pLoaiKhach = "LC" Then
            LoadCmb_VAL(Me.CmbRPTcust, MyCust.List_LC)
        End If
    End Sub

    Private Sub BarBOReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarBOReport.Click
        Me.CmdReportClose.Tag = "A"
        If myStaff.City = "SGN" Then
            GoDataConvert()
        End If

        DoReport()
    End Sub
    Private Sub BarApplyPayments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarApplyPayments.Click, BarQuickInvoicing.Click, BarPaymentFollowUp.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New frmDC_Pax(b.Tag)
        f.ShowDialog()
    End Sub

    Private Sub statusConnSate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles statusConnSate.Click
        If Conn.State <> ConnectionState.Open Then Conn.Open()
    End Sub


    Private Sub BarClearPendingPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarClearPendingPayment.Click, BarClearDEB.Click, BarAddCC2FOP.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New ClearPendingPmt(b.Tag)
        f.ShowDialog()
    End Sub

    Private Sub BarVNPCT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarCreditExtensionCounter.Click, BarDEBApproval.Click, BarPPandCreditLimit.Click, _
        BarPaymentOption.Click, BarCreditApprovalCS.Click, BarApproveITP.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        If b.Name.Contains("ITP") Then
            Dim whichApp As String = InputBox("RAS Or FLX?", msgTitle, "RAS")
            If whichApp = "RAS" Then
                Dim f As New frmMISC(b.Tag)
                f.ShowDialog()
            ElseIf whichApp = "FLX" Then
                FLX_DEBControl.ShowDialog()
            Else
                Exit Sub
            End If
        Else
            Dim f As New frmMISC(b.Tag)
            f.ShowDialog()
        End If
    End Sub

    Private Sub CmbOffice_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbLocation.SelectedIndexChanged
        If Me.CmbLocation.Text.Substring(0, 2) = "BO" Then
            Me.CmbBiz.Text = "XX"
        End If
    End Sub
    Private Sub BarFODepositHandler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarRQ4DepHandlingCounter.Click, BarRQForDepositHandlingSALES.Click, BarDepositHandler.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New RQXulyDatCoc(b.Tag)
        f.ShowDialog()
    End Sub

    Private Sub BarIssueInvoiceBO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
         BarIssueInvoiceFO.Click, BarInvIssuerKTT.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New InvoicePrinting(b.Tag, "")
        f.ShowDialog()
    End Sub

    Private Sub BarInvoiceHandlerFO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarInvoiceHandlerFO.Click, BarInvHandler_KTT.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New InvHandler(b.Tag)
        f.ShowDialog()
    End Sub
    Private Sub BarReportSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarReportSales.Click
        Me.CmdReportClose.Tag = "S"
        DoReport()
    End Sub

    Private Sub BarUpdateBegBalance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        NhapSoDu.ShowDialog()
    End Sub
    Private Sub BarUNCPrinting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarUNCPrinting.Click
        UNC.ShowDialog()
    End Sub

    Private Sub BarUNCCompanyName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarUNCAcctDetails.Click, BarUNCCannedText.Click, BarUNCReprinting.Click, BarINVFromSupplier.Click, BarALINVFollowUp.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New UNC_support(b.Tag, "ACT")
        f.ShowDialog()
    End Sub
    Private Sub BarPivotUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarPivotUser.Click
        PivotUser.ShowDialog()
    End Sub
    Private Sub BarReportingPeriod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarReportingPeriod.Click
        KyBaoCao.ShowDialog()
    End Sub

    Private Sub BarADMACMupdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarADMACMupdate.Click
        ACM_ADM.ShowDialog()
    End Sub

    Private Sub BarBalancePreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarBalancePreview.Click, BarViewBalance.Click, BarBalancePreviewKT.Click, BarBalancePreviewCS.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New BalancePreview(b.Tag)
        f.ShowDialog()
    End Sub

    Private Sub BarReportCS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarReportCS.Click

        Me.CmdReportClose.Tag = "S"
        DoReport()
    End Sub
    Private Sub BarCustInfor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarCustInfor.Click
        CustInfor.ShowDialog()
    End Sub
    Private Sub BarFoxSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarFoxSettingBLC.Click, BarFoxSettingSMS.Click, BarChangePSWAndStopSales.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New FoxSetting(b.Tag)
        f.ShowDialog()
    End Sub

    Private Sub BarFoxRefund_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarFoxRefund.Click
        FoxRefund.ShowDialog()
    End Sub
    Private Sub CmbReportCounter_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbReportCounter.VisibleChanged
        Me.Label11.Visible = Me.CmbReportCounter.Visible
        If Me.CmbReportCounter.Visible Then
            If Me.CmdReportClose.Tag = "C" Then ' quay thi shortlist ten quay lai khoi lam nham bao cao
                Me.CmbReportCounter.Items.Clear()
                If MySession.Domain = "EDU" Then
                    Me.CmbReportCounter.Items.Add("XX")
                ElseIf MySession.Domain = "TVS" Then
                    Me.CmbReportCounter.Items.Add("TS")
                Else
                    LoadCmb_MSC(Me.CmbReportCounter, myStaff.TVA & " and al not in ('TS','XX')")
                End If
            Else
                LoadCmb_MSC(Me.CmbReportCounter, myStaff.TVA)
            End If
        End If
    End Sub
    Private Sub BarCS_ServiceFee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarCS_ServiceFee.Click
        CWT_ServiceFee.ShowDialog()
    End Sub
    Private Sub BarLockTheUnlocked_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarLockTheUnlocked.Click
        LockTheUnLocked.ShowDialog()
    End Sub
    Private Sub BarForEx_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarForEx.Click
        UpdateForEx.ShowDialog()
    End Sub
    Private Sub BarChargesAndFees_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarChargesAndFees.Click, BarTVComm.Click
        Dim b As ToolStripItem = CType(sender, ToolStripItem)
        Dim f As New ChargeAndFee(b.Tag)
        f.ShowDialog()
    End Sub
    Private Sub BarCustChannel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarCustChannel.Click
        CustChannelLevel.ShowDialog()
    End Sub

    Private Sub BarServiceFee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarServiceFee.Click
        ServiceFee.ShowDialog()
    End Sub

    Private Sub BarBGKeeper_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarBGKeeper.Click
        BGUpdate.Show()
    End Sub
    Private Sub GridRPTname_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridRPTname.CellContentClick
        Try
            'Me.CmbReportCounter.Visible = False
            If Me.GridRPTname.CurrentRow.Cells(0).Value.ToUpper.Substring(0, 2) = "SR" Then
                Me.CmdReportPrint.Visible = True
                Me.CmdReportPrint.Tag = "SR"
                Me.CmbRPTcust.Visible = False
                If Me.GridRPTname.CurrentRow.Cells(0).Value.ToUpper.Substring(3, 2) = "YY" Then
                    Me.CmbReportCounter.Visible = True
                End If
            ElseIf Me.GridRPTname.CurrentRow.Cells(0).Value.ToUpper.Substring(0, 2) = "AR" Then
                Me.CmdReportPrint.Visible = True
                Me.CmbRPTcust.Visible = False
                Me.CmdReportPrint.Tag = "AR"
                Me.CmbReportCounter.Visible = True
                LoadCmb_MSC(Me.CmbRptNo, "Select distinct RPTNo as val from rcp where status<>'XX' and RPTNO<>'' and city+Counter+sbu='" & _
                    MySession.City + MySession.Counter + MySession.Domain & "' order by rptno desc")

            ElseIf Me.GridRPTname.CurrentRow.Cells(0).Value.ToUpper.Substring(0, 2) = "CC" Then
                Me.CmbRPTcust.Visible = True
                GenCustList4RPT(Me.GridRPTname.CurrentRow.Cells(0).Value.ToUpper.Substring(3, 2))
            Else
                Me.CmdReportPrint.Visible = False
                Me.CmbRPTcust.Visible = False
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BarPendingXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarPendingXX.Click
        PendingXX.ShowDialog()
    End Sub

    Private Sub BarF1SUserManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarF1SUserManager.Click
        F1SUserManager.ShowDialog()
    End Sub
    Private Sub PromoCodeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarPromoCode.Click
        PromoCode.ShowDialog()
    End Sub

    Private Sub BarFrontEndLoss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarFrontEndLoss.Click
        Dim f As New GSA_MISC()
        f.ShowDialog()
    End Sub
    Private Sub txtLogInPSW_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLogInPSW.Enter
        myStaff.SICode = Me.txtLogInSIcode.Text

        LoadComboBiz(myStaff.DAccess)
        If myStaff.SICode = "SYS" Then Me.CmbCounter.Visible = True
        'If myStaff.DAccess.Contains("GSA_TVS") Or myStaff.SICode = "SYS" Then
        '    'Me.CmbBiz.Visible = True
        'Else
        '    'Me.CmbBiz.Visible = False
        'End If
        'Me.Label12.Visible = Me.CmbBiz.Visible
        Me.Label14.Visible = Me.CmbCounter.Visible
        Me.StatusVersion.Text = myStaff.DAccess & "|" & myStaff.Counter
    End Sub
    Private Sub LoadComboBiz(strDomainAccess As String)
        Dim arrDomains As String() = strDomainAccess.Split("_")
        Dim i As Integer
        CmbBiz.Items.Clear()
        For i = 0 To arrDomains.Length - 1
            CmbBiz.Items.Add(arrDomains(i))
        Next
        CmbBiz.SelectedIndex = CmbBiz.Items.Count - 1
    End Sub
    Private Sub BarCosting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarCosting.Click
        Dim frmCosting As New Costing
        frmCosting.ShowDialog()
    End Sub
    Private Sub BarUnReportedTKTs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarUnReportedTKTs.Click
        unReportedTickets.ShowDialog()
    End Sub
    Private Sub BarDataFromSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarDataFromSQL.Click
        Dim f As New Sales("RPT")
        f.Show()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles BarQuickRef.Click
        MyDoc.ShowDialog()
    End Sub

    Private Sub BarReportNonAir_Click(sender As Object, e As EventArgs) Handles BarReportNonAir.Click
        Me.CmdReportClose.Tag = "N"
        DoReport()
    End Sub
    Private Sub BasrSalesCall_Click(sender As Object, e As EventArgs) Handles BarSalesCall.Click, BarMeetingLogs.Click
        SaleLog.ShowDialog()
    End Sub
    Private Sub OverCreditOverDueToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OverCreditOverDueToolStripMenuItem.Click
        OverDue_CreditReport.ShowDialog()
    End Sub

    Private Sub BarNewsUpdate_Click(sender As Object, e As EventArgs) Handles BarNewsUpdate.Click
        NewsUpdater.ShowDialog()
    End Sub

    Private Sub BarUserMapping_Click(sender As Object, e As EventArgs) Handles BarUserMapping.Click
        B2BUserMap.ShowDialog()
    End Sub

    Private Sub BarUpdateForExQuay_Click(sender As Object, e As EventArgs) Handles BarUpdateForExQuay.Click
        UpdateForEx.ShowDialog()
    End Sub

    Private Sub BarIATAMaintain_Click(sender As Object, e As EventArgs) Handles BarIATAMaintain.Click
        Dim f As New BSP_Agent("SALES")
        f.ShowDialog()
    End Sub

    Private Sub BarReportKTT_Click(sender As Object, e As EventArgs) Handles BarReportKTT.Click
        Me.CmdReportClose.Tag = "A"
        DoReport()
    End Sub

    Private Sub BarBulkPrinting_Click(sender As Object, e As EventArgs) Handles BarBulkPrinting.Click
        BulkPrint.ShowDialog()
    End Sub

    Private Sub BarAPGInv_KTT_Click(sender As Object, e As EventArgs) Handles BarAPGInv_KTT.Click
        APG_INV.ShowDialog()
    End Sub

    Private Sub BarInvChecker_Click(sender As Object, e As EventArgs) Handles BarInvChecker.Click
        InvChecker.ShowDialog()
    End Sub

    Private Sub BarVendorInfor_Click(sender As Object, e As EventArgs) Handles BarVendorInfor.Click
        Dim f As New VendorInfor_KTT("KTT")
        f.ShowDialog()

    End Sub

    Private Sub BarUpdateVoucher_Click(sender As Object, e As EventArgs) Handles BarUpdateVoucher.Click
        DiscountManager.ShowDialog()
    End Sub

    Private Sub BarPaymenForVoucher_Click(sender As Object, e As EventArgs) Handles BarPaymenForVoucher.Click
        Payment4VCR.ShowDialog()
    End Sub

    Private Sub BarMappingSupplierVendor_Click(sender As Object, e As EventArgs) Handles BarMappingSupplierVendor.Click
        Dim f As New Vendor_SupplierMapping(0)
        f.ShowDialog()
    End Sub

    Private Sub BarUpdateSupplier_Click(sender As Object, e As EventArgs) Handles BarUpdateSupplier.Click
        Supplier.ShowDialog()
    End Sub

    Private Sub BarUNCCompanyName_Click_1(sender As Object, e As EventArgs) Handles BarUNCCompanyName.Click
        Dim f As New VendorInfor_KTT("ACC")
        f.ShowDialog()

    End Sub

    Private Sub BarGroupBooking_Click(sender As Object, e As EventArgs) Handles BarGroupBooking.Click
        GroupBooking.ShowDialog()
    End Sub
    Private Sub BarSpecimen_Click(sender As Object, e As EventArgs) Handles BarSpecimen.Click
        Dim AL As String = InputBox("Enter Airline Code ", msgTitle, "NH")
        InHoaDon(DDAN, "R12_VATInvoice_Specimen.xlt", "V", AL, Now, Now, 0, "", "", "")
    End Sub
    Private Sub CheckToUploadNewVersion()
        Dim CurrSetting As String = ScalarToString("MISC", "VAL", "Cat='APPVERSION'")
        Dim AppList As String = "RAS-12|COS-14|TSP-15"
        Dim SourceFileName As String

        SourceFileName = Application.StartupPath & "\SharedFunctions12_1.dll"
        If System.IO.File.Exists(SourceFileName) Then
            For i As Int16 = 0 To AppList.Split("|").Length - 1
                UploadFileToFtp(SourceFileName, "ftp://42.117.5.86/" & AppList.Split("|")(i).Replace("-", "") & "/", "transviet", "Abcd1234", "APP")
            Next
            Kill(SourceFileName)
        End If
        If Application.ProductVersion <> CurrSetting Then
            Shell("D:\D_disc\Exe\forFTP.bat")
            SourceFileName = Application.StartupPath & "\Ras12.exe"
            UploadFileToFtp(SourceFileName, "ftp://42.117.5.86/ras12/", "transviet", "Abcd1234", "APP")

            cmd.CommandText = "update MISC set VAL=@VAL where cat=@Cat" & _
                                "; update [42.117.5.86].ras12.dbo.MISC set VAL=@VAL where cat=@Cat"
            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("@VAL", Application.ProductVersion)
            cmd.Parameters.AddWithValue("@Cat", "APPVERSION")
            cmd.ExecuteNonQuery()
        End If
    End Sub
    Private Sub BarNonTVSGNBSPAgents_Click(sender As Object, e As EventArgs) Handles BarNonTVSGNBSPAgents.Click
        Dim f As New BSP_Agent("ACC")
        f.ShowDialog()
    End Sub

    Private Sub BarVATNoForTourDesk_Click(sender As Object, e As EventArgs) Handles BarVATNoForTourDesk.Click
        VATForTD.ShowDialog()
    End Sub
    Private Sub BarGoToTSP_Click(sender As Object, e As EventArgs) Handles BarGoToTSP.Click
        Dim fName As String = GetLastFileName_FullPath("X:\Ras2k7\TSP15\", "TSP15_*.exe")
        fName = fName & " " & myStaff.SICode & "|" & myStaff.PSW
        Try
            Shell(fName)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub BarBulkPrintingPilot_Click(sender As Object, e As EventArgs) Handles BarBulkPrintingPilot.Click
        BulkPrint_New.ShowDialog()
    End Sub

    Private Sub BarUNCPreview_Click(sender As Object, e As EventArgs) Handles BarUNCPreview.Click
        Dim f As New UNC_support("EDIT", "KTT")
        f.ShowDialog()
    End Sub

    Private Sub BarExportToTSP_Click(sender As Object, e As EventArgs) Handles BarExportToTSP.Click
        ExportToTSP.ShowDialog()
    End Sub

    
    Private Sub txtLogInSIcode_DragLeave(sender As Object, e As EventArgs) Handles txtLogInSIcode.DragLeave

    End Sub

    
   
    Private Sub BarUpdateLastBalance_Click(sender As Object, e As EventArgs) Handles BarUpdateStartBalance.Click
        frmStartBalanceList.ShowDialog()
    End Sub


    Private Sub BarCcUpdate_Click(sender As Object, e As EventArgs) Handles BarCcUpdate.Click, BarCcUpdateNonAir.Click
        frmCcList.ShowDialog()
    End Sub

    
    Private Sub BarExportToSCB_Click(sender As Object, e As EventArgs) Handles BarExportToSCB.Click
        Dim frmExport2Bank As New SCB("SCB")
        frmExport2Bank.ShowDialog()
    End Sub

    Private Sub BarExportToVCB_Click(sender As Object, e As EventArgs) Handles BarExportToVCB.Click
        Dim frmExport2Bank As New SCB("VCB")
        frmExport2Bank.ShowDialog()
    End Sub

    Private Sub barVATInvoicePrint4CWT_Click_1(sender As Object, e As EventArgs) Handles barVATInvoicePrint4CWT.Click
        frmVatInvoicePrint4Cwt.ShowDialog()
    End Sub
End Class

