Imports SharedFunctions.MySharedFunctions
Imports System.Diagnostics
Imports SharedFunctions.MySharedFunctionsWzConn
Imports System.IO
Imports System.Data.SqlClient
Imports System.Object

Public Class Costing
    Private MyCust As New objCustomer
    Private CharEntered As Boolean = False
    Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private D1Visible As String = "Accommodations_Transfer_Insurance"
    Private SVC_WzDetail As String = D1Visible & "_Meal"
    Private QrySupplier As String = "select RecID as VAL, Shortname as DIS from UNC_Company where status='OK' and cat not in ('AL','AP','TV')"
    Private FirstClickQuote As Boolean = False, OverPay As Date = "01-Jan-2000"
    Private mblnFirstLoadCompleted As Boolean
    Private mintSelectedCrdRow As Integer

    Private Structure BillBy
        Public Shared BillByBundle As String = "Bundle"
        Public Shared BillByCOD As String = "COD"
        Public Shared BillByEvent As String = "Event"
        Public Shared BillByPeriod As String = "Period"
        Public Shared BillByCRD As String = "CRD"
    End Structure
    Private Structure SVCType
        Public Shared Accommodations As String = "Accommodations"
        Public Shared Transfer As String = "Transfer"
        Public Shared Insurance As String = "Insurance"
        Public Shared Meal As String = "Meal"
    End Structure
    Private Sub LoadgridTour()
        Dim strSQL As String = "select t.*, c.VAL as Channel from DuToan_Tour t" _
                                & " left join [Cust_Detail] c on t.CustId=c.CustId" _
                                & " where c.Status='OK' and c.Cat='Channel' "

        txtSearchTcode.Text = txtSearchTcode.Text.Trim

        strSQL = strSQL & "and TCode like '%" & Me.txtSearchTcode.Text.Trim & "%'"

        Select Case cboChannel.Text
            Case "CWT"
                strSQL = strSQL & " and c.VAL='CS'"
            Case "LCL"
                strSQL = strSQL & " and c.VAL='LC'"
        End Select



        If cboCustShortName.Text <> "" Then
            strSQL = strSQL & " and CustShortName= '" & cboCustShortName.Text & "'"
        End If
        If Me.cmbStatus.Text = "CXLD" Then
            strSQL = strSQL & " and t.Status='XX'"
        ElseIf Me.cmbStatus.Text = "Finalized" Then
            strSQL = strSQL & " and t.Status='RR'"
        ElseIf Me.cmbStatus.Text = "Pending" Then
            strSQL = strSQL & " and t.Status like 'O%'"
        Else
            strSQL = strSQL & " and t.Status='O" & Me.cmbStatus.Text.Substring(0, 1) & "'"
        End If
        For j As Int16 = 0 To Me.LstCCenter.Items.Count - 1
            Me.LstCCenter.SetItemChecked(j, False)
        Next
        If Me.ChkSelectedCustOnly.Checked Then strSQL = strSQL & " and t.custID=" & Me.CmbCust.SelectedValue
        If Me.ChkPastOnly.Checked Then strSQL = strSQL & " and EDate <getdate()"
        Me.GridTour.DataSource = GetDataTable(strSQL)
        Me.GridTour.Columns("RecID").Visible = False
        Me.GridTour.Columns("CustID").Visible = False
        Me.GridTour.Columns("Pax").Width = 32
        Me.LblDeleteTour.Visible = False
        Me.LckLblFinalize.Visible = False
        LblOrderBV.Visible = False
        Me.LckLblUndoFinalize.Visible = False
        Me.LblDocSent.Visible = False
        'Me.LblPreview.Visible = False
        Me.LblQuote.Visible = False
        Me.LblSettlement.Visible = False
        Me.LblSvcCfm.Visible = False
        Me.GridTour.Columns("SDate").Width = 64
        Me.GridTour.Columns("EDate").Width = 64
        Me.LblSaveTour.Visible = False
        If InStr(Me.cmbStatus.Text, "-") > 0 Or Me.cmbStatus.Text = "Pending" Then
            Dim Back3 As Date = Now.Date.AddDays(-3)
            Dim Back7 As Date = Now.Date.AddDays(-7)
            Dim Back10 As Date = Now.Date.AddDays(-10)
            For r As Int16 = 0 To Me.GridTour.RowCount - 1
                If Me.GridTour.Item("Status", r).Value = "OC" Then Me.GridTour.Rows(r).DefaultCellStyle.ForeColor = Color.DarkGray
                If Me.GridTour.Item("Status", r).Value = "OD" Then Me.GridTour.Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                If Me.GridTour.Item("BillingBy", r).Value.ToString.Substring(0, 1) = "P" AndAlso Me.GridTour.Item("Edate", r).Value < Back3 Then
                    Me.GridTour.Rows(r).DefaultCellStyle.ForeColor = Color.Red
                ElseIf Me.GridTour.Item("BillingBy", r).Value.ToString.Substring(0, 1) = "E" AndAlso Me.GridTour.Item("Edate", r).Value < Back10 Then
                    Me.GridTour.Rows(r).DefaultCellStyle.ForeColor = Color.Red
                ElseIf Me.GridTour.Item("BillingBy", r).Value.ToString.Substring(0, 1) = "E" AndAlso Me.GridTour.Item("Edate", r).Value < Back7 Then
                    Me.GridTour.Rows(r).DefaultCellStyle.ForeColor = Color.DarkRed
                End If
            Next
        End If
    End Sub

    Private Sub Costing_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub Costing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If HasNewerVersion_R12(Application.ProductVersion) Or SysDateIsWrong(Conn) Then
            Me.Close()
            Me.Dispose()
            End
        End If
        Me.cmbStatus.Text = Me.cmbStatus.Items(0).ToString
        Me.CmbCrd.Text = "??"
        'CheckRightForALLForm(Me)
        MyCust.GenCustList()
        LoadCmb_VAL(Me.CmbCountry, "select Country as VAL, CountryName as DIS from lib.dbo.Country order by DIS")
        LoadCmb_VAL(Me.CmbCust, MyCust.List_LC)
        LoadCmb_MSC(Me.CmbCurrItem, "CURR")
        LoadCmb_MSC(Me.CmbBUAdmin, "Select Value as VAL from cwt.dbo.GO_MiscWzDate where catergory='SANOFI_BU_ADMIN' and status='OK'")
        LoadCmb_MSC(Me.CmbService, "DT_DVU")
        LoadCmb_MSC(Me.CmbPmtRQSVC, "DT_DVU")
        LoadgridTour()
        LoadCmb_VAL(Me.CmbVendor, "select RecID as VAL, Shortname as DIS from UNC_Company where status='OK'")
        Me.CmbCurrItem.Text = "VND"
        If InStr("HDI_TKL", myStaff.SICode) > 0 Then
            DisableAllLinkLabel(Me)
        End If
        reset()
        mblnFirstLoadCompleted = True
    End Sub
    Private Sub Reset()
        cboChannel.SelectedIndex = 0
        cboCustShortName.SelectedIndex = -1
        cmbStatus.SelectedIndex = 3
        txtSearchTcode.Text = ""
    End Sub
    Private Sub DisableAllLinkLabel(ByVal pRoot As Control)
        For Each Ctrl As Control In pRoot.Controls
            If Ctrl.Controls.Count > 0 Then
                DisableAllLinkLabel(Ctrl)
            ElseIf TypeOf Ctrl Is LinkLabel Then
                Ctrl.Enabled = False
                Ctrl.Visible = False
            End If
        Next
    End Sub
    Private Sub TxtPax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPax.KeyDown, TxtHTLSF.KeyDown, _
        TxtQty.KeyDown, TxtUnitCost.KeyDown, TxtVAT.KeyDown, TxtTvSfPct.KeyDown, txtMU.KeyDown, TxtTVSfAmount.KeyDown, TxtAmtAdj.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub
    Private Sub TxtPax_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtPax.KeyPress, TxtHTLSF.KeyPress, _
        TxtQty.KeyPress, TxtUnitCost.KeyPress, TxtVAT.KeyPress, TxtTvSfPct.KeyPress, txtMU.KeyPress, TxtTVSfAmount.KeyPress, TxtAmtAdj.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub
    Private Function DuplicateEventCode(pEventCode As String) As Boolean
        If Not MyCust.ShortName.Contains("SANOFI") Then Return False
        If pEventCode = "TBA" Then Return False
        Dim RecIDofGivenPR As Integer
        If pEventCode = "" Or pEventCode.ToUpper = "_EVENT CODE" Then
            Return True
        Else
            RecIDofGivenPR = ScalarToInt("Dutoan_Tour", "count(recID)", "RefNo='" & pEventCode.ToUpper.Trim _
                                         & "' and status<>'XX' and tcode <>'" & _
                                         txtTcode.Text & "' and CustShortName='" & MyCust.ShortName & "'")
            If RecIDofGivenPR > 1 Then Return True
        End If
        Return False
    End Function
    Private Function defineKeymapFromCCenter() As String
        Dim KQ As String = ""
        For i As Int16 = 0 To Me.LstCCenter.Items.Count - 1
            If Me.LstCCenter.GetItemChecked(i) Then
                KQ = KQ & "|" & Me.LstCCenter.Items(i).ToString
            End If
        Next
        If KQ.Length > 2 Then KQ = KQ.Substring(1)
        Return KQ
    End Function

    Private Sub LblCreate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCreate.LinkClicked
        Dim txtBodyAck As String = ScalarToString("MISC", "Details", "cat='EMLACK' and val='NonAir'")
        Dim strKeyMap As String = defineKeymapFromCCenter()
        Dim txtSubj As String = "Acknowledgement-" & Me.TxtBrief.Text, txtCC As String = "cwtbos.sgn@transviet.com"
        txtBodyAck = String.Format("Dear {0}, %0d{1} %0dFor further correspondence on this request, please quote", Me.CmbBooker.Text, txtBodyAck.Trim)
        If CInt(Me.TxtPax.Text) = 0 Or Me.TxtBrief.Text = "" Or Me.CmbBilling.Text = "" Or _
                Me.TxtStartDate.Value.Date > Me.txtEndDate.Value.Date Or _
                (Me.OptCWT.Checked And Me.cmbLocation.Text = "") Or _
                (Me.CmbCust.Text.Contains("SANOFI") And DuplicateEventCode(Me.txtRefCode.Text)) Then
            MsgBox("Invalid NoOfPax or Brief or Billing or StartDate or PRNo or Empty Location for CWT client", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If Me.CmbBooker.Text.ToUpper.Contains("ZPERS") And InStr("COD_CRD", Me.CmbBilling.Text) = 0 Then
            MsgBox("Illogic Booker and Billing Method", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If CmbBooker.Text = "" Then
            MsgBox("Invalid Booker!")
            Exit Sub
        End If
        If Strings.Left(Me.TxtEmail.Text, 1) = "_" Then Me.TxtEmail.Text = ""
        If Strings.Left(Me.TxtPRNo.Text, 1) = "_" Then Me.TxtPRNo.Text = ""
        If Strings.Left(Me.txtRefCode.Text, 1) = "_" Then Me.txtRefCode.Text = ""
        If Strings.Left(Me.TxtOwner.Text, 1) = "_" Then Me.TxtOwner.Text = ""
        If Strings.Left(Me.TxtIONo.Text, 1) = "_" Then Me.TxtIONo.Text = ""

        If Me.TxtStartDate.Value < Now.Date And myStaff.SupOf = "" Then
            MsgBox("Invalid Start Date Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If Me.GridGO.Visible And InvalidSIR() Then
            MsgBox("Invalid G/O Data", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If


        If OptCWT.Checked _
            AndAlso (cmbTraveler.Text.StartsWith("MR.") Or cmbTraveler.Text.StartsWith("MR ")) Then
            MsgBox("Invalid Traveller Name", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        If CmbCust.Text.Contains("SANOFI") AndAlso (txtWindowId.Text = "") Then
            MsgBox("Must input W Id for Sanofi!")
            Exit Sub
        End If

        GenTourCode(Me.TxtStartDate.Value, Me.CmbCust.Text, Me.CmbBilling.Text.Substring(0, 1))
        txtBodyAck = String.Format("{0} {1} as reference number", txtBodyAck, Me.txtTcode.Text) & "%0d" & "Best Regards"
        Dim t As SqlClient.SqlTransaction = Conn.BeginTransaction, tmpDuToanID As Integer
        cmd.Transaction = t
        Dim location As String = Me.cmbLocation.Text
        If Me.OptCWT.Checked Then location = "N/A"
        Try
            cmd.CommandText = "insert DuToan_Tour (Tcode, SDate, CustShortName, CustID, Email, Contact, Brief, Pax, BillingBy, FstUser, " & _
                "EDate, Traveller, KeyMap, Owner, PRNO, IONo, RefNo" _
                & ", BUAdmin, Dept, Location, EventDate,WindowId)" _
                & " values (@Tcode, @SDate, @CustShortName, " & _
                "@CustID, @Email, @Contact, @Brief, @Pax, @BillingBy, @FstUser,@EDate, @Traveller, @KeyMap, @Owner, @PRNO,@IONo, @RefNo,@BUAdmin," & _
                "@Dept, @Location, @EventDate, @WindowId);SELECT SCOPE_IDENTITY() AS [RecID]"
            cmd.Parameters.Clear()
            cmd.Parameters.Add("@TCode", SqlDbType.VarChar).Value = Me.txtTcode.Text
            cmd.Parameters.Add("@SDate", SqlDbType.DateTime).Value = Me.TxtStartDate.Value.Date
            cmd.Parameters.Add("@CustShortName", SqlDbType.VarChar).Value = Me.CmbCust.Text
            cmd.Parameters.Add("@CustID", SqlDbType.VarChar).Value = Me.CmbCust.SelectedValue
            cmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = Me.TxtEmail.Text
            cmd.Parameters.Add("@Contact", SqlDbType.VarChar).Value = Me.CmbBooker.Text
            cmd.Parameters.Add("@Brief", SqlDbType.NVarChar).Value = Me.TxtBrief.Text
            cmd.Parameters.Add("@Pax", SqlDbType.Int).Value = CInt(Me.TxtPax.Text)
            cmd.Parameters.Add("@BillingBy", SqlDbType.VarChar).Value = Me.CmbBilling.Text
            cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
            cmd.Parameters.Add("@EDate", SqlDbType.DateTime).Value = Me.txtEndDate.Value.Date
            cmd.Parameters.Add("@Traveller", SqlDbType.VarChar).Value = Me.cmbTraveler.Text
            cmd.Parameters.Add("@Keymap", SqlDbType.VarChar).Value = strKeyMap
            cmd.Parameters.Add("@Owner", SqlDbType.VarChar).Value = Me.TxtOwner.Text
            cmd.Parameters.Add("@PRNO", SqlDbType.VarChar).Value = Me.TxtPRNo.Text
            cmd.Parameters.Add("@IONO", SqlDbType.VarChar).Value = Me.TxtIONo.Text
            cmd.Parameters.Add("@RefNO", SqlDbType.VarChar).Value = Me.txtRefCode.Text
            cmd.Parameters.Add("@BUAdmin", SqlDbType.VarChar).Value = Me.CmbBUAdmin.Text
            cmd.Parameters.Add("@Dept", SqlDbType.VarChar).Value = Me.CmbDept.Text
            cmd.Parameters.Add("@Location", SqlDbType.VarChar).Value = location
            cmd.Parameters.Add("@EventDate", SqlDbType.DateTime).Value = Me.txtEventDate.Value.Date
            cmd.Parameters.Add("@WindowId", SqlDbType.VarChar).Value = txtWindowId.Text.Trim

            tmpDuToanID = cmd.ExecuteScalar
            t.Commit()
            Process.Start(String.Format("mailto:{0}?subject={1}&cc={2}&body={3}", Me.TxtEmail.Text, txtSubj, txtCC, txtBodyAck))
            'SendKeys.SendWait("%S")
            If Me.GridGO.Visible Then
                InsertSIR(tmpDuToanID, False)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            t.Rollback()
        End Try
        LoadgridTour()
    End Sub
    Private Sub GridTour_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridTour.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblPreview.Visible = True
        Me.LblSaveTour.Visible = True
        Me.LblOrderBV.Visible = True
        Me.LblAddBancFee.Visible = False
        Me.LblAddMerchantFee.Visible = False
        MyCust.CustID = Me.GridTour.CurrentRow.Cells("CustID").Value

        With GridTour
            If .CurrentRow.Cells("Channel").Value = "CS" Then
                OptCWT.Checked = True
            Else
                OptLCL.Checked = True
            End If
            CmbCust.SelectedIndex = CmbCust.FindStringExact(.CurrentRow.Cells("CustShortName").Value)
        End With

        If Me.GridTour.CurrentRow.Cells("BillingBy").Value = "CRD" Then Me.LblAddMerchantFee.Visible = True

        LoadGridAdj()
        Me.CmbVendorAdj.DataSource = GetDataTable("select distinct vendorID as VAL, Vendor as DIS from dutoan_item where status<>'XX' and dutoanID=" & _
                Me.GridTour.CurrentRow.Cells("RecID").Value)
        Me.CmbVendorAdj.DisplayMember = "DIS"
        Me.CmbVendorAdj.ValueMember = "VAL"
        Me.CmbCurrAdj.DataSource = GetDataTable("select distinct cCurr from dutoan_item where status<>'XX' and dutoanID=" &
                Me.GridTour.CurrentRow.Cells("RecID").Value)
        Me.CmbCurrAdj.DisplayMember = "cCurr"

        Me.LblSettlement.Visible = True
        LoadGridSVC(Me.GridTour.CurrentRow.Cells("RecID").Value)

        If GridSVC.RowCount = 0 Then
            txtPaxName.Text = GridTour.CurrentRow.Cells("Traveller").Value
        End If

        For i As Int16 = 0 To Me.GridSVC.RowCount - 1
            If ScalarToString("Lib.dbo.Supplier", "Address_CountryCode", "RecID=" & Me.GridSVC.Item("SupplierID", i).Value) <> "VN" Then
                Me.LblAddBancFee.Visible = True
                Exit For
            End If
        Next

        LoadCmb_VAL(Me.CmbPmtRQVendor, "Select distinct VendorID as VAL, Vendor as DIS from dutoan_item where dutoanID=" & Me.GridTour.CurrentRow.Cells("recID").Value & " and status <>'XX' and vendorID <>2")
        Me.GridPmtRQTC.DataSource = Nothing

        Select Me.GridTour.CurrentRow.Cells("Status").Value
            Case "RR"
                If myStaff.SupOf <> "" Then Me.LckLblUndoFinalize.Visible = True
                Me.LblAddSVC.Visible = False
                Me.LblQuote.Visible = False
                Me.LblSvcCfm.Visible = False
                Me.LblDeleteTour.Visible = False
            Case Else
                Me.LblAddSVC.Visible = True
                Me.LblQuote.Visible = True
                Me.LblSvcCfm.Visible = True
                If GridTour.CurrentRow.Cells("Status").Value <> "XX" Then
                    Me.LblDeleteTour.Visible = True
                End If
        End Select
        
        Me.LblSave.Visible = Me.LblAddSVC.Visible
        Me.LblDeleteSVC.Visible = Me.LblSave.Visible

        If InStr("OK_OC", Me.GridTour.CurrentRow.Cells("Status").Value) > 0 Then
            Me.LblDocSent.Visible = True
        End If

        If InStr("OK_OD", Me.GridTour.CurrentRow.Cells("Status").Value) > 0 Then
            Me.LckLblFinalize.Visible = True
        ElseIf InStr("RR", Me.GridTour.CurrentRow.Cells("Status").Value) > 0 Then
            Me.LckLblFinalize.Visible = False
        End If

        Me.CmbBooker.Text = Me.GridTour.CurrentRow.Cells("Contact").Value
        Me.cmbTraveler.Text = Me.GridTour.CurrentRow.Cells("Traveller").Value
        Me.TxtBrief.Text = Me.GridTour.CurrentRow.Cells("Brief").Value
        Me.TxtEmail.Text = Me.GridTour.CurrentRow.Cells("Email").Value
        Me.CmbBilling.Text = Me.GridTour.CurrentRow.Cells("BillingBy").Value
        Me.TxtPRNo.Text = Me.GridTour.CurrentRow.Cells("PRNO").Value
        Me.TxtOwner.Text = Me.GridTour.CurrentRow.Cells("Owner").Value
        Me.TxtIONo.Text = Me.GridTour.CurrentRow.Cells("IONO").Value
        Me.txtRefCode.Text = Me.GridTour.CurrentRow.Cells("RefNo").Value
        Me.CmbBUAdmin.Text = Me.GridTour.CurrentRow.Cells("BUAdmin").Value
        Me.cmbLocation.Text = Me.GridTour.CurrentRow.Cells("Location").Value
        Me.CmbDept.Text = Me.GridTour.CurrentRow.Cells("Dept").Value
        Me.TxtStartDate.Value = Me.GridTour.CurrentRow.Cells("Sdate").Value
        Me.txtEndDate.Value = Me.GridTour.CurrentRow.Cells("EDate").Value
        Me.txtEventDate.Value = Me.GridTour.CurrentRow.Cells("EventDate").Value
        Me.txtTcode.Text = Me.GridTour.CurrentRow.Cells("TCode").Value
        txtFileId.Text = GridTour.CurrentRow.Cells("FileId").Value
        txtQuotationId.Text = GridTour.CurrentRow.Cells("QuotationFile").Value
        txtWindowId.Text = GridTour.CurrentRow.Cells("WindowId").Value
        TxtPax.Text = GridTour.CurrentRow.Cells("Pax").Value

        If Me.GridTour.CurrentRow.Cells("CustShortname").Value.ToString.Contains("SANOFI") Then
            If Me.LstCCenter.Items.Count = 0 Then
                GenListCCenter(Me.GridTour.CurrentRow.Cells("CustID").Value)
            End If
            For j As Int16 = 0 To Me.LstCCenter.Items.Count - 1
                Me.LstCCenter.SetItemChecked(j, False)
            Next
            For i As Int16 = 0 To Me.GridTour.CurrentRow.Cells("KeyMap").Value.ToString.Split("|").Length - 1
                For j As Int16 = 0 To Me.LstCCenter.Items.Count - 1
                    If Me.LstCCenter.Items(j).ToString = Me.GridTour.CurrentRow.Cells("KeyMap").Value.ToString.Split("|")(i) Then
                        Me.LstCCenter.SetItemChecked(j, True)
                    End If
                Next
            Next
        Else
            Me.LstCCenter.Items.Clear()
        End If
        If Me.GridTour.CurrentRow.Cells("Channel").Value = "CS" Then
            GridGO.Visible = True
            'If Me.GridGO.Visible Or Me.GridTour.CurrentRow.Cells("Channel").Value = "CS" Then
            Dim dTbl As DataTable = GetDataTable("select FName, FValue from cwt.dbo.sir where status='OK' and rcpid=" _
                                        & Me.GridTour.CurrentRow.Cells("RecID").Value _
                                        & " and Prod='NonAir' and CustID=" & MyCust.CustID)
            For i As Int16 = 0 To dTbl.Rows.Count - 1
                For j As Int16 = 0 To Me.GridGO.RowCount - 1
                    If Me.GridGO.Item(0, j).Value = dTbl.Rows(i)("FName") Then
                        Me.GridGO.Item(1, j).Value = dTbl.Rows(i)("FValue")
                        Exit For
                    End If
                Next
            Next
        End If
        Me.GridVendorInforUpdate.DataSource = GetDataTable("select RecID, ShortName, HD from UNC_Company where cat in ('NH','KS') and status='OK' " & _
            "and recID in (select VendorID from Dutoan_Item where DutoanID=" & Me.GridTour.CurrentRow.Cells("RecID").Value & ")")
        Me.GridVendorInforUpdate.Columns(0).Visible = False
        Me.GridVendorInforUpdate.Columns(1).Width = 200
        Me.GridVendorInforUpdate.Columns(2).Width = 56
        Me.LckLblUpdateVendorAddr.Visible = False

        'If TabControl1.SelectedTab.Name = "Data4GO" Then
        '    ShowCDRs()
        'End If

        LoadGridGO(False)
    End Sub
    Private Sub LoadGridSVC(ByVal pTourID As Integer)
        Dim strDK As String = IIf(Me.ChkXXOnly.Checked, "status ='XX'", "status <>'XX'")
        Me.GridSVC.DataSource = Nothing
        FirstClickQuote = True
        Me.GridSVC.DataSource = GetDataTable("select * from DuToan_Item where  DuToanID=" & pTourID & " and " & strDK)
        'Me.GridSVC.Columns("RecID").Visible = False
        Me.GridSVC.Columns("RecID").Width = 45
        Me.GridSVC.Columns("DuToanID").Visible = False
        Me.GridSVC.Columns("CCurr").Width = 32
        Me.GridSVC.Columns("Unit").Width = 32
        Me.GridSVC.Columns("Q").Width = 25
        Me.GridSVC.Columns("CCurr").Width = 32
        Me.GridSVC.Columns("VAT").Width = 56
        Me.GridSVC.Columns("Cost").Width = 70
        Me.GridSVC.Columns("Status").Width = 32
        Me.GridSVC.Columns("Qty").Width = 32
        Me.GridSVC.Columns("VAT").Width = 32
        Me.GridSVC.Columns("PmtMethod").Width = 56
        Me.GridSVC.Columns("isVATincl").Width = 56
        Me.GridSVC.Columns("Svc_status").Width = 64
        Me.GridSVC.Columns("Cost").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridSVC.Columns("Cost").DefaultCellStyle.Format = "#,##0.0"
        If Me.ChkXXOnly.Checked Then
            Me.GridSVC.Height = 477
        Else
            Me.GridSVC.Height = 240
        End If
        Me.LblDeleteSVC.Visible = False
        Me.LblQCSF.Visible = False
        Me.LblSave.Visible = False
        Me.LblAddSF.Visible = False
        Me.GrpCost.Enabled = True
    End Sub

    Private Sub GridSVC_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridSVC.CellContentClick
        Dim sBrief As String, LineI As String
        LoadCmb_VAL(Me.CmbVendor, "select RecID as VAL, Shortname as DIS from UNC_Company where status='OK'")
        Me.CmbService.Text = Me.GridSVC.CurrentRow.Cells("Service").Value
        Me.TxtUnit.Text = Me.GridSVC.CurrentRow.Cells("Unit").Value
        txtPaxName.Text = GridSVC.CurrentRow.Cells("PaxName").Value

        Me.TxtQty.Text = Me.GridSVC.CurrentRow.Cells("Qty").Value
        Me.TxtD1.Value = Me.GridSVC.CurrentRow.Cells("SVCDate").Value
        Me.CmbVendor.Text = Me.GridSVC.CurrentRow.Cells("Vendor").Value
        Me.CmbSupplier.SelectedValue = Me.GridSVC.CurrentRow.Cells("SupplierID").Value
        Me.TxtUnitCost.Text = Me.GridSVC.CurrentRow.Cells("Cost").Value
        Me.CmbCurrItem.Text = Me.GridSVC.CurrentRow.Cells("CCurr").Value
        Me.CmbPmtMethod.Text = Me.GridSVC.CurrentRow.Cells("PmtMethod").Value
        Me.TxtVAT.Text = Me.GridSVC.CurrentRow.Cells("VAT").Value
        Me.ChkDeposit.Checked = Me.GridSVC.CurrentRow.Cells("NeedDeposit").Value
        Me.TxtMotaCost.Text = Me.GridSVC.CurrentRow.Cells("SupplierRMK").Value
        Me.CmbVendor.Text = Me.GridSVC.CurrentRow.Cells("Vendor").Value
        Me.CmbVendor.SelectedValue = Me.GridSVC.CurrentRow.Cells("VendorID").Value

        If GridSVC.CurrentRow.Cells("ZeroFeeReason").Value = "" Then
            cboZeroFeeReason.SelectedIndex = -1
        Else
            cboZeroFeeReason.SelectedIndex = cboZeroFeeReason.FindStringExact(GridSVC.CurrentRow.Cells("ZeroFeeReason").Value)
        End If


        If Me.GridSVC.CurrentRow.Cells("Cost").Value <> 0 Then
            Me.TxtHTLSF.Text = Me.GridSVC.CurrentRow.Cells("MU").Value / Me.GridSVC.CurrentRow.Cells("Cost").Value * 100
        End If

        Me.txtMU.Text = Me.GridSVC.CurrentRow.Cells("MU").Value
        Me.chkBookOnly.Checked = Me.GridSVC.CurrentRow.Cells("BookOnly").Value
        Me.ChkCostOnly.Checked = Me.GridSVC.CurrentRow.Cells("CostOnly").Value
        If Me.GridSVC.CurrentRow.Cells("isVATincl").Value = 0 Then
            Me.OptVATIncl.Checked = False
        Else
            Me.OptVATIncl.Checked = True
        End If
        sBrief = Me.GridSVC.CurrentRow.Cells("Brief").Value
        If InStr(SVC_WzDetail, Me.GridSVC.CurrentRow.Cells("Service").Value) > 0 Then
            LineI = sBrief.Split("|")(0)
            Me.cmbStype.Text = LineI.Split("_")(0)
            Me.CmbSCat.Text = LineI.Split("_")(1)
            Me.TxtD1.Text = LineI.Split("_")(2)
            Me.TxtD2.Text = LineI.Split("_")(3)
            Me.TxtT1.Text = LineI.Split("_")(2)
            Me.TxtT2.Text = LineI.Split("_")(3)
            Me.TxtQty.Text = LineI.Split("_")(4)
            Me.TxtMoTaSVC.Text = sBrief.Split("|")(1)
        Else
            Me.TxtMoTaSVC.Text = sBrief
        End If
        Me.LblDeleteSVC.Visible = True
        Me.LblSave.Visible = True
        If GridSVC.CurrentRow.Cells("Service").Value = "TransViet SVC Fee" Then
            Me.LblAddSF.Visible = False
        Else
            Me.LblAddSF.Visible = True
        End If

        If Me.GridSVC.CurrentRow.Cells("VendorID").Value <> 2 And _
            Not Me.GridSVC.CurrentRow.Cells("Service").Value.ToString.Contains("SVC") Then
            If Me.GridSVC.CurrentRow.Cells("RelatedItem").Value <> Me.GridSVC.CurrentRow.Cells("RecID").Value Or _
                Me.GridSVC.CurrentRow.Cells("RelatedItem").Value = 0 Then
                Me.LblAddSF.Visible = Me.LblAddSVC.Visible
            End If
        End If
        If Me.GridSVC.CurrentRow.Cells("VendorID").Value = 2 Then
            Me.GrpCost.Enabled = False
        Else
            Me.GrpCost.Enabled = True
        End If
        If GridSVC.CurrentRow.Cells("RelatedItem").Value = 0 Then
            lbkLinkItems.Visible = True
        Else
            lbkLinkItems.Visible = False
        End If
    End Sub
    Private Sub LblSave_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSave.LinkClicked
        Me.LblSave.Visible = False
        DeleteService()
        AddService(True)
    End Sub

    Private Sub LblAddSVC_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAddSVC.LinkClicked
        If Me.LblAddSVC.Tag = "Wait" Then Exit Sub
        Me.LblAddSVC.Tag = "Wait"

        If GridTour.CurrentRow.Cells("Channel").Value = "CS" _
            AndAlso GridTour.CurrentRow.Cells("Traveller").Value <> "ZPERSONAL" Then
            If CmbService.Text = "TransViet SVC Fee" Then
                Me.LblAddSVC.Tag = "OK"
                MsgBox("Must use function AddSF for CWT Client on Business trip")
                Exit Sub
            ElseIf TxtQty.Text <> 1 Then
                Me.LblAddSVC.Tag = "OK"
                MsgBox("Must use Qty=1 for CWT Client on Business trip")
                Exit Sub
            End If
            
        End If
        If Me.GrpCost.Enabled Then
            Dim NoOfSvcInRQ As Int16 = ScalarToInt("Dutoan_Item", "count(*)", "DutoanID =" & Me.GridTour.CurrentRow.Cells("RecID").Value & _
                                                   " and Status='OK' and service <>'TransViet  SVC Fee'")
            If NoOfSvcInRQ = 0 And Me.CmbService.Text = "TransViet  SVC Fee" Then
                MsgBox("Cannot Input TV SVC Fee Without Service", MsgBoxStyle.Critical, msgTitle)
                Me.LblAddSVC.Tag = "OK"
                Exit Sub
            End If
            AddService(False)
        Else
            If Not CheckFormatTextBox(txtTvSfQty, True, 1, 1) Or txtTvSfQty.Text < 1 Then
                Exit Sub
            ElseIf GridTour.CurrentRow.Cells("Channel").Value = "CS" _
                AndAlso GridTour.CurrentRow.Cells("Owner").Value <> "ZPERSONAL" _
                AndAlso txtTvSfQty.Text <> 1 Then
                MsgBox("Must use Qty=1 for CWT Client on Business trip")
                Me.LblAddSVC.Tag = "OK"
                Exit Sub
            ElseIf TxtTVSfAmount.Text = 0 AndAlso cboZeroFeeReason.Text = "" Then
                MsgBox("You must select Reason for Zero Service Fee")
                Me.LblAddSVC.Tag = "OK"
                Exit Sub
            End If

            cmd.CommandText = "update Dutoan_Item set RelatedItem=reciD where recid=" & Me.GridSVC.CurrentRow.Cells("RecID").Value
            cmd.ExecuteNonQuery()

            AddFee("TransViet " & Me.GrpCost.Tag, "VND", txtTvSfQty.Text, Me.TxtTVSfAmount.Text, IIf(Me.OptVATIncl.Checked, 1, 0), _
                   Me.TxtVAT.Text, IIf(Me.GrpCost.Tag = "SVC", Me.GridSVC.CurrentRow.Cells("RecID").Value, 0), 1)
            MsgBox("Service Fee Added", MsgBoxStyle.Information, msgTitle)
            LoadGridSVC(Me.GridTour.CurrentRow.Cells("RecID").Value)
        End If

        Me.TxtUnitCost.Text = "0"
        Me.TxtQty.Text = "0"
        Me.LblAddSVC.Tag = "OK"
    End Sub
    Private Function InvalidInput(pSdate As Date, pQty As Integer) As Boolean
        Dim MyAns As Int16
        If pQty = 0 Then
            MsgBox("Qty must NOT be 0")
            Return True
        End If

        If txtPaxName.Text.Trim = "" AndAlso GridTour.CurrentRow.Cells("Channel").Value = "CS" _
            AndAlso TxtOwner.Text <> "ZPERSONAL" Then
            MsgBox("Invalid PaxName")
            Return True
        End If

        If CmbService.Text = "TransViet SVC Fee" AndAlso chkBookOnly.CheckState = CheckState.Checked Then
            MsgBox("TransViet SVC Fee CAN NOT be BookOnly!")
            Return True
        End If

        If Not Me.OptVATIncl.Checked And Not Me.OptVATNotIncl.Checked Then
            MsgBox("VAT option MUST be selected!")
            Return True
        End If

        If CInt(Me.TxtUnitCost.Text) = 0 Then
            MyAns = MsgBox("Oops! Zero Cost. Wanna Correct Your Input?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, msgTitle)
            If MyAns = vbYes Then Return True
        End If
        If pSdate.Date < Me.GridTour.CurrentRow.Cells("SDate").Value Then
            MsgBox("StartDate of service must be after StartDate Of Tour!")
            Return True
        End If

        If Me.CmbService.Text = "" Or Me.CmbVendor.Text = "" Or Me.CmbCurrItem.Text = "" Then Return True
        If InStr(D1Visible, Me.CmbService.Text) > 0 Then
            If Me.TxtD1.Value.Date > Me.TxtD2.Value.Date Then Return True
        End If
        Return False
    End Function
    Private Function GetSBrief() As String
        Dim KQ As String = "", dt1 As String, dt2 As String
        dt1 = Format(Me.TxtD1.Value, "dd-MMM-yy")
        dt2 = Format(Me.TxtD2.Value, "dd-MMM-yy")
        If Me.TxtT1.Checked Then dt1 = dt1 & " " & Format(Me.TxtT1.Value, "HH:mm")
        If Me.TxtT2.Checked Then dt2 = dt2 & " " & Format(Me.TxtT2.Value, "HH:mm")
        If InStr(SVC_WzDetail, Me.CmbService.Text) = 0 Then Return Me.TxtMoTaSVC.Text
        KQ = Me.cmbStype.Text & "_" & Me.CmbSCat.Text & "_" & dt1 & "_" & dt2 & "_" & Me.TxtQty.Text & "|" & Me.TxtMoTaSVC.Text
        Return KQ
    End Function
    Private Function DefineSVCDate_time(pDate As Date, pTime As String) As Date
        Dim strKQ As String
        strKQ = Format(pDate, "dd-MMM-yy") & " " & pTime
        Return CDate(strKQ)
    End Function
    Private Sub AddService(ByVal isEdit As Boolean)
        Dim SDate As Date = DefineSVCDate_time(Me.TxtD2.Value.Date, Format(Me.TxtT2.Value, "HH:mm")), Qty As Int16 = Me.TxtQty.Text
        Dim pmtMethod As String, SupplierID As Integer, Supplier As String
        If Me.CmbService.Text = SVCType.Accommodations Then Qty = Qty * DateDiff(DateInterval.Day, Me.TxtD1.Value.Date, Me.TxtD2.Value.Date)
        If InStr(D1Visible, Me.CmbService.Text) > 0 Then SDate = DefineSVCDate_time(Me.TxtD1.Value.Date, Format(Me.TxtT1.Value, "HH:mm"))
        If InvalidInput(SDate, Qty) Then
            MsgBox("invalid Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        If Me.CmbService.Text = "Registration" And Me.CmbCust.Text.Contains("SANOFI") Then
            Dim TourIDofSameRefNo As Integer = ScalarToInt("Dutoan_Tour", "recID", "CustID =" & Me.GridTour.CurrentRow.Cells("CustID").Value & _
                                                           " and RefNo='" & Me.txtRefCode.Text.ToUpper.Trim & _
                                                           "' and status<>'XX' and tcode <>'" & Me.GridTour.CurrentRow.Cells("Tcode").Value & "'")
            If TourIDofSameRefNo > 0 Then ' da tung co 1 EventCode, likely la Reg, can tim DV cua TourID nay 
                Dim ItemIDofReg As Integer = ScalarToInt("Dutoan_Item", "RecID", "DutoanID=" & TourIDofSameRefNo & _
                                                         " and status<>'XX' and Service='Registration' ")
                If ItemIDofReg > 0 And Me.CmbService.Text = "Registration" Then
                    MsgBox("Duplicate Registration/EventCode", MsgBoxStyle.Critical, msgTitle)
                    Exit Sub
                End If
            End If
        End If

        Dim NewRecNo As Integer, isVATincl As Boolean = Me.OptVATIncl.Checked, SBrief As String, VendorID As Integer = 0
        SBrief = GetSBrief()
        If Not Me.CmbVendor.SelectedValue Is Nothing Then VendorID = Me.CmbVendor.SelectedValue
        If Me.CmbSupplier.SelectedValue Is Nothing Then
            If VendorID = 2 Then
                SupplierID = 2
                Supplier = ""
            Else
                MsgBox("You Have to Specify Supplier", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        Else
            SupplierID = Me.CmbSupplier.SelectedValue
            Supplier = Me.CmbSupplier.Text
        End If
        If VendorID = 2 Then
            pmtMethod = "PSP"
            Supplier = ""
        Else
            pmtMethod = ScalarToString("UNC_Company", "FOP", "RecID=" & VendorID)
            If String.IsNullOrEmpty(pmtMethod) Then pmtMethod = "PPD"
        End If

        cmd.CommandText = "insert DuToan_Item (Service, CCurr, Unit, Qty, Cost, Supplier, VendorID" _
            & ", Vendor, FstUser, PmtMethod, isVATIncl, " _
            & " VAT, DuToanID, Brief, SupplierRMK, MU, SVCDate, NeedDeposit, SupplierID" _
            & ", BookOnly, CostOnly,TrxCount,PaxName) Values (@Service, @CCurr, @Unit, @Qty, @Cost," _
            & "@Supplier, @VendorID, @Vendor, @FstUser, @PmtMethod, @isVATIncl, @VAT, @DuToanID" _
            & ", @Brief, @SupplierRMK, @MU, @SVCDate, " _
            & "@NeedDeposit,@SupplierID, @BookOnly, @CostOnly,@TrxCount,@PaxName)" _
            & "; SELECT SCOPE_IDENTITY() AS [RecID]"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@Service", SqlDbType.VarChar).Value = Me.CmbService.Text
        cmd.Parameters.Add("@CCurr", SqlDbType.VarChar).Value = Me.CmbCurrItem.Text
        cmd.Parameters.Add("@Unit", SqlDbType.VarChar).Value = Me.TxtUnit.Text
        cmd.Parameters.Add("@Qty", SqlDbType.Decimal).Value = Qty
        cmd.Parameters.Add("@Cost", SqlDbType.Decimal).Value = CDec(Me.TxtUnitCost.Text)
        cmd.Parameters.Add("@Supplier", SqlDbType.VarChar).Value = Supplier
        cmd.Parameters.Add("@Vendor", SqlDbType.VarChar).Value = Me.CmbVendor.Text
        cmd.Parameters.Add("@VendorID", SqlDbType.Int).Value = VendorID
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@PmtMethod", SqlDbType.VarChar).Value = pmtMethod
        cmd.Parameters.Add("@isVATIncl", SqlDbType.Bit).Value = IIf(isVATincl, 1, 0)
        cmd.Parameters.Add("@VAT", SqlDbType.Decimal).Value = CDec(Me.TxtVAT.Text)
        cmd.Parameters.Add("@DuToanID", SqlDbType.Int).Value = Me.GridTour.CurrentRow.Cells("RecID").Value
        cmd.Parameters.Add("@SupplierID", SqlDbType.Int).Value = SupplierID
        cmd.Parameters.Add("@NeedDeposit", SqlDbType.Int).Value = Me.ChkDeposit.Checked
        cmd.Parameters.Add("@Brief", SqlDbType.NVarChar).Value = SBrief
        cmd.Parameters.Add("@SupplierRMK", SqlDbType.NVarChar).Value = Me.TxtMotaCost.Text
        cmd.Parameters.Add("@MU", SqlDbType.Decimal).Value = CDec(Me.txtMU.Text)
        cmd.Parameters.Add("@SVCDate", SqlDbType.DateTime).Value = SDate
        cmd.Parameters.Add("@BookOnly", SqlDbType.Bit).Value = IIf(Me.chkBookOnly.Checked, 1, 0)
        cmd.Parameters.Add("@CostOnly", SqlDbType.Bit).Value = IIf(Me.ChkCostOnly.Checked, 1, 0)
        cmd.Parameters.Add("@TrxCount", SqlDbType.Int).Value = TxtQty.Text
        cmd.Parameters.Add("@PaxName", SqlDbType.VarChar).Value = txtPaxName.Text.Trim
        NewRecNo = cmd.ExecuteScalar
        If NewRecNo = 0 Then

        End If
        If isEdit Then
            Dim BeingEditTedRecNo As Integer = Me.GridSVC.CurrentRow.Cells("RecID").Value
            Dim OldROE As Decimal = Me.GridSVC.CurrentRow.Cells("ROE").Value
            Dim OldCurr As String = Me.GridSVC.CurrentRow.Cells("CCurr").Value
            Dim DaTra As Decimal = ScalarToDec("dutoan_pmt", "isnull(sum(VND),0)", "ItemID=" & BeingEditTedRecNo & " and status<>'XX'")
            Try
                If DaTra <> 0 Then ' gan lai ban ghi datra cho item cu thanh cho item moi
                    cmd.CommandText = " Update Dutoan_pmt set itemID=" & NewRecNo & " where ItemID=" & BeingEditTedRecNo & _
                        "; update Dutoan_Item set VNDPaid= " & DaTra & " where recid=" & NewRecNo
                    cmd.ExecuteNonQuery()
                End If
                If OldCurr = Me.CmbCurrItem.Text And Me.CmbCurrItem.Text <> "VND" And OldROE <> 0 Then
                    cmd.CommandText = String.Format("update dutoan_item set ROE={0} where RecID={1}", OldROE, NewRecNo)
                    cmd.ExecuteNonQuery()
                End If
                If Me.GridSVC.CurrentRow.Cells("recID").Value = Me.GridSVC.CurrentRow.Cells("RelatedItem").Value Then ' da co sf
                    cmd.CommandText = String.Format("update dutoan_item set relatedItem={0} where relatedItem={1} and vendorID=2; " & _
                        "update dutoan_item set relatedItem={0} where RecID={0}", NewRecNo, BeingEditTedRecNo)
                    cmd.ExecuteNonQuery()
                End If
            Catch ex As Exception
                GoTo ErrHandler
            End Try
        End If
        Me.CmbVendor.Visible = True
        LoadGridSVC(Me.GridTour.CurrentRow.Cells("RecID").Value)
        Me.CmbCurrItem.Text = "VND"
        MsgBox("Service Item Added", MsgBoxStyle.Information, msgTitle)

        If Me.CmbService.Text = "Visa" Or Me.CmbService.Text = "Registration" Then
            If MsgBox("Wanna Order a Messgenger to Collect Travel Docs?", MsgBoxStyle.Question Or vbYesNo, msgTitle) = vbYes Then
                Book_a_MSGR(myStaff.SICode, myStaff.PSW, "N/A", Me.GridTour.CurrentRow.Cells("Tcode").Value, 0)
            End If
        End If
        Exit Sub
ErrHandler:
        MsgBox(NewRecNo.ToString & ": Error Occurs During Saving Changes. " & cmd.CommandText & ". Plz take a screenshot and email VTH", MsgBoxStyle.Information, msgTitle)
    End Sub


    Private Sub LblDeleteSVC_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDeleteSVC.LinkClicked
        Dim myAns As Integer = MsgBox("Clicking Delete by Mistake?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
        If myAns = vbYes Then Exit Sub
        DeleteService()
        LoadGridSVC(Me.GridTour.CurrentRow.Cells("RecID").Value)
    End Sub
    Private Sub DeleteService()
        If Me.GridTour.CurrentRow.Cells("Status").Value = "RR" Then Exit Sub
        Dim ExistingPmtID As Integer = ScalarToInt("Dutoan_pmt", "RecID", "ItemID=" & Me.GridSVC.CurrentRow.Cells("recID").Value & " and status='OK'")
        If Me.GridSVC.CurrentRow.Cells("VNDPaid").Value > 0 Or ExistingPmtID > 0 Then
            If myStaff.SICode <> "PNT" Then
                MsgBox("This Item Has Been Paid. Ask Supervisor to Delete.", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            Else
                MsgBox("This Item Has Been Paid. Be Careful.", MsgBoxStyle.Critical, msgTitle)
            End If
        End If
        cmd.CommandText = ChangeStatus_ByID("DuToan_Item", "XX", Me.GridSVC.CurrentRow.Cells("recID").Value)
        cmd.ExecuteNonQuery()
        If Me.GridSVC.CurrentRow.Cells("RelatedItem").Value <> 0 Then ' da co SF
            If Me.GridSVC.CurrentRow.Cells("RelatedItem").Value = Me.GridSVC.CurrentRow.Cells("RecID").Value Then ' ban ghi bi huy la chinh, se huy ban ghi sf
                cmd.CommandText = ChangeStatus_ByDK("DuToan_Item", "XX", String.Format("relatedItem={0} and status='OK'", Me.GridSVC.CurrentRow.Cells("recID").Value))
            Else ' ban ghi bi huy la sf, se danh dau ban ghi chinh thanh chua co sf
                cmd.CommandText = "update DuToan_Item set relatedItem=0 where status='OK' and recid=" & Me.GridSVC.CurrentRow.Cells("RelatedItem").Value
            End If
            cmd.ExecuteNonQuery()
        End If
    End Sub
    Private Sub LblDeleteTour_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDeleteTour.LinkClicked
        If myStaff.SupOf = "" Then Exit Sub
        Dim myAns As Integer = MsgBox("Clicking Delete by Mistake?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
        Dim AirID As Integer
        If myAns = vbYes Then Exit Sub
        If Me.GridTour.CurrentRow.Cells("custID").Value < 0 Then
            cmd.CommandText = String.Format("Delete from dutoan_tour where recid={0}; delete from dutoan_item where dutoanID={0}", Me.GridTour.CurrentRow.Cells("recID").Value)
            cmd.ExecuteNonQuery()
            GoTo ResumeHere
        End If
        If Me.GridTour.CurrentRow.Cells("billingBy").Value = BillBy.BillByBundle Or _
                Me.GridTour.CurrentRow.Cells("billingBy").Value = BillBy.BillByEvent Then
            AirID = ScalarToInt("FOP", "RecID", "status<>'XX' and Document='" & Me.GridTour.CurrentRow.Cells("TCode").Value & "'")
            If AirID > 0 Then
                MsgBox("This TourCode Has Air Part. Plz Ask Air Team To Change FOP Before Continue", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        End If
        Dim LyDo As String = InputBox("Plz Enter Valid Reason for Deleting", msgTitle, "By Customer RQST")
        If LyDo = "" Then Exit Sub
        cmd.CommandText = "Update DuToan_Tour set status='XX', RMK=RMK+'|" & LyDo & "' where recid=" & _
            Me.GridTour.CurrentRow.Cells("recID").Value & _
            ";" & ChangeStatus_ByDK("DuToan_Item", "XX", String.Format("DuToanID={0}", Me.GridTour.CurrentRow.Cells("recID").Value))
        cmd.ExecuteNonQuery()
ResumeHere:
        LoadgridTour()
    End Sub
    Private Sub GridSVC_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridSVC.CellContentDoubleClick
        Me.GridSVC.Width = 773
        Me.GridSVC.BringToFront()
    End Sub
    Private Sub UpdateROE_ItemPricing(QuoteOrSvcCfm As String)
        Dim strDKWhatRow As String = String.Format("DuToanID={0} and status='OK'", Me.GridTour.CurrentRow.Cells("RecID").Value)
        Dim strDKZeroROE As String = String.Format("{0} and ROE=0", strDKWhatRow)
        Dim MyAns As Integer, tmpROE As Decimal
        cmd.CommandText = String.Format("update dutoan_item set roe=1 where {0} and cCurr='VND'", strDKWhatRow)
        cmd.ExecuteNonQuery()
        Dim tmpCurr As String = ScalarToString("DuToan_Item", "top 1 CCurr", strDKZeroROE)
        If String.IsNullOrEmpty(tmpCurr) Then
            If QuoteOrSvcCfm = "Q" Or myStaff.SupOf = "" Then Exit Sub
            MyAns = MsgBox("ROE Has Been Updated. Wanna Edit?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
            If MyAns = vbNo Then
                Exit Sub
            Else
                cmd.CommandText = String.Format("update dutoan_item set roe=0 where {0} and cCurr<>'VND'", strDKWhatRow)
                cmd.ExecuteNonQuery()
            End If
        End If
        Do
            tmpCurr = ScalarToString("DuToan_Item", "top 1 CCurr", strDKZeroROE)
            If String.IsNullOrEmpty(tmpCurr) Then Exit Do
            tmpROE = ForEX_12(Now, tmpCurr, "BSR", "TS").Amount
            If tmpROE = 0 Then
                MsgBox("ROE for " & tmpCurr & " not Found. Ask Accounting Dept to Update", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
            cmd.CommandText = String.Format("update dutoan_item set roe={0} where {1} and cCurr='{2}'", tmpROE, strDKWhatRow, tmpCurr)
            cmd.ExecuteNonQuery()
        Loop
    End Sub
    Private Sub CmbService_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbService.SelectedIndexChanged
        Me.TxtHTLSF.Text = 0
        Me.LblD2.Text = "Svc Date"
        Me.TxtUnit.Text = ""
        Me.TxtT1.Checked = False
        Me.TxtT2.Checked = False
        Me.TxtT1.Text = "00:00"
        Me.TxtT2.Text = "00:00"
        If Me.CmbService.Text = "Miscellaneous" Then
            Me.CmbVendor.DropDownStyle = ComboBoxStyle.DropDown
        Else
            Me.CmbVendor.DropDownStyle = ComboBoxStyle.DropDownList
        End If
        If InStr(SVC_WzDetail, Me.CmbService.Text) = 0 Then
            Me.LblSCat.Visible = False
            'LoadCmb_VAL(Me.CmbVendor, QrySupplier)
            If Me.CmbService.Text = "Air Tickets" Then
                LoadCmb_VAL(Me.CmbVendor, String.Format("{0} and cat='{1}'", QrySupplier, "AR"))
                Me.LblD1.Visible = False
                Me.TxtD1.Visible = False
                Me.LblD2.Visible = True
                Me.TxtD2.Visible = True
                Me.TxtUnit.Text = "Pax"
            End If
        Else
            Me.LblSCat.Visible = True
            Me.LblD1.Visible = True
            Me.TxtD1.Visible = True
            Me.LblD2.Visible = True
            Me.TxtD2.Visible = True
            If Me.CmbService.Text = SVCType.Accommodations Then
                Me.LblSType.Text = "RoomType"
                Me.LblSCat.Text = "RoomCat"
                LblD1.Text = "ChkIn"
                LblD2.Text = "ChkOut"
                Me.TxtHTLSF.Text = 5
                Me.TxtUnit.Text = "R/N"
                'LoadCmb_VAL(Me.CmbVendor, String.Format("{0} and cat='{1}'", QrySupplier, "KS"))
            ElseIf Me.CmbService.Text = SVCType.Transfer Then
                Me.LblSType.Text = "CarType"
                Me.LblSCat.Text = "CarCat"
                LblD1.Text = "PickUp"
                LblD2.Text = "Drop"
                'LoadCmb_VAL(Me.CmbVendor, String.Format("{0} and cat='{1}'", QrySupplier, "XE"))
                Me.TxtUnit.Text = "Car"
            ElseIf Me.CmbService.Text = SVCType.Meal Then
                Me.LblSType.Text = "Cusine"
                Me.LblSCat.Text = "MenuType"
                LblD1.Visible = False
                LblD2.Text = "Time"
                Me.TxtD1.Visible = False
                Me.TxtUnit.Text = "Pax"
                Me.TxtUnit.Visible = True
                'LoadCmb_VAL(Me.CmbVendor, String.Format("{0} and cat='{1}'", QrySupplier, "NH"))
            ElseIf Me.CmbService.Text = SVCType.Insurance Then
                Me.LblSType.Text = "Type"
                Me.LblSCat.Text = "Cat"
                LblD1.Visible = True
                LblD1.Text = "From"
                LblD2.Text = "Thru"
                Me.TxtD1.Visible = True
                Me.TxtUnit.Text = "Person"
            End If
            Me.cmbStype.DataSource = GetDataTable("Select VAL from Misc where cat='STYPE' and val1='" & Me.CmbService.Text & "'")
            Me.cmbStype.DisplayMember = "VAL"
            Me.CmbSCat.DataSource = GetDataTable("Select VAL from Misc where cat='SCAT' and val1='" & Me.CmbService.Text & "'")
            Me.CmbSCat.DisplayMember = "VAL"
        End If
        Me.CmbVendor.SelectedValue = 2
    End Sub

    Private Sub LblFilter_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblFilter.LinkClicked
        LoadCmb_VAL(Me.CmbVendor, String.Format("{0} and shortname like '%{1}%'", QrySupplier, Me.TxtFilter.Text))
    End Sub
    Private Sub GenTourCode(ByVal pStart As Date, ByVal pCustName As String, ByVal pBillType As String)
        Dim LastTC As String, strThang As String = pStart.Month.ToString.Trim
        Dim LastChar As String
        If strThang.Length = 2 Then strThang = Chr(CInt(strThang) + 55)
        Dim tmpTC As String = pCustName & (pStart.Year - 2010).ToString.Trim & strThang & pStart.Day.ToString.Trim & pBillType
        cmd.CommandText = "select top 1 TCode from DuToan_Tour where TCode like '" & tmpTC & "%' order by RecID desc"
        LastTC = cmd.ExecuteScalar
        If String.IsNullOrEmpty(LastTC) Then
            LastChar = "A"
        ElseIf Strings.Right(LastTC, 3).Contains(".") Then
            LastChar = Strings.Right(LastTC, 2)
            LastChar = "." & Format(CInt(LastChar) + 1, "00")
        ElseIf Strings.Right(LastTC, 1) = "Z" Then
            LastChar = "0"
        ElseIf Strings.Right(LastTC, 1) = "9" Then
            LastChar = ".10"
        Else
            LastChar = Strings.Right(LastTC, 1)
            LastChar = Chr(Asc(LastChar) + 1)
        End If
        tmpTC = tmpTC & LastChar
        Me.txtTcode.Text = tmpTC
    End Sub
    Private Sub GenListCCenter(pCustID As Integer)
        Dim pSQL As String = "select distinct Traveler as VAL from employeeid where custid=" & pCustID & " and status='OK'"
        Dim tbl As DataTable = GetDataTable(pSQL)
        For i As Int16 = 0 To tbl.Rows.Count - 1
            Me.LstCCenter.Items.Add(tbl.Rows(i)("VAL"))
        Next
    End Sub
    Private Sub CmbCust_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCust.SelectedIndexChanged
        'Private Sub CmbCust_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCust.LostFocus _
        ', CmbCust.SelectedIndexChanged
        If Not mblnFirstLoadCompleted Then
            Exit Sub
        End If
        MyCust.CustID = Me.CmbCust.SelectedValue
        Dim NeedCheckTravelerName As Int16
        If MyCust.ShortName.Contains("SANOFI") Then
            'LoadCmb_MSC(Me.CmbCCenter, "select distinct Traveler as VAL from employeeid where custid=" & MyCust.CustID & " and status='OK'")
            GenListCCenter(MyCust.CustID)
            LoadCmb_MSC(Me.CmbDept, "SELECT [value] as VAL FROM [CWT].[dbo].[GO_MiscWzDate] where catergory='SANOFI_DEPT' and status='OK'")
            Me.CmbBooker.DropDownStyle = ComboBoxStyle.DropDown
            Me.TxtPRNo.Text = "_PR No."
            Me.txtRefCode.Text = "_EVENT Code"
            Me.TxtOwner.Text = "_EVENT OWNER"
        ElseIf MyCust.ShortName = "LOR VN" Then
            'LoadCmb_MSC(Me.CmbCCenter, "select '' as VAL ")
            Me.LstCCenter.Items.Clear()
            LoadCmb_MSC(Me.CmbDept, "SELECT '' as VAL ")

            Me.TxtPRNo.Text = "_GL"
            Me.txtRefCode.Text = "_TER"
            Me.TxtOwner.Text = "_BUDGET"
        Else
            Me.CmbBooker.DropDownStyle = ComboBoxStyle.DropDownList
            Me.LstCCenter.Items.Clear()
        End If
        NeedCheckTravelerName = ScalarToInt("cwt.dbo.go_companyinfo1", "Empl4NonAir", "Status='OK' and custid=" & MyCust.CustID)
        Me.cmbTraveler.DataSource = Nothing
        If NeedCheckTravelerName = 0 Then
            Me.cmbTraveler.DropDownStyle = ComboBoxStyle.DropDown
        Else
            Me.cmbTraveler.DropDownStyle = ComboBoxStyle.DropDownList
            LoadCmb_MSC(Me.cmbTraveler, "SELECT Traveler as VAL FROM cwt.dbo.GO_EmployeeID where status<>'XX' and CustID=" & MyCust.CustID)
        End If
        Try
            LoadCmb_MSC(Me.CmbBooker, "Select BookerName as VAL from [42.117.5.86].ft.dbo.Cwt_Bookers where Status='OK' and CustId=" & Me.CmbCust.SelectedValue)
        Catch ex As Exception
            LoadCmb_MSC(Me.CmbBooker, "select distinct fValue as VAL from SIR where fName+status='BOOKEROK' and custID=" & Me.CmbCust.SelectedValue)
        End Try
        LoadGridGO(True)
        'LoadgridTour()
    End Sub
    Private Function GetQryMCE2PSP(ppTcode As String) As String
        Return "; update FOP set fop='" & MyCust.DelayType & "', RMK='TT.INV.BDL' where status='OK' and fop='MCE' and document='" & ppTcode & "'"
    End Function
    Private Sub LblFinalize_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblFinalize.LinkClicked
        If HasNewerVersion_R12(Application.ProductVersion) Or SysDateIsWrong(Conn) Then
            Me.Close()
            Me.Dispose()
            End
        End If
        Dim MyAns As Integer = MsgBox("Did You Click This By Accident", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
        If MyAns = vbYes Then Exit Sub
        Dim DuToanID As String = ScalarToString("FOP", "RMK", " Document='" & Me.GridTour.CurrentRow.Cells("TCode").Value & "' and status='OK' and fop <>'MCE' and rmk <>'MCE'")
        If Not String.IsNullOrEmpty(DuToanID) Then
            MsgBox("This Tour Code Has Been Finalised", MsgBoxStyle.Information, msgTitle)
            Exit Sub
        End If
        Dim tmpROE As Decimal, tmpPrice As Decimal, strSQL As String
        Dim i As Integer
        Dim tblItems As DataTable = GetDataTable("Select * from dutoan_Item where Status='OK'" _
                                                 & " and DutoanId=" & Me.GridTour.CurrentRow.Cells("recID").Value)

        strSQL = "DuToanID=" & Me.GridTour.CurrentRow.Cells("recID").Value & " and status='OK' "

        tmpROE = ScalarToDec("Dutoan_Item", "top 1 ROE", strSQL & " order by ROE")
        strSQL = ChangeStatus_ByID("Dutoan_Tour", "RR", Me.GridTour.CurrentRow.Cells("recID").Value)
        If tblItems.Rows.Count = 0 Then
            MyAns = MsgBox("This Tourcode is Empty. Wanna Recheck?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, msgTitle)
            If MyAns = vbYes Then Exit Sub
            cmd.CommandText = strSQL
            If Me.GridTour.CurrentRow.Cells("BillingBy").Value = BillBy.BillByBundle Then
                cmd.CommandText = cmd.CommandText & GetQryMCE2PSP(Me.GridTour.CurrentRow.Cells("TCode").Value)
            End If
            GoTo UpdateDbHere
        ElseIf GridTour.CurrentRow.Cells("Channel").Value = "CS" _
                AndAlso GridTour.CurrentRow.Cells("Contact").Value <> "ZPERSONAL" Then
            For i = 0 To tblItems.Rows.Count - 1
                If tblItems.Rows(i)("RelatedItem") = 0 _
                    AndAlso Not ("Merchant Fee").Contains(tblItems.Rows(i)("Service")) _
                    AndAlso tblItems.Rows(i)("ZeroFeeReasion") = "" Then
                    MsgBox("Must link SVC with FEE for item" & tblItems.Rows(i)("RecId"))
                    Exit Sub
                End If
            Next
        End If

        If tmpROE = 0 Then
            MsgBox("You Have Finished Pricing First", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        Me.LckLblFinalize.Visible = False

        tmpPrice = ItemPriceToTotalPrice(Me.GridTour.CurrentRow.Cells("recID").Value, "TTL")

        If Me.GridTour.CurrentRow.Cells("BillingBy").Value <> BillBy.BillByEvent Then
            strSQL = TaoBanGhiRasEvent(strSQL, Me.GridTour.CurrentRow.Cells("recID").Value, Me.GridTour.CurrentRow.Cells("TCode").Value, _
                "VND", tmpPrice, Me.GridTour.CurrentRow.Cells("BillingBy").Value, Me.GridTour.CurrentRow.Cells("Contact").Value)
        End If
        cmd.CommandText = strSQL
UpdateDbHere:
        Dim t As SqlClient.SqlTransaction = Conn.BeginTransaction
        cmd.Transaction = t
        Try
            cmd.ExecuteNonQuery()
            t.Commit()
        Catch ex As Exception
            t.Rollback()
            MsgBox("Error Finalizing This Tour", MsgBoxStyle.Critical, msgTitle)
        End Try
        LoadgridTour()
    End Sub
    Private Function ItemPriceToTotalPrice(ByVal pTourID As Integer, ByVal pWhat As String) As Decimal
        Dim strSQL As String = String.Format("select * from DuToan_Item where DutoanID={0} and status='OK' and TrvlPay=0 and BookOnly=0 and CostOnly=0", pTourID)
        Dim dTable As DataTable = GetDataTable(strSQL)
        Dim Cost As Decimal = 0, Gia As Decimal, TVCharge As Decimal
        For i As Int16 = 0 To dTable.Rows.Count - 1
            Gia = dTable.Rows(i)("Cost") + dTable.Rows(i)("MU")
            If Not dTable.Rows(i)("isVATIncl") Then
                Gia = Gia + Gia * dTable.Rows(i)("VAT") / 100
            End If
            If InStr(dTable.Rows(i)("Service"), "SVC Fee") > 0 Then
                TVCharge = TVCharge + dTable.Rows(i)("qty") * dTable.Rows(i)("ROE") * Gia
            Else
                Cost = Cost + dTable.Rows(i)("qty") * dTable.Rows(i)("ROE") * Gia
            End If
        Next
        If pWhat = "TTL" Then
            Return TVCharge + Cost
        ElseIf pWhat = "TVC" Then
            Return TVCharge
        ElseIf pWhat = "COST" Then
            Return Cost
        End If
    End Function
    Private Function TaoBanGhiRasEvent(ByVal pSQL As String, ByVal pTourID As Integer, ByVal pTCode As String, ByVal pCurr As String, ByVal pPrice As Decimal, ByVal pBilling As String, pBooker As String) As String
        Dim KQ As String = pSQL, RCPNo As String, LocalRCPID As Integer, tmpROE As Decimal = 1, ROEID As Integer
        Dim TKNO As String = GenPseudoTKT("EVT", "TV")

        Dim Fare As Decimal = ItemPriceToTotalPrice(pTourID, "COST")
        Dim TVCharge As Decimal = ItemPriceToTotalPrice(pTourID, "TVC")
        Dim RasFOP As String = ""
        If pBilling = BillBy.BillByCOD Then
            RasFOP = "DEB"
        ElseIf pBilling = BillBy.BillByCRD Then
            RasFOP = "CRD"
        Else
            RasFOP = MyCust.DelayType
        End If
        RCPNo = GenRCPNo(MySession.TRXCode, "0")
        ROEID = ForEX_12(Now.Date, "USD", "RECID", "YY").Id
        If pCurr <> "VND" Then
            tmpROE = ScalarToDec("forEx", "BSR", "Recid=" & ROEID)
        End If
        If RCPNo <> "" Then
            LocalRCPID = ScalarToInt("RCP", "RecID", "RCPNO='" & RCPNo & "'")
            KQ = KQ & "; Update RCP set CustID=" & MyCust.CustID & ", DeliveryStatus='" & pTCode & _
                "', FstUser='AUT', CustType='" & MyCust.CustType & "', Counter='N-A', status='NA', SRV='S'," & _
                " DOS='" & Format(Now, "dd-MMM-yy") & "', Stock='01', CustshortName='" & MyCust.ShortName & _
                 "', PrintedCustName='" & MyCust.FullName & "', PrintedCustAddrr='" & MyCust.Addr & "', PrintedTaxCode='" & _
                 MyCust.taxCode & "', Currency='" & pCurr & "', ROE=" & tmpROE & ", City='" & MySession.City & _
                 "', Location='TVH', TTLDue=" & pPrice & ", ROEID=" & ROEID & ", CA='" & pBooker.Replace("--", "") & "' where recid=" & LocalRCPID
            Dim Amt As Decimal = pPrice
            If MyCust.AdhType.Trim <> "" Then
                Amt = Fare
                KQ = KQ & "; insert fop (fop, currency, roe, amount, RCPID, RCPNO, document, RMK, customerID, FstUser) values ('" & _
                    MyCust.AdhType & "','" & pCurr & "'," & tmpROE & "," & TVCharge & "," & LocalRCPID & _
                    ",'" & RCPNo & "','" & pTCode & "','" & pTourID.ToString & "'," & MyCust.CustID & ",'" & myStaff.SICode & "')"
            End If
            KQ = KQ & "; insert fop (fop, currency, roe, amount, RCPID, RCPNO, document, RMK, customerID, Status, FstUser) values ('" & _
                    RasFOP & "','" & pCurr & "'," & tmpROE & "," & Amt & "," & LocalRCPID & _
                    ",'" & RCPNo & "','" & pTCode & "','" & pTourID.ToString & "'," & MyCust.CustID & ",'" & _
                    IIf(RasFOP = "CRD", "QQ", "OK") & "','" & myStaff.SICode & "')"
            If pBilling = BillBy.BillByBundle Then
                KQ = KQ & GetQryMCE2PSP(pTCode)
            End If
            Return KQ
        Else
            Return "Err. Unable to crate RCP"
        End If
    End Function
    Private Sub TxtUnitCost_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitCost.LostFocus, txtMU.LostFocus, TxtTVSfAmount.LostFocus
        Dim aa As Decimal = CDec(Me.TxtUnitCost.Text)
        Me.TxtUnitCost.Text = Format(aa, "#,##0.0")
    End Sub
    Private Sub TxtUnitCost_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtUnitCost.TextChanged, txtMU.TextChanged, TxtTVSfAmount.TextChanged
        Me.TxtTvSfPct.Text = 0
    End Sub
    Private Sub TxtSFPCT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTvSfPct.LostFocus
        Dim tmpSF As Decimal
        With GridSVC.CurrentRow
            tmpSF = (.Cells("Cost").Value + .Cells("MU").Value) * .Cells("Qty").Value / txtTvSfQty.Text
        End With

        'phai tim related item roi tinh % theo do
        If Not Me.GridSVC.CurrentRow.Cells("isVATIncl").Value Then
            tmpSF = tmpSF + tmpSF * Me.GridSVC.CurrentRow.Cells("VAT").Value / 100
        End If
        Me.TxtTVSfAmount.Text = tmpSF * CDec(Me.TxtTvSfPct.Text) / 100
    End Sub
    Private Sub LblQuote_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblQuote.LinkClicked
        If updateDL_CDM_Quote(Me.GridTour.CurrentRow.Cells("recID").Value, "QuoteValid", "Quotation Validity") Then
            UpdateROE_ItemPricing("Q")
            Dim myPath As String = Application.StartupPath
            Dim ItemList As String = ""
            If FirstClickQuote Then
                For i As Int16 = 0 To Me.GridSVC.RowCount - 1
                    Me.GridSVC.Item("Q", i).Value = Not Me.GridSVC.Item("Q", i).Value
                Next
                FirstClickQuote = False
                Dim MyAns As Integer = MsgBox("Are You OK To Make Quotation for Selected Items?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
                If MyAns = vbNo Then Exit Sub
            End If
            For i As Int16 = 0 To Me.GridSVC.RowCount - 1
                If Me.GridSVC.Item("Q", i).Value Then
                    ItemList = ItemList & "," & Me.GridSVC.Item("RecID", i).Value
                End If
            Next
            If ItemList.Length > 2 Then
                ItemList = "(" & ItemList.Substring(1) & ")"
                cmd.CommandText = "update dutoan_Item set q=-1 where recid in " & ItemList
                cmd.ExecuteNonQuery()
            End If
            If myStaff.Counter = "ALL" Then
                If Me.GridTour.CurrentRow.Cells("CustShortName").Value = "ROCHE" Then
                    InHoaDon(myPath, "Quotation_Roche.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, ItemList, "", "")
                ElseIf Me.GridTour.CurrentRow.Cells("CustShortName").Value = "MAST" Then
                    InHoaDon(myPath, "Quotation_MAST.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, ItemList, "", "")
                Else
                    InHoaDon(myPath, "Quotation.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, ItemList, "", "")
                End If

            Else
                If Me.GridTour.CurrentRow.Cells("CustShortName").Value = "ROCHE" Then
                    InHoaDon(myPath, "Quotation_Roche.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, ItemList, "", "E")
                ElseIf Me.GridTour.CurrentRow.Cells("CustShortName").Value = "MAST" Then
                    InHoaDon(myPath, "Quotation_MAST.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, ItemList, "", "E")
                Else
                    InHoaDon(myPath, "Quotation.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, ItemList, "", "E")
                End If
            End If
        End If
    End Sub

    Private Sub TxtStart_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtStartDate.LostFocus, txtEventDate.LostFocus
        GenTourCode(Me.TxtStartDate.Value, Me.CmbCust.Text, "")
    End Sub

    Private Sub GridTour_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridTour.CellContentDoubleClick
        'Me.TabControl1.SelectTab("TabPage2")
    End Sub


    Private Sub LblSettlement_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSettlement.LinkClicked
        Dim myPath As String = Application.StartupPath
        InHoaDon(myPath, "QuyetToanTour.xlt", "V", "", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "")
    End Sub
    Private Function updateDL_CDM_Quote(ByVal pTourID As Integer, ByVal pWhat As String, ByVal pMsg As String) As Boolean
        Dim fstUpdate As Date = ScalarToDate("DuToan_Tour", "FstUpdate", "RecID=" & pTourID)
        Dim tmpDeadLine As Date = ScalarToDate("DuToan_Tour", pWhat, "RecID=" & pTourID)
        If fstUpdate <> tmpDeadLine Then Return True
        fstUpdate = Now.Date.AddDays(4)
        tmpDeadLine = InputBox("Please Input " & pMsg, msgTitle, Format(fstUpdate, "dd-MMM-yy"))
        If tmpDeadLine < Now.Date Then Return False
        cmd.CommandText = "update DuToan_Tour set " & pWhat & "='" & tmpDeadLine & "' where recid=" & pTourID & "; " & UpdateLogFile("dutoan_Tour", pWhat, pTourID, tmpDeadLine, "", "", "")
        cmd.ExecuteNonQuery()
        If pWhat = "DLCFM" And Me.GridTour.CurrentRow.Cells("Status").Value = "OK" Then
            cmd.CommandText = "update DuToan_Tour set status='OC' where RecID=" & pTourID
            cmd.ExecuteNonQuery()
        End If
        Return True
    End Function
    Private Sub LblSvcCfm_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSvcCfm.LinkClicked
        Dim myAns As Int16 = MsgBox("Be Carefull. Click OK ONLY if you REALLY want to send confirmation to client" & vbCrLf & "Otherwise, click Cancel and hit Preview", MsgBoxStyle.Critical Or MsgBoxStyle.OkCancel, msgTitle)
        If myAns = vbCancel Then Exit Sub
        If Me.GridTour.CurrentRow.Cells("Traveller").Value = "" Then Exit Sub
        If updateDL_CDM_Quote(Me.GridTour.CurrentRow.Cells("recID").Value, "DLCFM", "Dead Line to Confirm") Then
            UpdateROE_ItemPricing("S")
            Dim myPath As String = Application.StartupPath
            If myStaff.Counter = "ALL" Then
                If Me.GridTour.CurrentRow.Cells("CustShortName").Value = "ROCHE" Then
                    InHoaDon(myPath, "Quotation_Roche.xlt", "V", "S", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "F")
                ElseIf Me.GridTour.CurrentRow.Cells("CustShortName").Value = "MAST" Then
                    InHoaDon(myPath, "Quotation_MAST.xlt", "V", "S", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "F")
                Else
                    InHoaDon(myPath, "Quotation.xlt", "V", "S", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "F")
                End If

            Else
                If Me.GridTour.CurrentRow.Cells("CustShortName").Value = "ROCHE" Then
                    InHoaDon(myPath, "Quotation_Roche.xlt", "V", "S", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "E")
                ElseIf Me.GridTour.CurrentRow.Cells("CustShortName").Value = "MAST" Then
                    InHoaDon(myPath, "Quotation_MAST.xlt", "V", "S", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "E")
                Else
                    InHoaDon(myPath, "Quotation.xlt", "V", "S", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "E")
                End If

            End If
        End If

    End Sub
    Private Sub LblSCat_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblSCat.VisibleChanged
        Me.LblSType.Visible = Me.LblSCat.Visible
        Me.CmbSCat.Visible = Me.LblSCat.Visible
        Me.cmbStype.Visible = Me.LblSCat.Visible
        Me.LblD1.Visible = Me.LblSCat.Visible
        Me.TxtD1.Visible = Me.LblSCat.Visible
        Me.LblUnit.Visible = Not Me.LblSCat.Visible
        Me.TxtUnit.Visible = Not Me.LblSCat.Visible
    End Sub
    Private Sub CmbVendor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbVendor.SelectedIndexChanged
        Dim wholeSaler As Integer = ScalarToInt("MISC", "RecID", "cat='NonAirWS' and val='" & Me.CmbVendor.SelectedValue.ToString.Trim & "'")
        Dim strDK As String = " where status='OK' and left(fullName,5)<>'(SAI)' "
        Try
            If wholeSaler = 0 Then strDK = strDK & " and vendorID=" & Me.CmbVendor.SelectedValue
            LoadCmb_VAL(Me.CmbSupplier, "select RecID as VAL, FullName as DIS from Supplier " & strDK)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub TxtD2_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtD2.VisibleChanged
        Me.TxtT2.Visible = Me.TxtD2.Visible
    End Sub

    Private Sub TxtD1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtD1.VisibleChanged
        Me.TxtT1.Visible = Me.TxtD1.Visible
    End Sub

    Private Sub TxtHTLSF_LostFocus(sender As Object, e As EventArgs) Handles TxtHTLSF.LostFocus
        Dim aa As Decimal = CDec(Me.TxtHTLSF.Text) * CDec(Me.TxtUnitCost.Text) / 100
        Me.txtMU.Text = Format(aa, "#,##0.0")
    End Sub

    Private Sub LblSaveTour_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSaveTour.LinkClicked
        Dim strCmd As String, strKeyMap = defineKeymapFromCCenter()
        If Me.GridGO.Visible And InvalidSIR() Then Exit Sub
        If Me.CmbBooker.Text.ToUpper.Contains("ZPERS") And InStr("COD_CRD", Me.CmbBilling.Text) = 0 Then Exit Sub
        If CInt(Me.TxtPax.Text) = 0 Or Me.TxtBrief.Text = "" Or Me.CmbBilling.Text = "" Or _
            Me.TxtStartDate.Value.Date > Me.txtEndDate.Value.Date Or DuplicateEventCode(Me.txtRefCode.Text) Or _
            (Me.OptCWT.Checked And Me.cmbLocation.Text = "") Then
            MsgBox("Invalid NoOfPax or Brief or Billing or StartDate or PRNo or Empty Location for CWT client", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If myStaff.SupOf = "" And Me.TxtStartDate.Value < Now.Date Then
            MsgBox("Invalid StartDate", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        strCmd = UpdateLogFile("Dutoan_tour", "Edit", Me.GridTour.CurrentRow.Cells("RecID").Value, strKeyMap, _
                Me.TxtEmail.Text, Me.TxtBrief.Text.Replace("'", ""), Me.cmbLocation.Text, Me.CmbBooker.Text, Me.cmbTraveler.Text, Me.TxtIONo.Text, _
                Me.txtRefCode.Text, Me.TxtOwner.Text, Me.TxtPRNo.Text, Me.CmbBUAdmin.Text)
        cmd.CommandText = strCmd
        cmd.ExecuteNonQuery()

        strCmd = String.Format("update Dutoan_Tour set contact='{0}', Traveller='{1}', IONO='{2}', RefNo='{3}',Owner='{4}'," & _
            "PRNo='{5}', BUAdmin='{6}', Dept='{7}', Location='{8}', KeyMap='{9}'" _
            & ", Email='{10}', Brief='{11}' , WindowId='{12}' ", _
            Me.CmbBooker.Text.Replace("--", ""), Me.cmbTraveler.Text.Replace("--", ""), _
            Me.TxtIONo.Text.Replace("--", ""), Me.txtRefCode.Text.Replace("--", ""), Me.TxtOwner.Text.Replace("--", ""), _
            Me.TxtPRNo.Text.Replace("--", ""), Me.CmbBUAdmin.Text.Replace("--", ""), Me.CmbDept.Text.Replace("--", ""), _
            IIf(Me.OptCWT.Checked, "N/A", Me.cmbLocation.Text.Replace("--", "")), strKeyMap.Replace("--", ""), Me.TxtEmail.Text.Replace("--", ""), _
            Me.TxtBrief.Text.Replace("--", "").Replace("'", ""), txtWindowId.Text.Trim)
        If myStaff.SupOf <> "" Then
            strCmd = String.Format(" {0}, Sdate='{1}', EDate='{2}', EventDate='{3}' ", strCmd, Me.TxtStartDate.Value.Date, Me.txtEndDate.Value.Date, Me.txtEventDate.Value.Date)
            If Me.GridTour.CurrentRow.Cells("Status").Value <> "RR" Then
                strCmd = String.Format("{0}, BillingBy='{1}'", strCmd, Me.CmbBilling.Text.Replace("--", ""))
            End If
        End If
        strCmd = strCmd & " where RecID=" & Me.GridTour.CurrentRow.Cells("RecID").Value
        cmd.CommandText = strCmd
        cmd.ExecuteNonQuery()

        If Me.GridGO.Visible Then
            InsertSIR(Me.GridTour.CurrentRow.Cells("RecID").Value, True)
        End If

        LoadgridTour()
    End Sub

    Private Sub ChkPastOnly_Click(sender As Object, e As EventArgs) Handles ChkPastOnly.Click _
        , ChkSelectedCustOnly.Click
        LoadgridTour()
    End Sub

    Private Sub LblPreview_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblPreview.LinkClicked
        Dim myPath As String = Application.StartupPath
        If Me.GridTour.CurrentRow.Cells("CustShortName").Value = "ROCHE" Then
            InHoaDon(myPath, "Quotation_Roche.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "V")
        ElseIf Me.GridTour.CurrentRow.Cells("CustShortName").Value = "MAST" Then
            InHoaDon(myPath, "Quotation_MAST.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "V")
        Else
            InHoaDon(myPath, "Quotation.xlt", "V", "Q", Now.Date, Now.Date, Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "V")
        End If
    End Sub

    Private Sub ChkXXOnly_CheckedChanged(sender As Object, e As EventArgs) Handles ChkXXOnly.CheckedChanged
        If Me.ChkXXOnly.Checked Then
            Me.ChkXXOnly.Left = 496
            Me.ChkXXOnly.Top = 480
            Me.GridSVC.BringToFront()
        Else
            Me.ChkXXOnly.Left = 43
            Me.ChkXXOnly.Top = 278
        End If
        'LoadGridSVC(Me.GridTour.CurrentRow.Cells("RecID").Value)
    End Sub

    Private Sub cmbStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStatus.SelectedIndexChanged
        If Not mblnFirstLoadCompleted Then
            Exit Sub
        End If
        If Me.cmbStatus.Text = "CXLD" Then
            Me.TabControl1.Enabled = False
            'Me.GridTour.Height = 470
            Me.ChkXXOnly.Checked = True
            Me.TabControl1.SelectTab("TabCosting")
        Else
            Me.TabControl1.Enabled = True
            'Me.GridTour.Height = 454
            Me.ChkXXOnly.Checked = False
        End If
        LoadgridTour()
    End Sub
    Private Sub LblDocSent_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDocSent.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("Dutoan_Tour", "OD", Me.GridTour.CurrentRow.Cells("recID").Value) & _
            "; " & UpdateLogFile("dutoan_Tour", "DocSent", Me.GridTour.CurrentRow.Cells("recID").Value, "", "", "", "")
        cmd.ExecuteNonQuery()
        LoadgridTour()
    End Sub
    Private Sub LckLblUndoFinalize_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LckLblUndoFinalize.LinkClicked
        If MsgBox("Are You Sure To Undo Finalize", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle) = vbNo Then Exit Sub
        Dim DKDocument As String = " and (Document = '" & Me.GridTour.CurrentRow.Cells("TCode").Value & "' or RMK like '%" & _
            Me.GridTour.CurrentRow.Cells("TCode").Value & "%')"
        Dim dTbl As DataTable = GetDataTable("Select * from FOP where status in ('OK','QQ') and RMK like '%" & _
                                             Me.GridTour.CurrentRow.Cells("RecID").Value & "%'" & DKDocument)
        If Me.GridTour.CurrentRow.Cells("BillingBy").Value <> BillBy.BillByEvent Then
            If VeCongNoDaInv(dTbl.Rows(0)("RCPID")) Then
                MsgBox("This TRX has Been Invoiced And Cant Be Changed", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
            If dTbl.Rows(0)("FOP") = "CRD" And dTbl.Rows(0)("Status") = "OK" Then
                MsgBox("This TRX Has Been Paid By CC And Cant Be Changed", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        End If
        cmd.CommandText = ChangeStatus_ByID("Dutoan_Tour", "OK", Me.GridTour.CurrentRow.Cells("recID").Value)
        If Me.GridTour.CurrentRow.Cells("BillingBy").Value <> BillBy.BillByEvent Then
            cmd.CommandText = cmd.CommandText & "; update RCP set status='XX' where counter='N-A' and recid=" & dTbl.Rows(0)("RCPID") & _
                ";" & ChangeStatus_ByDK("FOP", "XX", "RcpID=" & dTbl.Rows(0)("RCPID") & DKDocument)
            If Me.GridTour.CurrentRow.Cells("BillingBy").Value = BillBy.BillByBundle Then
                cmd.CommandText = cmd.CommandText & "; Update FOP set FOP='MCE' where status='OK' and RMK='TT.INV.BDL' " & DKDocument
            End If
        End If
        cmd.ExecuteNonQuery()
        LoadgridTour()
    End Sub

    Private Sub TxtEmail_Enter(sender As Object, e As EventArgs) Handles TxtEmail.Enter
        If Strings.Left(Me.TxtEmail.Text, 1) = "_" Then
            Me.TxtEmail.Text = ""
            Me.TxtEmail.ForeColor = Color.Black
        End If
    End Sub


    Private Sub LblAddSF_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAddSF.LinkClicked
        Me.GrpCost.Tag = "SVC"
        Me.GrpCost.Enabled = False
        Me.chkBookOnly.Checked = False
    End Sub

    Private Sub GrpCost_EnabledChanged(sender As Object, e As EventArgs) Handles GrpCost.EnabledChanged
        Me.Label7.Visible = Not Me.GrpCost.Enabled
        Me.TxtTvSfPct.Visible = Not Me.GrpCost.Enabled
        Me.TxtTVSfAmount.Visible = Not Me.GrpCost.Enabled
        txtTvSfQty.Visible = Not Me.GrpCost.Enabled
        lblTvSfQty.Visible = Not Me.GrpCost.Enabled
    End Sub

    Private Sub TxtPRNo_Enter(sender As Object, e As EventArgs) Handles TxtPRNo.Enter
        If Strings.Left(Me.TxtPRNo.Text, 1) = "_" Then
            Me.TxtPRNo.Text = ""
            Me.TxtPRNo.ForeColor = Color.Black
        End If
    End Sub
    Private Sub TxtOwner_Enter(sender As Object, e As EventArgs) Handles TxtOwner.Enter
        If Me.TxtOwner.Text.Contains("_") Then
            Me.TxtOwner.Text = ""
            Me.TxtOwner.ForeColor = Color.Black
        End If
    End Sub
    Private Function defineTCList(wzAmt As Boolean) As String
        Dim KQ As String = ""
        For i As Int16 = 0 To Me.GridPmtRQTC.RowCount - 1
            If Me.GridPmtRQTC.Item("Q", i).Value Then
                KQ = KQ & "_" & Me.GridPmtRQTC.Item("TCode", i).Value
                If wzAmt Then
                    KQ = KQ & ": " & Format(Me.GridPmtRQTC.Item("BeingPaidThisTime", i).Value, "#,##0")
                End If
            End If
        Next
        If KQ.Length > 2 Then KQ = KQ.Substring(1)
        Return KQ
    End Function
    Private Function DefineTCList_Combine(pPmtID As Integer, wzAmt As Boolean) As String
        Dim dtbl As DataTable = GetDataTable("select TCode, sum(vnd) as Amt from dutoan_pmt where status='OK' and pmtID=" & pPmtID & " group by TCode")
        Dim KQ As String = ""
        For i As Int16 = 0 To dtbl.Rows.Count - 1
            KQ = KQ & "_" & dtbl.Rows(i)("Tcode")
            If wzAmt Then
                KQ = KQ & ": " & Format(dtbl.Rows(i)("Amt"), "#,##0")
            End If
        Next
        Return KQ.Substring(1)
    End Function
    Private Sub LblPrintPmtRQ_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblPrintPmtRQ.LinkClicked
        If HasNewerVersion_R12(Application.ProductVersion) Or SysDateIsWrong(Conn) Then Exit Sub
        Dim vendorID As Integer, ShortName As String = "", MoTa As String, TCode As String, PmtID As Integer, RMKNoiBo As String
        Dim CAT As String = "", myPath As String = Application.StartupPath
        Dim AccountName As String = "", AccountNumber As String = "", BankName As String = ""
        Dim BankAddress As String = "", swift As String = "", PayeeAccID As Integer = 0

        If CDec(Me.txtTraLanNay.Text) = 0 Then Exit Sub
        ShortName = Me.CmbPmtRQVendor.Text
        vendorID = Me.CmbPmtRQVendor.SelectedValue
        If Not isLogicAmt() Then Exit Sub
        TCode = defineTCList(False)
        RMKNoiBo = defineTCList(True)

        Me.txtTraLanNay.Text = Format(XacDinhAmtTraLanNay(), "#,##0")
        If CDec(Me.txtTraLanNay.Text) < 0 Then
            MoTa = "Hoan Ung " & TCode & ". So phieu thu: " & InputBox("Please Enter [PhieuThu] No.", msgTitle)
            Me.CmbPmtRQFOP.Text = "CSH"
        Else
            MoTa = "TransViet TToan " & TCode
        End If

        If Me.GridAcct.Rows.Count = 0 Then
            MsgBox("No Bank Details. Please Update ", MsgBoxStyle.Critical, msgTitle)
            If Me.CmbPmtRQFOP.Text = "BTF" Then
                Exit Sub
            End If
        Else
            AccountName = Me.GridAcct.CurrentRow.Cells("AccountName").Value
            AccountNumber = Me.GridAcct.CurrentRow.Cells("AccountNumber").Value
            BankName = Me.GridAcct.CurrentRow.Cells("BankName").Value
            BankAddress = Me.GridAcct.CurrentRow.Cells("BankAddress").Value
            swift = Me.GridAcct.CurrentRow.Cells("swift").Value
            PayeeAccID = Me.GridAcct.CurrentRow.Cells("RecID").Value
            CAT = ScalarToString("UNC_Company", "CAT", "RecID=" & vendorID)
            CAT = "[" & CAT & "] "
            If String.IsNullOrEmpty(AccountName) Then
                MsgBox("You Must Specify Payee Account", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        End If

        Dim strVendorShortName As String = Me.CmbPmtRQVendor.Text
        If strVendorShortName = "SO Y TE" Then
            strVendorShortName = strVendorShortName & "-" & ScalarToString("dutoan_tour", "Location" _
                                                                           , "where status<>'xx' and tcode='" & TCode & "'")
        End If

        InHoaDon(myPath, "DeNghiThanhToan.xlt", "V", strVendorShortName, OverPay, Now, vendorID, myStaff.SICode, MoTa, "VND " & Me.txtTraLanNay.Text)
        Dim myAns As Int16 = MsgBox("Are You OK With the PrintOut?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
        If myAns = vbNo Then Exit Sub

        cmd.CommandText = "Insert into UNC_Payments (PayerAccountID, AccountName, AccountNumber, BankName, BankAddress, " & _
            " Curr, Amount, Description, swift, Charge, shortname, InvNo, TRX_TC, FstUser, Status, FOP, PayeeAccountID) values (" & _
            " 0, @AccountName, @AccountNumber, @BankName, @BankAddress, 'VND', @Amount, @Description, @swift, 'IPOB', @shortname," & _
            "'', @TRX_TC, @FstUser,'QQ',@FOP, @PayeeAccountID);SELECT SCOPE_IDENTITY() AS [RecID]"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@AccountName", SqlDbType.NVarChar).Value = AccountName
        cmd.Parameters.Add("@AccountNumber", SqlDbType.VarChar).Value = AccountNumber
        cmd.Parameters.Add("@BankName", SqlDbType.NVarChar).Value = BankName
        cmd.Parameters.Add("@BankAddress", SqlDbType.NVarChar).Value = BankAddress
        cmd.Parameters.Add("@Amount", SqlDbType.Decimal).Value = CDec(Me.txtTraLanNay.Text)
        cmd.Parameters.Add("@Description", SqlDbType.VarChar).Value = MoTa
        cmd.Parameters.Add("@swift", SqlDbType.VarChar).Value = swift
        cmd.Parameters.Add("@shortname", SqlDbType.VarChar).Value = ShortName
        cmd.Parameters.Add("@FOP", SqlDbType.VarChar).Value = Me.CmbPmtRQFOP.Text
        cmd.Parameters.Add("@TRX_TC", SqlDbType.VarChar).Value = "PPD: NA"
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@PayeeAccountID", SqlDbType.Int).Value = PayeeAccID
        PmtID = cmd.ExecuteScalar
        cmd.CommandText = ""
        cmd.Parameters.Clear()
        For i As Int16 = 0 To Me.GridPmtRQTC.RowCount - 1
            If Me.GridPmtRQTC.Item("Q", i).Value Then
                cmd.CommandText = cmd.CommandText & "; insert Dutoan_pmt (DutoanID, TCode, ItemID, VendorID, Vendor, PmtID, VND, FstUser)" & _
                    " values (" & Me.GridPmtRQTC.Item("DuToanID", i).Value & ",'" & Me.GridPmtRQTC.Item("TCode", i).Value & "'," & _
                    Me.GridPmtRQTC.Item("RecID", i).Value & "," & Me.CmbPmtRQVendor.SelectedValue & ",'" & _
                    Me.CmbPmtRQVendor.Text & "'," & PmtID & "," & Me.GridPmtRQTC.Item("BeingPaidThisTime", i).Value & _
                    ",'" & myStaff.SICode & "')"
            End If
        Next
        cmd.CommandText = cmd.CommandText.Substring(1)
        cmd.ExecuteNonQuery()

        For i As Int16 = 0 To Me.GridPmtRQTC.RowCount - 1
            If Me.GridPmtRQTC.Item("Q", i).Value Then
                RefreshSoDaTra(Me.GridPmtRQTC.Item("RecID", i).Value)
            End If
        Next
        TCode = DefineTCList_Combine(PmtID, False)
        RMKNoiBo = DefineTCList_Combine(PmtID, True)
        If CDec(Me.txtTraLanNay.Text) < 0 Then
            RMKNoiBo = "Hoan Ung " & RMKNoiBo & ". So phieu thu: " & InputBox("Please Enter [PhieuThu] No.", msgTitle)
            Me.CmbPmtRQFOP.Text = "CSH"
        Else
            MoTa = "TransViet TToan " & TCode
        End If
        cmd.CommandText = "update UNC_Payments set Description='" & MoTa.Replace("--", "") & "', RMKNoibo='" & _
            RMKNoiBo.Replace("--", "") & "' where recid=" & PmtID
        cmd.ExecuteNonQuery()
        InHoaDon(myPath, "DeNghiThanhToan.xlt", "O", CAT & Me.CmbPmtRQVendor.Text, OverPay, Now, vendorID, myStaff.SICode, RMKNoiBo, "VND " & Me.txtTraLanNay.Text, PmtID)
        Me.LblPreview.Visible = True
        If Me.CmbPmtRQFOP.Text = "CSH" Then
            If MsgBox("Wanna Order a Messenger To Fulfill This Payment?", MsgBoxStyle.Question Or vbYesNo, msgTitle) = vbYes Then
                Book_a_MSGR(myStaff.SICode, myStaff.PSW, "N/A", IIf(TCode.Length > 16, "PmtByVendor", TCode), PmtID)
            End If
        End If
    End Sub
    Private Sub RefreshSoDaTra(pItemID As Integer)
        Dim DaTra As Decimal = ScalarToDec("dutoan_pmt", "isnull(sum(VND),0)", "ItemID=" & pItemID & " and status<>'XX'")
        cmd.CommandText = "Update Dutoan_item set VNDPaid=" & DaTra & " where RecID=" & pItemID
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub LblQCSF_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblQCSF.LinkClicked
        Me.GrpCost.Tag = "QC"
        Me.GrpCost.Enabled = False
        Dim CT As String, MealTime As Integer = CInt(Format(Me.GridSVC.CurrentRow.Cells("SVCDate").Value, "HH"))
        CT = Me.GridSVC.CurrentRow.Cells("City").Value
        Dim QCFee As Decimal = ScalarToDec("MISC", "VAL1", "CAT='QCFEE' and VAL='" & CT & "'")
        Dim OverTime As Decimal = ScalarToDec("MISC", "VAL2", "CAT='QCFEE' and VAL='" & CT & "'")
        If MealTime > 17 Then QCFee = QCFee + OverTime
        Me.TxtTVSfAmount.Text = Format(QCFee, "#,##0")
    End Sub
    Private Sub LblAddSF_VisibleChanged(sender As Object, e As EventArgs) Handles LblAddSF.VisibleChanged
        Try
            If MyCust.ShortName.Contains("SANOFI") Or MyCust.ShortName.Contains("TEST") Then
                Me.LblQCSF.Visible = Me.LblAddSF.Visible
            Else
                Me.LblQCSF.Visible = False
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Function isLogicAmt() As Boolean
        Dim TotalToVendor As Decimal, DaTra As Decimal
        Dim LanNay As Decimal, MyAns As Int16
        OverPay = "01-jan-2000"
        For i As Int16 = 0 To Me.GridPmtRQTC.RowCount - 1
            If Me.GridPmtRQTC.Item("Q", i).Value Then
                LanNay = Me.GridPmtRQTC.Item("BeingPaidThisTime", i).Value
                TotalToVendor = Me.GridPmtRQTC.Item("toVendorWoTax", i).Value
                DaTra = Me.GridPmtRQTC.Item("VNDPaid", i).Value
                If LanNay > TotalToVendor - DaTra Then
                    MyAns = MsgBox("Invalid Amount. Wanna Correct Your Input?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, msgTitle)
                    If MyAns = vbYes Then Return False
                    OverPay = "31-Dec-2000"
                End If
            End If
        Next
        Return True
    End Function
    Private Sub LblAddAdj_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAddAdj.LinkClicked
        If CDec(Me.TxtAmtAdj.Text) = 0 Or Me.TxtRMKAdj.Text = "" Then Exit Sub
        cmd.CommandText = "insert DuToan_Item (Service, CCurr, Status, Qty, Cost, Supplier, VendorID, Vendor, FstUser, PmtMethod, isVATIncl, " & _
            " VAT, DuToanID, Brief, SupplierRMK, MU, SVCDate, City, BookOnly) Values (@Service, @CCurr, 'TV', @Qty, @Cost, @Supplier, @VendorID, " & _
            "@Vendor, @FstUser, @PmtMethod, @isVATIncl, @VAT, @DuToanID, @Brief, @SupplierRMK, @MU, @SVCDate, @City, @BookOnly) "
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@Service", SqlDbType.VarChar).Value = "ADJ"
        cmd.Parameters.Add("@CCurr", SqlDbType.VarChar).Value = Me.CmbCurrAdj.Text
        cmd.Parameters.Add("@Qty", SqlDbType.Decimal).Value = 1
        cmd.Parameters.Add("@Cost", SqlDbType.Decimal).Value = CDec(Me.TxtAmtAdj.Text)
        cmd.Parameters.Add("@Supplier", SqlDbType.VarChar).Value = Me.CmbVendorAdj.Text
        cmd.Parameters.Add("@Vendor", SqlDbType.VarChar).Value = Me.CmbVendorAdj.Text
        cmd.Parameters.Add("@VendorID", SqlDbType.Int).Value = Me.CmbVendorAdj.SelectedValue
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@PmtMethod", SqlDbType.VarChar).Value = "PSP"
        cmd.Parameters.Add("@City", SqlDbType.VarChar).Value = "ALL"
        cmd.Parameters.Add("@isVATIncl", SqlDbType.Bit).Value = 1
        cmd.Parameters.Add("@VAT", SqlDbType.Decimal).Value = 0
        cmd.Parameters.Add("@DuToanID", SqlDbType.Int).Value = Me.GridTour.CurrentRow.Cells("RecID").Value
        cmd.Parameters.Add("@Brief", SqlDbType.VarChar).Value = Me.TxtRMKAdj.Text
        cmd.Parameters.Add("@SupplierRMK", SqlDbType.NVarChar).Value = ""
        cmd.Parameters.Add("@MU", SqlDbType.Decimal).Value = 0
        cmd.Parameters.Add("@SVCDate", SqlDbType.DateTime).Value = Now.Date
        cmd.Parameters.Add("@BookOnly", SqlDbType.Bit).Value = 0
        cmd.ExecuteNonQuery()
        LoadGridAdj()
    End Sub
    Private Sub LoadGridAdj()
        Me.GridAdj.DataSource = GetDataTable("select RecID, VendorID, cCurr as Curr, Cost as Amount, Vendor, Brief from dutoan_item where status='TV' and dutoanID=" & Me.GridTour.CurrentRow.Cells("RecID").Value)
        Me.GridAdj.Columns(0).Visible = False
        Me.GridAdj.Columns(1).Visible = False
        Me.GridAdj.Columns("Curr").Width = 32
        Me.GridAdj.Columns("Amount").Width = 75
        Me.GridAdj.Columns("Vendor").Width = 128
        Me.GridAdj.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridAdj.Columns("Amount").DefaultCellStyle.Format = "#,###.0"
        Me.LblDeleteAdj.Visible = False
    End Sub

    Private Sub GridAdj_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridAdj.CellContentClick
        Me.LblDeleteAdj.Visible = True
    End Sub
    Private Sub LblDeleteAdj_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDeleteAdj.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("Dutoan_Item", "XX", Me.GridAdj.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        LoadGridAdj()
    End Sub

    Private Sub OptCWT_CheckedChanged(sender As Object, e As EventArgs) Handles OptCWT.CheckedChanged
        LoadCmb_VAL(Me.CmbCust, MyCust.List_CS)
    End Sub
    Private Sub OptLCL_Click(sender As Object, e As EventArgs) Handles OptLCL.CheckedChanged
        If mblnFirstLoadCompleted Then
            LoadCmb_VAL(Me.CmbCust, MyCust.List_LC)
        End If

    End Sub
    Private Sub txtRefCode_Enter(sender As Object, e As EventArgs) Handles txtRefCode.Enter
        If Me.txtRefCode.Text.Substring(0, 1) = "_" Then
            Me.txtRefCode.Text = ""
            Me.txtRefCode.ForeColor = Color.Black
        End If
    End Sub

    Private Sub OptPmtByTCode_Click(sender As Object, e As EventArgs) Handles OptPmtByTCode.Click, OptPmtByVendor.Click
        Me.CmbPmtRQSVC.Visible = Me.OptPmtByVendor.Checked
    End Sub
    Private Sub CmbPmtRQVendor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbPmtRQVendor.SelectedIndexChanged
        Try
            If Me.OptPmtByVendor.Checked Then
                LoadGridPmtRQ(0, Me.CmbPmtRQVendor.SelectedValue)
            Else
                LoadGridPmtRQ(Me.GridTour.CurrentRow.Cells("RecID").Value, Me.CmbPmtRQVendor.SelectedValue)
            End If
            Me.GridAcct.DataSource = GetDataTable("select AccountNumber, BankName, AccountName, BankAddress,swift, RecID " & _
                                      " from unc_accounts where companyID=" & Me.CmbPmtRQVendor.SelectedValue & " and status='OK'")

        Catch ex As Exception
        End Try
    End Sub

    Private Sub CmbPmtRQSVC_LostFocus(sender As Object, e As EventArgs) Handles CmbPmtRQSVC.LostFocus
        LoadCmb_VAL(Me.CmbPmtRQVendor, "select RecId as VAL, ShortName as DIS from unc_company" _
                    & " where status='OK' and RecId<>2 and RecId in " _
                    & "(Select distinct VendorId from Dutoan_item where Status='OK' and Service='" & Me.CmbPmtRQSVC.Text & "')")
    End Sub
    Private Sub OptPmtByVendor_CheckedChanged(sender As Object, e As EventArgs) Handles OptPmtByVendor.CheckedChanged
        Me.TabControl1.Left = 0
        Me.TabControl1.Width = 975

        EnableSearchControl(Not OptPmtByVendor.Checked)
        
    End Sub

    Private Sub EnableSearchControl(blnEnable As Boolean)
        cboChannel.Visible = blnEnable
        cboCustShortName.Visible = blnEnable
        txtSearchTcode.Visible = blnEnable
        lblTcode.Visible = blnEnable
        LblSearch.Visible = blnEnable
        lbkReset.Visible = blnEnable
    End Sub

    Private Sub OptPmtByTCode_CheckedChanged(sender As Object, e As EventArgs) Handles OptPmtByTCode.CheckedChanged
        Me.TabControl1.Left = 431
        Me.TabControl1.Width = 546
    End Sub


    Private Sub TabPage5_Leave(sender As Object, e As EventArgs) Handles TabPage5.Leave
        Me.OptPmtByTCode.PerformClick()
    End Sub
    Private Function XacDinhAmtTraLanNay() As Decimal
        Dim AmtThisTime As Decimal = 0
        For i As Int16 = 0 To Me.GridPmtRQTC.RowCount - 1
            If Me.GridPmtRQTC.Item("BeingPaidThisTime", i).Value = 0 Then
                Me.GridPmtRQTC.Item("Q", i).Value = False
            End If
        Next
        For i As Int16 = 0 To Me.GridPmtRQTC.RowCount - 1
            If Me.GridPmtRQTC.Item("Q", i).Value Then
                AmtThisTime = AmtThisTime + Me.GridPmtRQTC.Item("BeingPaidThisTime", i).Value
            End If
        Next
        Return AmtThisTime
    End Function
    Private Sub txtTraLanNay_Enter(sender As Object, e As EventArgs) Handles txtTraLanNay.Enter
        Me.txtTraLanNay.Text = Format(XacDinhAmtTraLanNay(), "#,##0")
        Me.CmbPmtRQFOP.Text = "BTF"
    End Sub
    Private Sub LoadGridPmtRQ(ByVal pTourID As Integer, pVendorID As Integer)
        Me.GridPmtRQTC.DataSource = Nothing
        Dim StrQry As String = "select i.RecID, DuToanID, Q, TCode, CCurr,Cost, Qty, VAT, isVATIncl, toVendorWoTax, VNDPaid," & _
            " toVendorWoTax - VNDPaid as BeingPaidThisTime" & _
            " from Dutoan_item i inner join dutoan_tour t on t.recid=i.dutoanid and i.status<>'XX' "
        StrQry = StrQry & " where vendorID<>2 and bookonly=0 and toVendorWoTax - VNDPaid <>0"
        If pTourID > 0 Then
            StrQry = StrQry & " and DutoanID=" & pTourID
        End If
        StrQry = StrQry & " and VendorID=" & pVendorID & " and (pmtMethod='PPD' or NeedDeposit <>0) and svcdate >'1-Jan-16'"
        If pVendorID = 5042 Then ' so YTe thi ko lay HN
            StrQry = StrQry & " and dutoanID not in (select RecID from dutoan_tour where location='Ha Noi')"
        End If
        Me.GridPmtRQTC.DataSource = GetDataTable(StrQry)
        If pTourID > 0 Then Me.GridPmtRQTC.Columns("TCode").Visible = False
        Me.GridPmtRQTC.Columns("RecID").Visible = False
        Me.GridPmtRQTC.Columns("DuToanID").Visible = False
        Me.GridPmtRQTC.Columns("CCurr").Width = 32
        Me.GridPmtRQTC.Columns("Q").Width = 25
        Me.GridPmtRQTC.Columns("VAT").Width = 56
        Me.GridPmtRQTC.Columns("Cost").Width = 70
        Me.GridPmtRQTC.Columns("toVendorWoTax").Width = 75
        Me.GridPmtRQTC.Columns("VNDPaid").Width = 75
        Me.GridPmtRQTC.Columns("Qty").Width = 32
        Me.GridPmtRQTC.Columns("VAT").Width = 32
        Me.GridPmtRQTC.Columns("isVATincl").Width = 56
        Me.GridPmtRQTC.Columns("Cost").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridPmtRQTC.Columns("Cost").DefaultCellStyle.Format = "#,##0.0"
        Me.GridPmtRQTC.Columns("toVendorWoTax").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridPmtRQTC.Columns("toVendorWoTax").DefaultCellStyle.Format = "#,##0.0"
        Me.GridPmtRQTC.Columns("VNDPaid").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridPmtRQTC.Columns("VNDPaid").DefaultCellStyle.Format = "#,##0.0"
        Me.GridPmtRQTC.Columns("BeingPaidThisTime").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridPmtRQTC.Columns("BeingPaidThisTime").DefaultCellStyle.Format = "#,##0.0"
        For i As Int16 = 0 To Me.GridPmtRQTC.RowCount - 1
            Me.GridPmtRQTC.Item("Q", i).Value = False
        Next
        For i As Int16 = 3 To Me.GridPmtRQTC.Columns.Count - 2
            Me.GridPmtRQTC.Columns(i).ReadOnly = True
        Next
    End Sub

    Private Sub TxtIONo_Enter(sender As Object, e As EventArgs) Handles TxtIONo.Enter
        If Me.TxtIONo.Text.ToUpper.Substring(0, 1) = "_" Then
            Me.TxtIONo.Text = ""
            Me.TxtIONo.ForeColor = Color.Black
        End If
    End Sub

    Private Sub LblRefreshPaid_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblRefreshPaid.LinkClicked
        If Me.OptPmtByVendor.Checked Then Exit Sub
        Dim StrQry As String = "select RecID from Dutoan_item where status <>'XX' and vendorID= " & Me.CmbPmtRQVendor.SelectedValue & _
                " and DutoanID=" & Me.GridTour.CurrentRow.Cells("RecID").Value
        Dim dTbl As DataTable = GetDataTable(StrQry)
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            RefreshSoDaTra(dTbl.Rows(i)("RecID"))
        Next
    End Sub
    Private Sub LoadGridGO(blnCreate As Boolean)
        Dim chkLocalCorpID As Integer
        Dim strQry As String
        Me.GridGO.Visible = False
        If Me.OptLCL.Checked Then
            chkLocalCorpID = ScalarToInt("CWT.dbo.go_companyinfo1", "RecID", "CustID=" & Me.CmbCust.SelectedValue & " and NoData4NonAir=0 and status='OK'")
            If chkLocalCorpID = 0 Then Exit Sub
            chkLocalCorpID = ScalarToInt("CWT.dbo.GO_RequiredData", "RecID", "CustID=" & Me.CmbCust.SelectedValue & " and collectionMethod='agtinput' and status='OK'")
            If chkLocalCorpID = 0 Then Exit Sub
            Me.GridGO.Visible = True
            If blnCreate Then
                strQry = "Select NameByCustomer, '' as FieldValues,DataCode, MinLength" _
                    & ", MaxLength, CharType,CheckValues,'RequiredData' as Source" & _
                " from [CWT].[dbo].[GO_RequiredData]" _
                & " where status='OK' and collectionmethod='agtinput'" _
                & " and ApplyTo in ('ALL','N-A','HTL','CAR')" _
                & " and custID=" & Me.CmbCust.SelectedValue
            Else
                strQry = " Select NameByCustomer, FValue as FieldValues,DataCode, MinLength" _
                    & ", MaxLength, CharType,CheckValues,'RequiredData'  as Source" _
                    & " from [CWT].[dbo].[GO_RequiredData] r left join cwt.dbo.SIR s" _
                    & " on r.NameByCustomer =s.Fname and s.Prod='nonair' and s.Status='OK' and RCPID=" & GridTour.CurrentRow.Cells("Recid").Value _
                    & " where r.status='OK' and collectionmethod='agtinput'" _
                    & " and ApplyTo in ('ALL','N-A','HTL','CAR')" _
                    & " and r.custID=" & Me.CmbCust.SelectedValue
            End If
            
        ElseIf blnCreate Then
            strQry = "select CdrName, '' as FieldValues, CdrNbr as DataCode, MinLength, MaxLength" _
                & ", CharType,CheckValues,'CDR' as Source" _
                & " from cwt.dbo.go_cdrs c " _
                & " where c.status='OK' and collectionmethod='agtinput'" _
                & " and CMC in (select CMC from cwt.dbo.GO_CompanyInfo1 where custid=" _
                & Me.CmbCust.SelectedValue & " and status='OK' " & _
                " and go_Client<>0 and NoData4NonAir=0)" _
                & " union all Select NameByCustomer, '' as FieldValues,DataCode, MinLength" _
                & ", MaxLength, CharType,CheckValues,'RequiredData'  as Source" _
                & " from [CWT].[dbo].[GO_RequiredData]" _
                & " where status='OK' and collectionmethod='agtinput'" _
                & " and ApplyTo in ('ALL','N-A','HTL','CAR')" _
                & " and custID=" & Me.CmbCust.SelectedValue
        Else
            strQry = "select CdrName, FValue as FieldValues, CdrNbr as DataCode, MinLength, MaxLength" _
                & ", CharType,CheckValues,'CDR' as Source" _
                & " from cwt.dbo.go_cdrs c left join cwt.dbo.SIR s" _
                & " on c.CdrName =s.Fname and s.Prod='nonair' and s.Status='OK' and RCPID=" & GridTour.CurrentRow.Cells("Recid").Value _
                & " where c.status='OK' and collectionmethod='agtinput'" _
                & " and CMC in (select CMC from cwt.dbo.GO_CompanyInfo1 where custid=" & Me.CmbCust.SelectedValue _
                & " and status='OK' and go_Client<>0 and NoData4NonAir=0)" _
                & " union all " _
                & " Select NameByCustomer, FValue as FieldValues,DataCode, MinLength" _
                & ", MaxLength, CharType,CheckValues ,'RequiredData'  as Source" _
                & " from [CWT].[dbo].[GO_RequiredData] r left join cwt.dbo.SIR s" _
                & " on r.NameByCustomer =s.Fname and s.Prod='nonair' and s.Status='OK' and RCPID=" & GridTour.CurrentRow.Cells("Recid").Value _
                & " where r.status='OK' and collectionmethod='agtinput'" _
                & " and ApplyTo in ('ALL','N-A','HTL','CAR')" _
                & " and r.custID=" & Me.CmbCust.SelectedValue

        End If
        Me.GridGO.DataSource = GetDataTable(strQry)
        If Me.GridGO.RowCount = 0 Then
            Me.GridGO.Visible = False
        Else
            Me.GridGO.Visible = True
            For i As Int16 = 2 To Me.GridGO.ColumnCount - 1
                Me.GridGO.Columns(i).Visible = False
            Next
            
            Me.GridGO.Columns(0).ReadOnly = True
            Me.GridGO.Columns(0).Width = 200
            Me.GridGO.Columns(1).Width = 256
        End If
    End Sub
    Private Function InvalidSIR() As Boolean
        Dim KQ As Boolean = False, FVal As String
        For i As Int16 = 0 To Me.GridGO.RowCount - 1
            FVal = Me.GridGO.Item(1, i).Value.ToString.Trim
            If FVal <> "" And FVal <> "NIL" Then
                If FVal.Length < Me.GridGO.Item("MinLength", i).Value Or _
                    FVal.Length > Me.GridGO.Item("MaxLength", i).Value Then
                    MsgBox("Invalid length for " & GridGO.Item("NameByCustomer", i).Value)
                    Return True
                End If

                If Me.GridGO.Item("CharType", i).Value = "ALPHA" Then
                    For j As Int16 = 0 To FVal.Length - 1
                        If InStr("0123456789", FVal.Substring(j, 1)) > 0 Then Return True
                    Next
                End If
                If Me.GridGO.Item("CharType", i).Value = "NUMERIC" Then
                    For j As Int16 = 0 To FVal.Length - 1
                        If InStr("0123456789", FVal.Substring(j, 1)) = 0 Then Return True
                    Next
                End If
            End If
        Next
        Return KQ
    End Function
    Private Sub InsertSIR(pDutoanID As Integer, isDeleteB4 As Boolean)
        Dim strQry As String = ""
        If isDeleteB4 Then strQry = ";update cwt.dbo.SIR set status='XX', LstUpdate=getdate(), LstUser='" & myStaff.SICode & "' where PROD='NonAir' and RCPID=" & pDutoanID
        For i As Int16 = 0 To Me.GridGO.RowCount - 1
            If Me.GridGO.Item(1, i).Value.ToString.ToUpper.Trim.Length > 0 Then
                strQry = strQry & "; Insert cwt.dbo.SIR (RCPID, PROD, FName, FValue, CustID, fstUser) values (" & pDutoanID _
                    & ",'NonAir','" & Me.GridGO.Item(0, i).Value _
                    & "','" & Me.GridGO.Item(1, i).Value.ToString.ToUpper _
                    & "'," & Me.CmbCust.SelectedValue & ",'" & myStaff.SICode & "')"
            End If
        Next
        If strQry.Length > 2 Then
            cmd.CommandText = strQry.Substring(1)
            cmd.ExecuteNonQuery()
        End If
    End Sub

    Private Sub GridVendorInforUpdate_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridVendorInforUpdate.CellContentClick
        Me.LckLblUpdateVendorAddr.Visible = False
        If e.RowIndex < 0 Then Exit Sub
        Me.CmbLoaiChungTu.Text = Me.GridVendorInforUpdate.CurrentRow.Cells("HD").Value
        If GridVendorInforUpdate.CurrentRow.Cells(2).Value.ToString.Length > 16 Then Exit Sub
        Me.LckLblUpdateVendorAddr.Visible = True
    End Sub

    Private Sub LblUpdateVendorAddr_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LckLblUpdateVendorAddr.LinkClicked
        If myStaff.SupOf = "" Then Exit Sub
        Me.LckLblUpdateVendorAddr.Visible = False
        cmd.CommandText = "update UNC_Company set HD=@HD where recID=" & Me.GridVendorInforUpdate.CurrentRow.Cells(0).Value
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@HD", SqlDbType.NVarChar).Value = Me.CmbLoaiChungTu.Text
        cmd.ExecuteNonQuery()
    End Sub
    Private Function DefineBankFee(pSupplierID As Integer, pItemID As Integer) As Decimal
        Dim KQ As Decimal, Country As String, AmtToVendor As Decimal
        Country = ScalarToString("supplier", "Address_CountryCode", "RecID=" & pSupplierID)
        AmtToVendor = ScalarToDec("Dutoan_Item", "TTLToVendor", "RecID=" & pItemID)
        KQ = 0.0014 * AmtToVendor
        If KQ < 5 Then KQ = 5
        If KQ > 75 Then KQ = 75
        KQ = KQ + 10
        KQ = KQ + IIf(Country = "JP", 55, 25)
        Return KQ
    End Function

    Private Sub LblAddMerchantFee_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAddMerchantFee.LinkClicked
        Dim decKhachPhaiTra As Decimal = ScalarToDec("Dutoan_item", "Sum(TTLToPax)", "DutoanID=" & Me.GridTour.CurrentRow.Cells("RecID").Value & " and status<>'XX'")
        Dim decMF As Decimal
        Dim decPct As Decimal

        Select Case Me.CmbCrd.Text

            Case "VI", "MC", "CA"
                decPct = 0.0155
            Case "AX"
                decPct = 0.03
            Case "DC"
                decPct = 0.0385
            Case ""
                MsgBox("You must select Credit Card Type")
                Exit Sub
            Case Else
                MsgBox("Unable to add MerchantFee. Fee level is NOT specified")
                Exit Sub
        End Select

        decMF = Math.Round(decKhachPhaiTra / (1 - decPct) - decKhachPhaiTra, 0)
        
        AddFee("Merchant", "VND", 1, decMF, True, TxtVAT.Text, 0, 1)
        LoadGridSVC(Me.GridTour.CurrentRow.Cells("RecID").Value)
    End Sub
    Private Sub LblAddBancFee_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAddBancFee.LinkClicked
        Dim BF As Decimal
        For i As Int16 = 0 To Me.GridSVC.RowCount - 1
            If Me.GridSVC.Item("VendorID", i).Value <> 2 And _
                ScalarToString("supplier", "Address_CountryCode", "VendorID=" & Me.GridSVC.Item("VendorID", i).Value) <> "VN" Then
                BF = BF + DefineBankFee(Me.GridSVC.Item("SupplierID", i).Value, Me.GridSVC.Item("RecID", i).Value)
            End If
        Next
        AddFee("Bank", "USD", 1, BF, True, 0, 0, 0)
        LoadGridSVC(Me.GridTour.CurrentRow.Cells("RecID").Value)
    End Sub
    Private Sub AddFee(pFeeName As String, pCurr As String, pQty As Decimal, pAmt As Decimal, pVATInclude As Boolean, pVATAmt As Decimal, pRelatedItem As Integer, pROE As Decimal)
        cmd.CommandText = "insert DuToan_Item (Service, CCurr, Unit, Qty, Cost, Supplier, VendorID, Vendor, FstUser, PmtMethod, isVATIncl, " & _
            " VAT, DuToanID, SVCDate,  RelatedItem, ROE, SupplierID,BookOnly,PaxName,ZeroFeeReason)" _
            & "Values (@Service, @CCurr,@Unit, @Qty, @Cost, @Supplier, @VendorID, " & _
            "@Vendor, @FstUser, @PmtMethod, @isVATIncl, @VAT, @DuToanID, @SVCDate" _
            & ", @RelatedItem, @ROE, @SupplierID, @BookOnly,@PaxName,@ZeroFeeReason)"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@Service", SqlDbType.VarChar).Value = pFeeName & " Fee"
        cmd.Parameters.Add("@CCurr", SqlDbType.VarChar).Value = pCurr
        cmd.Parameters.Add("@Unit", SqlDbType.VarChar).Value = "Service"
        cmd.Parameters.Add("@Qty", SqlDbType.Decimal).Value = pQty
        cmd.Parameters.Add("@Cost", SqlDbType.Decimal).Value = pAmt
        cmd.Parameters.Add("@Supplier", SqlDbType.VarChar).Value = ""
        cmd.Parameters.Add("@Vendor", SqlDbType.VarChar).Value = "TransViet"
        cmd.Parameters.Add("@VendorID", SqlDbType.Int).Value = 2
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@PmtMethod", SqlDbType.VarChar).Value = "PSP"
        cmd.Parameters.Add("@isVATIncl", SqlDbType.Bit).Value = pVATInclude
        cmd.Parameters.Add("@BookOnly", SqlDbType.Bit).Value = 0
        cmd.Parameters.Add("@VAT", SqlDbType.Decimal).Value = pVATAmt
        cmd.Parameters.Add("@ROE", SqlDbType.Decimal).Value = pROE
        cmd.Parameters.Add("@DuToanID", SqlDbType.Int).Value = Me.GridTour.CurrentRow.Cells("RecID").Value
        cmd.Parameters.Add("@SVCDate", SqlDbType.DateTime).Value = Me.GridTour.CurrentRow.Cells("SDate").Value
        cmd.Parameters.Add("@RelatedItem", SqlDbType.Int).Value = pRelatedItem
        cmd.Parameters.Add("@SupplierID", SqlDbType.Int).Value = 2
        cmd.Parameters.Add("@PaxName", SqlDbType.VarChar).Value = txtPaxName.Text
        If pAmt = 0 Then
            cmd.Parameters.Add("@ZeroFeeReason", SqlDbType.VarChar).Value = ""
        Else
            cmd.Parameters.Add("@ZeroFeeReason", SqlDbType.VarChar).Value = cboZeroFeeReason.Text
        End If

        cmd.ExecuteNonQuery()
    End Sub
    Private Sub CmbSupplier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbSupplier.SelectedIndexChanged
        Try
            Me.TxtMoTaSVC.Text = ScalarToString("Supplier", "Address", "RecID=" & Me.CmbSupplier.SelectedValue)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CmbBooker_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbBooker.SelectedIndexChanged
        If Me.CmbBooker.Text.Contains("ZPERSONAL") Then Me.GridGO.Visible = False
    End Sub

    Private Sub LblSearch_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSearch.LinkClicked
        LoadgridTour()

        'Me.GridTour.DataSource = GetDataTable("select t.*,c.VAL as Channel from Dutoan_Tour t" _
        '                                      & " left join [Cust_Detail] c on t.CustId=c.CustId" _
        '                                      & " where " & strDK & " and c.Status='OK' and c.Cat='Channel'")
        'With GridTour
        '    If .RowCount > 0 Then
        '        .Columns("RecId").Width = 50
        '        .Columns("Pax").Width = 30
        '        .Columns("SDate").Width = 60
        '        .Columns("EDate").Width = 60
        '    End If
        'End With
        'LoadGridGO()
    End Sub
    Private Sub LblOrderBV_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblOrderBV.LinkClicked
        Book_a_MSGR(myStaff.SICode, myStaff.PSW, "N/A", Me.GridTour.CurrentRow.Cells("Tcode").Value, 0)
    End Sub

    Private Sub Data4GO_Enter(sender As Object, e As EventArgs) Handles TabData4GO.Enter

        'ShowCDRs()
        'FillCDRs()
    End Sub
    Private Sub ShowCDRs()
        Dim colCdrs As New Collection
        Dim strCmc As String = ScalarToString("cwt.dbo.go_CompanyInfo1", "Cmc", " Status='OK' and CustId=" _
                                              & GridTour.CurrentRow.Cells("CustId").Value)
        flpCdr.Controls.Clear()
        If strCmc <> "" Then
            colCdrs = GetCDRs(Conn, strCmc, True, True)
            For Each objCdr As clsCwtCdr In colCdrs
                Dim ucCdr As New ucCdr(objCdr)
                flpCdr.Controls.Add(ucCdr)
            Next
        End If
    End Sub
    
    Private Sub FillRequiredData(tblRequiredData As DataTable)

        If GridTour.CurrentRow.Cells("RequiredData").Value = "" Then
            Exit Sub
        End If
        Dim arrRequiredData As String() = GridTour.CurrentRow.Cells("RequiredData").Value.ToString.Split("|")
        Dim colAvailData As New Collection
        


        For Each ucCdr As ucCdr In flpCdr.Controls
            If colAvailData.Contains("CDR" & ucCdr.lblName.Tag) Then
                ucCdr.txtValue.Text = colAvailData("CDR" & ucCdr.lblName.Tag)
            End If
        Next

    End Sub
    Private Sub lbkSaveCDRs_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkSaveCDRs.LinkClicked
        'Dim strRqData As String = String.Empty
        'For Each ucCdr As ucCdr In flpCdr.Controls
        '    If ucCdr.txtValue.Text <> "" Then
        '        strRqData = strRqData & "CDR" & ucCdr.lblName.Tag & "/" & ucCdr.txtValue.Text & "|"
        '    End If
        'Next
        'If strRqData.Length > 0 Then
        '    strRqData = Mid(strRqData, 1, strRqData.Length - 1)
        'End If
        'ExecuteNonQuerry("Update DuToan_Tour set RequiredData='" & strRqData & "' where RecId=" _
        '                 & GridTour.CurrentRow.Cells("RecId").Value, Conn)
    End Sub



    Private Sub lbkUploadFile_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkUploadFile.LinkClicked
        If GridTour.CurrentRow IsNot Nothing Then
            Dim frmUpload As New frmUploadFile4NonAir(GridTour.CurrentRow, True)
            If frmUpload.ShowDialog = Windows.Forms.DialogResult.OK Then
                Select Case frmUpload.FileType
                    Case "Registration"
                        txtFileId.Text = frmUpload.FileId
                        GridTour.CurrentRow.Cells("FileId").Value = frmUpload.FileId
                    Case "Quotation"
                        txtQuotationId.Text = frmUpload.FileId
                        GridTour.CurrentRow.Cells("QuotationFile").Value = frmUpload.FileId
                End Select

            End If
        End If

    End Sub
    

    Private Sub lbkViewUploadedFile_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkViewUploadedFile.LinkClicked
        If txtFileId.Text <> 0 Or txtQuotationId.Text <> 0 Then
            Dim frmView As New frmUploadFile4NonAir(GridTour.CurrentRow, False)
            frmView.ShowDialog()

        End If
    End Sub
    

    Private Sub TabRequiredData_Enter(sender As Object, e As EventArgs) Handles TabRequiredData.Enter
        Dim i As Integer
        Dim blnRqData As Boolean
        Dim tblRequiredData As System.Data.DataTable
        Dim colAvailData As New Collection
        Dim arrSlashBreaks As String()
        Dim arrRequiredData As String()

        flpRequiredData.Controls.Clear()
        If GridTour.CurrentRow.Cells("Contact").Value <> "ZPERSONAL" Then
            blnRqData = True
        End If

        If GridTour.CurrentRow.Cells("RequiredData").Value <> "" Then


            arrRequiredData = GridTour.CurrentRow.Cells("RequiredData").Value.Split("|")
            For i = 0 To arrRequiredData.Length - 1
                arrSlashBreaks = Split(arrRequiredData(i), "/", 1)
                colAvailData.Add(arrSlashBreaks(1), arrSlashBreaks(0))
            Next
        End If

        If blnRqData Then
            tblRequiredData = GetDataTable("Select *,'' as CurrentValue from CWT.dbo.Go_RequiredData where Status='OK'" _
                                    & " and CustId=" & GridTour.CurrentRow.Cells("CustId").Value, Conn)

            If tblRequiredData.Rows.Count > 0 Then
                AddData(tblRequiredData, colAvailData)
            End If
        End If

    End Sub

    Private Sub lbkSaveRequiredData_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkSaveRequiredData.LinkClicked
        Dim lstRequiredData As New List(Of String)

        If Not CheckInputValues4RequiredData() Then
            Exit Sub
        End If
        For Each objCtrl As Control In flpRequiredData.Controls
            If objCtrl.Name = "ucRqDataCombo" Then
                Dim objUc As ucRqDataCombo = objCtrl
                If objUc.cboValue.Text <> "" Then
                    lstRequiredData.Add(objUc.Tag & "|" & objUc.cboValue.Text)
                End If

            ElseIf objCtrl.Name = "ucRqDataText" Then
                Dim objUc As ucRqDataText = objCtrl
                If objUc.txtValue.Text <> "" Then
                    lstRequiredData.Add(objUc.Tag & "|" & objUc.txtValue.Text)
                End If
            End If
        Next
        If lstRequiredData.Count > 0 Then
            If Not ExecuteNonQuerry("Update DuToan_Tour set RequiredData='" _
                             & Join(lstRequiredData.ToArray, "|") & "' where RecId=" _
                         & GridTour.CurrentRow.Cells("RecId").Value, Conn) Then
                MsgBox("Unable to update RequiredData")
            End If
        End If

    End Sub
    Private Function CheckInputValues4RequiredData() As Boolean
        Dim lstError As New List(Of String)
        Dim strError As String = ""
        Dim strErrMsg As String

        Dim i As New List(Of String)

        For Each objCtrl As Control In flpRequiredData.Controls
            With objCtrl
                If .Name = "ucRqDataCombo" Then
                    Dim ucRd As ucRqDataCombo = objCtrl
                    strError = ucRd.CheckInput

                ElseIf .Name = "ucRqDataText" Then
                    Dim ucRd As ucRqDataText = objCtrl
                    strError = ucRd.CheckInput
                End If
            End With
            If strError <> "" Then
                lstError.Add(strError)
            End If
        Next
        strErrMsg = Join(lstError.ToArray, vbNewLine)
        If strErrMsg <> "" Then
            MsgBox(strErrMsg)
            Return False
        End If

        Return True
    End Function
    Private Sub AddData(tblRequiredData As DataTable, colAvailData As Collection)
        For Each objRow As DataRow In tblRequiredData.Rows

            If objRow("CheckValues") Then
                Dim ucRqData As New ucRqDataCombo
                ucRqData.Tag = objRow("DataCode")
                ucRqData.Row = objRow
                If objRow("Mandatory") = "M" Then
                    ucRqData.lblName.Text = objRow("NameByCustomer") & "(*)"
                ElseIf objRow("Mandatory") = "C" Then
                    ucRqData.lblName.Text = objRow("NameByCustomer") & "(" _
                                            & objRow("ConditionOfUse") & ")"
                End If
                LoadComboDisplay(ucRqData.cboValue, "Select Value" _
                        & ",Description as Display from cwt.dbo.GO_RequiredDataValues" _
                        & " where Status='OK' and CustId=" & objRow("CustId") _
                        & " and DataCode='" & objRow("DataCode") & "'", Conn)
                ucRqData.cboValue.SelectedIndex = -1
                If colAvailData.Contains(objRow("DataCode")) Then
                    ucRqData.cboValue.SelectedIndex = ucRqData.cboValue.FindStringExact(colAvailData(objRow("DataCode")))
                End If
                flpRequiredData.Controls.Add(ucRqData)
            Else
                Dim ucRqData As New ucRqDataText
                ucRqData.Tag = objRow("DataCode")
                ucRqData.Row = objRow
                If objRow("Mandatory") = "M" Then
                    ucRqData.lblName.Text = objRow("NameByCustomer") & "(*)"
                ElseIf objRow("Mandatory") = "C" Then
                    ucRqData.lblName.Text = objRow("NameByCustomer") & "(" _
                                            & objRow("ConditionOfUse") & ")"
                End If
                If colAvailData.Contains(objRow("DataCode")) Then
                    ucRqData.txtValue.Text = colAvailData(objRow("DataCode"))
                End If

                flpRequiredData.Controls.Add(ucRqData)
            End If
        Next
    End Sub

    Private Sub cboChannel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboChannel.SelectedIndexChanged
        Select Case cboChannel.Text
            Case "CWT"
                LoadCmb_VAL(Me.cboCustShortName, MyCust.List_CS)
            Case "LCL"
                LoadCmb_VAL(Me.cboCustShortName, MyCust.List_LC)
        End Select
        cboCustShortName.SelectedIndex = -1
    End Sub

    Private Sub blkReset_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkReset.LinkClicked
        Reset()
    End Sub

    
    Private Sub cboCustShortName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCustShortName.SelectedIndexChanged
        If mblnFirstLoadCompleted Then
            LoadgridTour()
        End If
    End Sub

    Private Sub GridGO_SelectionChanged(sender As Object, e As EventArgs) Handles GridGO.SelectionChanged
        With GridGO
            If .RowCount = 0 Then
                mintSelectedCrdRow = -1
            Else
                mintSelectedCrdRow = .CurrentRow.Index
                cboCdrValues.Enabled = .CurrentRow.Cells("CheckValues").Value
                lbkSelectCdrValues.Enabled = .CurrentRow.Cells("CheckValues").Value

                If .CurrentRow.Cells("CheckValues").Value Then
                    Dim strQuerry As String = ""
                    Select Case .CurrentRow.Cells("Source").Value
                        Case "CDR"
                            strQuerry = "Select Value from Cwt.dbo.GO_MiscWzDate" _
                                & " where Status='OK' and Catergory='CDR" & .CurrentRow.Cells("DataCode").Value _
                                & "' and Value1= (Select top 1 CMC from CompanyInfo where Status='ok'" _
                                & " and CustId=" & CmbCust.SelectedValue & ") order by Value"
                            LoadCombo(cboCdrValues, strQuerry, Conn)

                        Case "RequiredData"
                            strQuerry = "Select Value from Cwt.dbo.GO_RequiredDataValues" _
                                & " where Status='OK' and DataCode='" & .CurrentRow.Cells("DataCode").Value _
                                & "' and CustId =" & CmbCust.SelectedValue & " order by Value"
                            LoadCombo(cboCdrValues, strQuerry, Conn)
                    End Select
                End If
            End If
        End With

    End Sub

    Private Sub lbkSelectCdrValues_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkSelectCdrValues.LinkClicked
        If mintSelectedCrdRow > -1 AndAlso cboCdrValues.Text <> "" Then
            GridGO.Rows(mintSelectedCrdRow).Cells("FieldValues").Value = cboCdrValues.Text
        End If
    End Sub

    Private Sub lbkQuickRef_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkQuickRef.LinkClicked
        Dim frmQuickRef As New frmQuickRef(CmbCust.SelectedValue)
        frmQuickRef.ShowDialog()
    End Sub

    Private Sub lbkLinkItems_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkLinkItems.LinkClicked
        If GridSVC.CurrentRow Is Nothing Then Exit Sub
        Dim frmLink As New frmLinkTourItems(GridSVC.CurrentRow)
        frmLink.ShowDialog()
    End Sub

    
    
End Class

