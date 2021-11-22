Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.MySharedFunctionsWzConn
Imports System.IO
Public Class UNC
    Private CharEntered As Boolean = False
    Private Recno As Integer, SelectedCurr As String
    Private iCounter As Int16 = 0
    Private PayerAccouID_frm As Integer, PayerAccouID_To As Integer
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub UNC_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub UNC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If HasNewerVersion_R12(Application.ProductVersion) Or SysDateIsWrong(Conn) Then
            Me.Close()
            Me.Dispose()
            End
        End If
        GenCmbValue()
        LoadGridRQ("")
        LoadGridNewSupplier()
        Dim dtable As DataTable = GetDataTable("select distinct TCode as val from dutoan_Tour where custid>0 and status<>'XX' and SDate >dateadd(d,-128,getdate()) " & _
                                                "Union select VAL from misc where cat='UNC_WHO'")
        For i As Int16 = 0 To dtable.Rows.Count - 1
            Me.LstTRX_TC.Items.Add(dtable.Rows(i)("VAL"))
        Next
        If myStaff.City = "SGN" Then
            PayerAccouID_frm = 1140
            PayerAccouID_To = 1145
        Else
            PayerAccouID_frm = 4182
            PayerAccouID_To = 4187
        End If

    End Sub
    Private Sub GenCmbValue()
        LoadCmb_VAL(Me.CmbTVC, "select RecID as VAL, shortName as DIs from UNC_Company where CAT='TV'")
        LoadCmb_MSC(Me.CmbCurr, "Curr")
        Me.CmbCurr.Text = "VND"
        Me.CmbTVC.SelectedValue = 20
    End Sub
    Private Sub LoadGridRQ(pDK As String)
        Dim strSQL = "select RecID, pmtLotID, Shortname, Curr, Amount, FOP as RQ4, TRX_TC, AccountName, AccountNumber, " & _
            "BankName, BankAddress, Description, FstUpdate, FstUser, RMK, App, SupplierID " & _
            "from unc_payments where "
        If Me.ChkShowLstAccepted.Checked Then
            strSQL = strSQL & " Status='OK' and LstUpdate >'" & Format(Now.AddDays(-8), "dd-MMM-yy") & "'"
        Else
            strSQL = strSQL & " status='QQ'"
            strSQL = strSQL & " and (fop in ('CRD','CSH') or AccountNumber<>'')"
        End If
        strSQL = strSQL & pDK & " order by recid desc"
        Me.GridRQ.DataSource = GetDataTable(strSQL)
        Me.GridRQ.Columns(0).Visible = False
        Me.GridRQ.Columns(1).Visible = False
        Me.GridRQ.Columns("TRX_TC").Width = 128
        Me.GridRQ.Columns("ShortName").Width = 128
        Me.GridRQ.Columns("RQ4").Width = 36
        Me.GridRQ.Columns("Curr").Width = 36
        Me.GridRQ.Columns("FstUser").Width = 36
        Me.GridRQ.Columns("pmtLotID").Width = 56
        Me.GridRQ.Columns("SupplierID").Width = 56
        Me.GridRQ.Columns("App").Width = 56
        Me.GridRQ.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridRQ.Columns("Amount").DefaultCellStyle.Format = "#,##0.00"
        Me.LblAddPayerDetail.Visible = False
        Me.LblCash.Visible = False
        Me.LblCRD.Visible = False
        Me.LblRejectRQ.Visible = False
        Me.LblChangeToVND.Visible = False
        Me.LckLblUndoAccept.Visible = False
        Me.txtNewDescr.Enabled = False
        Me.txtNewAmt.Enabled = False
        If pDK = "" Then
            Me.LblCashAdv2Guide.Visible = False
        ElseIf Me.GridRQ.RowCount > 0 And CInt(Me.txtPmtLotID.Text) > 32000 Then
            Me.LblCashAdv2Guide.Visible = True
        End If
    End Sub
    Private Sub LoadGridNewSupplier()
        Dim strSQL = "select RecID, Curr, Amount, FOP, TRX_TC, Description, RefNo, FstUpdate, FstUser, SupplierID " & _
            " from unc_payments where status='QQ' and SupplierID >0 and accountNumber=''  " & _
            " order by recid desc"
        Me.GridNewSupplier.DataSource = GetDataTable(strSQL)
        Me.GridNewSupplier.Columns(0).Visible = False
        Me.GridNewSupplier.Columns("TRX_TC").Width = 128
        Me.GridNewSupplier.Columns("FOP").Width = 36
        Me.GridNewSupplier.Columns("Curr").Width = 36
        Me.GridNewSupplier.Columns("FstUser").Width = 36
        Me.GridNewSupplier.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridNewSupplier.Columns("Amount").DefaultCellStyle.Format = "#,##0.00"
        Me.LblAssignAccount.Visible = False
        Me.LblGo2Mapping.Visible = False
    End Sub

    Private Sub OptAL_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        OptAL.Click, OptOT.Click, OptLD.Click, OptTO.Click, OptZZ.Click, OptUS.Click, _
        OptXE.Click, OptNH.Click, OptKS.Click, OptAR.Click, OptOS.Click, OptBO.Click, OptAP.Click, OptYT.Click
        Dim tmp As String = "", opt As RadioButton = CType(sender, RadioButton)
        tmp = opt.Name.Substring(3, 2)
        GenBeneficiary(tmp)
        If tmp <> "ZZ" Then
            Me.CmbBeneficiary.Enabled = True
        Else
            Me.CmbBeneficiary.Enabled = False
            ClearTextBox()
        End If
        Me.GridBeneficiary.Visible = Me.CmbBeneficiary.Enabled
        Me.TxtAccountName.Enabled = Not Me.CmbBeneficiary.Enabled
        Me.TxtAccNo.Enabled = Not Me.CmbBeneficiary.Enabled
        Me.TxtBank.Enabled = Not Me.CmbBeneficiary.Enabled
        Me.TxtCity.Enabled = Not Me.CmbBeneficiary.Enabled
        Me.TxtSwift.Enabled = Not Me.CmbBeneficiary.Enabled
    End Sub
    Private Sub GenBeneficiary(ByVal pCat As String)
        LoadCmb_VAL(Me.CmbBeneficiary, "select RecID as VAL, shortName as DIS from UNC_Company where status ='OK' and CAT='" & pCat & "'")
        Me.GridCanText.DataSource = GetDataTable("select VAL from MISC where cat='CANTXT' and VAL1='" & pCat & "'")
        Me.GridCanText.Columns(0).Width = 320
        Me.TxtDescript.Text = ""
    End Sub

    Private Sub CmbTVC_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbTVC.SelectedIndexChanged
        Dim Cmd As SqlClient.SqlCommand = Conn.CreateCommand
        On Error GoTo ErrHandler
        Dim strSQL As String = "select RecID, LstUser as Bank, AccountNumber from UNC_Accounts where status='OK' and CompanyID=" & _
            Me.CmbTVC.SelectedValue & " order by LstUser, AccountNumber"
        Me.GridPayer.DataSource = GetDataTable(strSQL)
        Me.GridPayer.Columns(0).Visible = False
        Me.GridPayer.Columns(1).Width = 64
        Me.GridPayer.Columns(2).Width = 256
        Me.GrpBeneficiaryDetail.Enabled = False
        ClearTextBox()

        If Me.CmbTVC.Text = "TVTR" Then
            Me.CmbInvNo.Text = ""
            Me.CmbPmtType.Visible = True
            Me.CmbPmtType.Text = Me.CmbPmtType.Items(0).ToString
            Me.GridRQ.Visible = True
        Else
            Me.GridRQ.Visible = False
            Me.CmbPmtType.Visible = False
            Me.CmbPmtType.Text = ""
            Me.CmbInvNo.Text = "N/A"
        End If
        Me.LstTRX_TC.Visible = Me.CmbPmtType.Visible
        Me.Label10.Visible = Me.CmbPmtType.Visible
ErrHandler:
        On Error GoTo 0
    End Sub

    Private Sub TxtAmt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAmt.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub

    Private Sub TxtAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAmt.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub GridCanText_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCanText.CellContentClick
        Me.TxtDescript.Text = Me.GridCanText.CurrentCell.Value & " "
        Me.TxtDescript.Focus()
        Me.TxtDescript.SelectionStart = Me.TxtDescript.TextLength
    End Sub

    Private Sub CmbBeneficiary_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbBeneficiary.SelectedIndexChanged
        On Error GoTo ErrHandler
        Dim strSQL As String = "select RecID, AccountName,AccountNumber, BankName, BankAddress,swift  from UNC_Accounts " & _
            " where status ='OK' and CompanyID=" & Me.CmbBeneficiary.SelectedValue & " order by BankName, AccountNumber"
        Me.GridBeneficiary.DataSource = GetDataTable(strSQL)
        Me.GridBeneficiary.Columns(0).Visible = False
        Me.GridBeneficiary.Columns(1).Width = 200
        Me.GridBeneficiary.Columns(2).Width = 128
        Me.GridBeneficiary.Columns(3).Visible = False
        Me.GridBeneficiary.Columns(4).Visible = False
        Me.GridBeneficiary.Columns(5).Width = 192
        Me.Label12.Text = ScalarToString("UNC_company", "FOP", "RecID=" & Me.CmbBeneficiary.SelectedValue)
        ClearTextBox()
        Exit Sub
ErrHandler:
        On Error GoTo 0
    End Sub

    Private Sub GridBeneficiary_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles _
        GridBeneficiary.CellClick
        Me.TxtBank.Text = Me.GridBeneficiary.CurrentRow.Cells("BankName").Value
        Me.TxtCity.Text = Me.GridBeneficiary.CurrentRow.Cells("BankAddress").Value
        Me.TxtAccNo.Text = Me.GridBeneficiary.CurrentRow.Cells("AccountNumber").Value
        Me.TxtAccountName.Text = Me.GridBeneficiary.CurrentRow.Cells("AccountName").Value
        Me.TxtSwift.Text = Me.GridBeneficiary.CurrentRow.Cells("Swift").Value
        Dim LstPmt As Date = ScalarToDate("UNC_payments", "top 1 FstUpdate", "AccountNumber='" & Me.TxtAccNo.Text & _
                                            "' order by fstupdate desc")
        If LstPmt < Now.AddMonths(-3) Then
            MsgBox("This Account Seems to Be Asleep. Plz Check its Validity", MsgBoxStyle.Exclamation, msgTitle)
        End If
    End Sub
    Private Sub GenTemplate()
        Dim strSQL As String = "select VAL from misc where cat='UNC' and left(val,3)='" & Me.GridPayer.CurrentRow.Cells("Bank").Value & "'"
        LoadCmb_MSC(Me.CmbTemplate, strSQL)
    End Sub

    Private Sub GridPayer_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles _
        GridPayer.CellContentClick, GridPayer.CellClick
        If e.RowIndex < 0 Then Exit Sub
        iCounter = 0
        GenTemplate()
        If Me.GridPayer.CurrentRow.Cells("Bank").Value = "TCB" Then
            If InStr(Me.GridPayer.CurrentRow.Cells("AccountNumber").Value, "VND") > 0 Then
                Me.CmbTemplate.Text = Me.CmbTemplate.Items(1).ToString
            End If
        End If
        Me.GrpBeneficiaryDetail.Enabled = True
        Me.GridRQ.Enabled = True
        For i As Int16 = 0 To Me.GridRQ.RowCount - 1
            If Me.GridRQ.Item("RQ4", i).Value <> "BTF" Then
                Me.GridRQ.Rows(i).DefaultCellStyle.ForeColor = Color.Red
            End If
        Next
    End Sub

    Private Sub TxtAmt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAmt.LostFocus
        Dim aa As Double = CDbl(Me.TxtAmt.Text)
        Me.TxtAmt.Text = Format(aa, "#,##0.00")
        iCounter = 0
    End Sub
    Private Function DefineWhoBearCharge() As String
        Dim KQ As String = ""
        If Me.OptInVNBeneficiary.Checked Then KQ = "IB"
        If Me.OptInVNPayer.Checked Then KQ = KQ & "IP"
        If Me.OptOutVNBeneficiary.Checked Then KQ = KQ & "OB"
        If Me.OptOutVNPayer.Checked Then KQ = KQ & "OP"
        Return KQ
    End Function

    Private Sub LblSave_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSave.LinkClicked
        If Me.TxtAccountName.Text = "" Or Me.TxtAccNo.Text = "" Or Me.TxtDescript.Text = "" Or _
           Me.TxtBank.Text = "" Or CDec(Me.TxtAmt.Text) = 0 Or Me.TxtCity.Text = "" Or _
           Me.TxtAccountName.Text = "ACCOUNT NAME" Or Me.TxtAccNo.Text = "ACCOUNT NO" Or _
           Me.TxtBank.Text = "BANK" Or Me.TxtCity.Text = "LOCATION" Then
            MsgBox("Invalid Input!", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        Me.LblSave.Visible = False
        Dim Cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim strSQL As String, tmpNoiDung As String
        Dim t As SqlClient.SqlTransaction
        Dim tmpCharge As String = "", RefNo As String, TRX_TC As String = ""
        tmpNoiDung = Me.TxtDescript.Text.Replace("'", "").Replace("--", "")
        tmpCharge = DefineWhoBearCharge()
        RefNo = DefineRefNo_UNC(Me.CmbTVC.Text)
        For i As Int16 = 0 To Me.LstTRX_TC.Items.Count - 1
            If Me.LstTRX_TC.GetItemChecked(i) Then
                TRX_TC = TRX_TC & "_" & Me.LstTRX_TC.Items(i).ToString
            End If
        Next
        If TRX_TC.Length > 2 Then
            TRX_TC = Me.CmbPmtType.Text & ":" & TRX_TC.Substring(1)
        End If
        t = Conn.BeginTransaction
        Cmd.Transaction = t
        strSQL = "Insert into UNC_Payments (PayerAccountID, AccountName, AccountNumber, BankName, BankAddress, " & _
            " Curr, Amount, Description, swift, Charge, shortname, InvNo, TRX_TC, PayeeAccountID, RMKNoiBo, FstUser) values ('" & _
            Me.GridPayer.CurrentRow.Cells("Recid").Value & "',N'" & _
            Me.TxtAccountName.Text & "','" & _
            Me.TxtAccNo.Text & "',N'" & _
            Me.TxtBank.Text & "',N'" & _
            Me.TxtCity.Text & "','" & _
            Me.CmbCurr.Text & "','" & _
            CDec(Me.TxtAmt.Text) & "',N'" & _
            tmpNoiDung & "','" & _
            Me.TxtSwift.Text & "','" & _
            tmpCharge & "',N'" & _
            Me.CmbBeneficiary.Text & "','" & _
            Me.CmbInvNo.Text & "','" & _
            TRX_TC & "',"
        If Me.OptZZ.Checked Then
            strSQL = strSQL & "0,'"
        Else
            strSQL = strSQL & Me.GridBeneficiary.CurrentRow.Cells("RecID").Value & ",'"
        End If
        strSQL = strSQL & tmpNoiDung & "','" & myStaff.SICode & "')"
        Cmd.CommandText = strSQL & "; SELECT SCOPE_IDENTITY() AS [RecID]"
        On Error GoTo ErrorHandler
        Recno = Cmd.ExecuteScalar
        Cmd.CommandText = "update UNC_Payments set refno='" & Replace(RefNo, "--", "") & "' where recid =" & Recno
        Cmd.ExecuteNonQuery()
        t.Commit()
        Me.LblPreview.Visible = True
        Exit Sub
ErrorHandler:
        t.Rollback()
        MsgBox("Unable Writting into DataBase. Please Check Your Input or If You Ensure the Correct Input, Contact Supervisor", MsgBoxStyle.Critical, msgTitle)
        On Error GoTo 0
    End Sub

    Private Sub LblNew_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblNew.LinkClicked
        Me.LblSave.Visible = True
        Me.LblPreview.Visible = False
        ClearTextBox()
    End Sub

    Private Sub LblPreview_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblPreview.LinkClicked
        Dim intPayerAccountId As Integer = ScalarToInt("UNC_payments", "PayerAccountID", "RecID=" & Recno)

        If Me.CmbTemplate.Text.Contains("SCB") And _
             intPayerAccountId > PayerAccouID_frm And _
            intPayerAccountId < PayerAccouID_To Then
            InHoaDon(Application.StartupPath, Me.CmbTemplate.Text, "V", Recno, Now.Date, Now.Date, 0, "NEW", MySession.Domain)

            If MsgBox("Wanna Print it Out?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle) = vbNo Then Exit Sub

            InHoaDon(Application.StartupPath, Me.CmbTemplate.Text, "O", Recno, Now.Date, Now.Date, 0, "NEW", MySession.Domain)

            Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
            cmd.CommandText = "update UNC_payments set HasPrinted=1 where recid=" & Recno
            cmd.ExecuteNonQuery()

        Else
            InHoaDon(Application.StartupPath, Me.CmbTemplate.Text, "O", Recno, Now.Date, Now.Date, 0, "NEW", MySession.Domain)
        End If

        Dim strBankName As String = ScalarToString("UNC_Accounts", "BankName", "RecID=" & intPayerAccountId)
        If strBankName.Contains("CITIBANK") Or strBankName.Contains("Standard Chartered Bank") Then
            CreateSms4UNC(Recno, strBankName, False)
        End If

    End Sub

    Private Sub ClearTextBox()
        Me.TxtAccountName.Text = "ACCOUNT NAME"
        Me.TxtAccNo.Text = "ACCOUNT NO"
        Me.TxtDescript.Text = ""
        Me.TxtBank.Text = "BANK"
        Me.TxtSwift.Text = "SWIFT"
        Me.TxtCity.Text = "LOCATION"
        Me.TxtAmt.Text = 0
        For i As Int16 = 0 To Me.LstTRX_TC.Items.Count - 1
            Me.LstTRX_TC.SetItemChecked(i, False)
        Next
    End Sub
    Private Sub LblAddPayerDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAddPayerDetail.LinkClicked
        AcceptRQ("BTF", Me.GridRQ.CurrentRow.Cells("recID").Value, True)
        Me.LblPreview.Visible = True
    End Sub
    Private Sub AcceptRQ(pFOP As String, RQID As Integer, NeedReload As Boolean)
        Dim RefNo As String = "", PayerID As Integer = 0
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        If pFOP = "BTF" Then
            RefNo = DefineRefNo_UNC("TVTR")
            PayerID = Me.GridPayer.CurrentRow.Cells("RecID").Value
        End If
        cmd.CommandText = "update UNC_Payments set status='OK', Refno='" & RefNo & "', PayerAccountID= " & PayerID & _
            ", FOP='" & pFOP & "',FstUpdate=getdate(), LstUpdate=getdate(), LstUser='" & myStaff.SICode & "', Description=@Description" & _
            ", Amount=" & CDec(Me.txtNewAmt.Text) & " where recid=" & RQID
        cmd.Parameters.Add("@Description", SqlDbType.VarChar).Value = Me.txtNewDescr.Text
        cmd.ExecuteNonQuery()
        Recno = RQID
        If NeedReload Then LoadGridRQ("")
        Me.GroupBox1.Enabled = False
        Me.txtNewDescr.Enabled = False
        Me.txtNewAmt.Enabled = False
    End Sub
    Private Sub GridRQ_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridRQ.CellClick
        If e.RowIndex < 0 Then Exit Sub
        iCounter = 0
        Me.LblAddPayerDetail.Visible = Not Me.ChkShowLstAccepted.Checked
        Me.LblRejectRQ.Visible = Not Me.ChkShowLstAccepted.Checked
        Me.LblCash.Visible = Not Me.ChkShowLstAccepted.Checked
        Me.LblCRD.Visible = Not Me.ChkShowLstAccepted.Checked
        Me.LckLblUndoAccept.Visible = Me.ChkShowLstAccepted.Checked
        Me.txtNewDescr.Text = Me.GridRQ.CurrentRow.Cells("Description").Value
        Me.txtNewAmt.Text = Format(Me.GridRQ.CurrentRow.Cells("Amount").Value, "#,##0.00")
        If Me.GridRQ.CurrentRow.Cells("Curr").Value <> "VND" Then
            Me.LblChangeToVND.Visible = True
            Me.TxtROE.Text = ForEX_12(Now, Me.GridRQ.CurrentRow.Cells("Curr").Value, "BSR", "YY").Amount
        End If
        Me.GroupBox1.Enabled = True
        Me.txtPmtLotID.Text = -1
        Me.LblCashAdv2Guide.Visible = False
        Me.txtNewDescr.Enabled = False
        If Me.GridRQ.CurrentRow.Cells("supplierID").Value <> 0 Then
            LblCreditType.Text = "CreditType: " & ScalarToString("UNC_Company", "FOP", " RecID in (select VendorID from supplier where RecID=" &
                                        Me.GridRQ.CurrentRow.Cells("supplierID").Value & ")")

        End If
    End Sub
    Private Function DefineRefNo_UNC(pCongTy As String) As String
        Dim tmpCty As String = Strings.Right(pCongTy, 2) & Format(Now, "MMyy") & "-"
        Dim KQ As String = ScalarToString("UNC_payments", "top 1 RefNo", "left(refno,7)='" & tmpCty & "' order by refno desc")
        If KQ <> "" Then
            KQ = tmpCty & Format(CInt(KQ.Substring(7, 3)) + 1, "000")
        Else
            KQ = tmpCty & "001"
        End If
        Return KQ
    End Function
    Private Sub RejectRQ()
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim ServerName As String = IIf(myStaff.City = "SGN", "[42.117.5.86].", "")
        Dim DaTra As Decimal, pmtLotID As Integer = Me.GridRQ.CurrentRow.Cells("pmtLotID").Value
        If pmtLotID = 0 Then
            cmd.CommandText = ChangeStatus_ByID("UNC_payments", "XX", Me.GridRQ.CurrentRow.Cells("recID").Value)
        Else
            cmd.CommandText = ChangeStatus_ByDK("UNC_payments", "XX", " PmtLotID=" & Me.GridRQ.CurrentRow.Cells("pmtLotID").Value)
        End If
        If Me.GridRQ.CurrentRow.Cells("App").Value = "RAS" Then
            cmd.CommandText = cmd.CommandText & ";" & ChangeStatus_ByDK("dutoan_pmt", "XX", "PmtID=" & Me.GridRQ.CurrentRow.Cells("recID").Value)
            cmd.ExecuteNonQuery()

            Dim tbPmt As DataTable = GetDataTable("select ItemID from dutoan_Pmt where pmtID=" & Me.GridRQ.CurrentRow.Cells("recID").Value)
            cmd.CommandText = ""
            For i As Int16 = 0 To tbPmt.Rows.Count - 1
                DaTra = ScalarToDec("dutoan_pmt", "isnull(sum(VND),0)", "ItemID=" & tbPmt.Rows(i)("ItemID") & " and status<>'XX'")
                cmd.CommandText = cmd.CommandText & "; Update Dutoan_item set VNDPaid=" & DaTra & _
                    " where RecID=" & tbPmt.Rows(i)("ItemID")
            Next
            If cmd.CommandText.Length > 2 Then
                cmd.CommandText = cmd.CommandText.Substring(1)
                cmd.ExecuteNonQuery()
            End If
        Else
            If pmtLotID = 0 Then
                cmd.CommandText = cmd.CommandText & "; update " & ServerName & "flx.dbo.tos_Pmt set status='XX' where DNTTID =" & _
                    Me.GridRQ.CurrentRow.Cells("recID").Value & " and city='" & myStaff.City & "'"
            Else
                For r As Int16 = 0 To Me.GridRQ.RowCount - 1
                    If Me.GridRQ.Item("PmtLotID", r).Value = pmtLotID Then
                        cmd.CommandText = cmd.CommandText & "; update " & ServerName & "flx.dbo.tos_Pmt set status='XX' where DNTTID =" & _
                            Me.GridRQ.Item("recID", r).Value & " and city='" & myStaff.City & "'"
                    End If
                Next
            End If
            cmd.ExecuteNonQuery()
        End If
        LoadGridRQ("")
    End Sub
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblRejectRQ.LinkClicked
        RejectRQ()
    End Sub

    Private Sub LblCash_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCash.LinkClicked
        AcceptRQ("CSH", Me.GridRQ.CurrentRow.Cells("recID").Value, True)
    End Sub

    Private Sub LblCRD_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCRD.LinkClicked
        AcceptRQ("CRD", Me.GridRQ.CurrentRow.Cells("recID").Value, True)
    End Sub

    Private Sub LblGo2Mapping_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblGo2Mapping.LinkClicked
        Dim SupplierID As Integer = Me.GridNewSupplier.CurrentRow.Cells("SupplierID").Value
        Dim f As New Vendor_SupplierMapping(SupplierID)
        f.ShowDialog()
    End Sub

    Private Sub AssignAccountToPmtRQ(pPmtRQID As Integer, pSupplierID As Integer, pAccountID As Integer)
        Dim VendorID As Integer, dtblPayee As DataTable
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        VendorID = ScalarToInt("supplier", "VendorID", "RecID=" & pSupplierID)
        Dim VendorName As String = ScalarToString("UNC_Company", "shortName", "RecID=" & VendorID)
        cmd.CommandText = "update UNC_Payments set ShortName=@ShortName where recID=" & pPmtRQID
        cmd.Parameters.AddWithValue("@ShortName", VendorName)
        cmd.ExecuteNonQuery()
        If pAccountID = 0 Then
            dtblPayee = GetDataTable("select top 1 * from unc_accounts where companyID=" & VendorID & " and status='OK'")
        Else
            dtblPayee = GetDataTable("select * from unc_accounts where RecID=" & pAccountID)
        End If
        If dtblPayee.Rows.Count = 1 Then
            cmd.CommandText = "update UNC_Payments set AccountName =@AccountName, AccountNumber =@AccountNumber, " & _
                "BankName =@BankName , BankAddress =@BankAddress , Swift=@Swift, PayeeAccountID=@PayeeAccID " & _
                "where recID=" & pPmtRQID
            cmd.Parameters.Clear()
            cmd.Parameters.Add("@AccountName", SqlDbType.VarChar).Value = dtblPayee.Rows(0)("AccountName")
            cmd.Parameters.Add("@AccountNumber", SqlDbType.VarChar).Value = dtblPayee.Rows(0)("AccountNumber")
            cmd.Parameters.Add("@BankName", SqlDbType.VarChar).Value = dtblPayee.Rows(0)("BankName")
            cmd.Parameters.Add("@BankAddress", SqlDbType.NVarChar).Value = dtblPayee.Rows(0)("BankAddress")
            cmd.Parameters.Add("@Swift", SqlDbType.VarChar).Value = dtblPayee.Rows(0)("Swift")
            cmd.Parameters.Add("@PayeeAccID", SqlDbType.Int).Value = dtblPayee.Rows(0)("RecID")
            cmd.ExecuteNonQuery()
        End If
    End Sub
    Private Sub LblAssignAccount_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAssignAccount.LinkClicked
        AssignAccountToPmtRQ(Me.GridNewSupplier.CurrentRow.Cells("RecID").Value, Me.GridNewSupplier.CurrentRow.Cells("SupplierID").Value, 0)
        LoadGridRQ("")
        LoadGridNewSupplier()
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridNewSupplier.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LblAssignAccount.Visible = True
        Me.LblGo2Mapping.Visible = True
    End Sub

    Private Sub LblChangeToVND_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblChangeToVND.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim VNDEquiv As Decimal = Me.GridRQ.CurrentRow.Cells("Amount").Value * CDec(Me.TxtROE.Text)
        VNDEquiv = Math.Round(VNDEquiv, 0)
        If myStaff.City = "SGN" AndAlso Me.GridRQ.CurrentRow.Cells("AccountName").Value <> "" Then
            Dim CustomeRounding As String = ScalarToString("MISC", "VAL1", "cat='CUSTRNDG' and charindex(VAL,'" & _
                                                           Me.GridRQ.CurrentRow.Cells("AccountName").Value & "')<>0 ")
            If Not String.IsNullOrEmpty(CustomeRounding) Then
                Dim strVND As String = VNDEquiv.ToString
                Dim HangNgan As String = Strings.Left(strVND, strVND.Length - 3)
                Dim PhanTram As String = Strings.Right(strVND, 3)
                If CInt(PhanTram) > CInt(CustomeRounding) Then
                    VNDEquiv = (CDec(HangNgan) + 1) * 1000
                Else
                    VNDEquiv = CDec(HangNgan) * 1000
                End If
            End If
        End If
        cmd.CommandText = UpdateLogFile("UNC_payments", "Change2VND", Me.TxtROE.Text, Me.GridRQ.CurrentRow.Cells("RecID").Value, "", "", "") & _
            "; update UNC_Payments set Curr='VND', Amount=" & VNDEquiv & " Where RecID=" & _
            Me.GridRQ.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridRQ("")
    End Sub
    Private Sub ChkShowLstAccepted_CheckedChanged(sender As Object, e As EventArgs) Handles ChkShowLstAccepted.CheckedChanged
        LoadGridRQ("")
    End Sub
    Private Sub LblUndoAccept_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LckLblUndoAccept.LinkClicked
        If (InStr("SYS_NMH_HXT", myStaff.SICode) > 0 And myStaff.City = "SGN") Or _
            (InStr("SYS_NTH", myStaff.SICode) > 0 And myStaff.City = "HAN") Then
            RejectRQ()
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        iCounter = iCounter + 1
        If iCounter = 2 And myStaff.City = "SGN" Then Me.Close()
    End Sub
    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        iCounter = -4
    End Sub

    Private Sub LblLoadByLotID_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblLoadByLotID.LinkClicked
        If CInt(Me.txtPmtLotID.Text) < 32000 Then Exit Sub
        LoadGridRQ(" and PmtLotID=" & CInt(Me.txtPmtLotID.Text))
    End Sub

    Private Sub LblCashAdv2Guide_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCashAdv2Guide.LinkClicked
        If Me.GridRQ.RowCount = 0 Then Exit Sub
        Dim tmpCurr As String = Me.GridRQ.Item("Curr", 0).Value
        For i As Int16 = 1 To Me.GridRQ.RowCount - 1
            If Me.GridRQ.Item("Curr", i).Value <> tmpCurr Then Exit Sub
        Next
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update UNC_Payments set status='OK', FOP='CSH', FstUpdate=getdate(), LstUpdate=getdate(), LstUser='" & _
            myStaff.SICode & "' where pmtLotID<>0 and PmtLotID=" & CInt(Me.txtPmtLotID.Text)
        cmd.ExecuteNonQuery()
        LoadGridRQ("")
        Me.GroupBox1.Enabled = False
        Me.txtNewDescr.Enabled = False
        Me.txtNewAmt.Enabled = False
    End Sub
    Private Sub Label11_DoubleClick(sender As Object, e As EventArgs) Handles Label11.DoubleClick
        Me.txtNewDescr.Enabled = True
    End Sub
    Private Sub txtNewDescr_TextChanged(sender As Object, e As EventArgs) Handles txtNewDescr.TextChanged
        If Me.txtNewDescr.Text.Contains("can tru") Then
            Me.txtNewAmt.Enabled = True
        Else
            Me.txtNewAmt.Enabled = False
            Me.txtNewAmt.Text = Format(Me.GridRQ.CurrentRow.Cells("Amount").Value, "#,##0.00")
        End If
    End Sub

    Private Sub TxtAccountName_Enter(sender As Object, e As EventArgs) Handles TxtAccountName.Enter
        If Me.TxtAccountName.Text = "ACCOUNT NAME" Then Me.TxtAccountName.Text = ""
    End Sub

    Private Sub TxtBank_Enter(sender As Object, e As EventArgs) Handles TxtBank.Enter
        If Me.TxtBank.Text = "BANK" Then Me.TxtBank.Text = ""
    End Sub
    Private Sub TxtAccNo_Enter(sender As Object, e As EventArgs) Handles TxtAccNo.Enter
        If Me.TxtAccNo.Text = "ACCOUNT NO" Then Me.TxtAccNo.Text = ""
    End Sub
    Private Sub TxtSwift_Enter(sender As Object, e As EventArgs) Handles TxtSwift.Enter
        If Me.TxtSwift.Text = "SWIFT" Then Me.TxtSwift.Text = ""
    End Sub
    Private Sub TxtCity_Enter(sender As Object, e As EventArgs) Handles TxtCity.Enter
        If Me.TxtCity.Text = "LOCATION" Then Me.TxtCity.Text = ""
    End Sub

    Private Sub OptAR_CheckedChanged(sender As Object, e As EventArgs) Handles OptAR.CheckedChanged

    End Sub
End Class