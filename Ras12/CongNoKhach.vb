Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.Crd_Ctrl
Imports SharedFunctions.MySharedFunctionsWzConn
Public Class frmDC_Pax
    Private MyCust As New objCustomer
    Private CharEntered As Boolean = False
    Private TiGiaGi As String = "INVOICE"
    Private varAction As String
    Private Const MaxWriteOffValueVND As Decimal = 5000
    Private Const MaxWriteOffValueUSD As Decimal = 1
    Private QryCust As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private iCounter As Int16 = 0
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal parAction As String)
        InitializeComponent()
        varAction = parAction
    End Sub
    Private Sub frmDC_Pax_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        AutoUploadPPD2TransVietVN()
        MyCust.CustID = 0
        Me.Dispose()
    End Sub

    Private Sub frmBO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        If myStaff.City = "SGN" Then Shell("route change 0.0.0.0 mask 0.0.0.0 172.16.1.252")
        If HasNewerVersion_R12(Application.ProductVersion) Then
            Me.Close()
            Me.Dispose()
            End
        End If
        MyCust.GenCustList()
        If InStr(varAction.ToUpper, "INVOIC") > 0 Then
            QryCust = MyCust.List_CC & " and (CRCoef > 0 or adh='PSP')"
            Me.PadAction.Enabled = False
        ElseIf InStr(varAction.ToUpper, "PAYMENT") > 0 Then
            Me.Timer1.Enabled = False
            QryCust = MyCust.List_CC & " and CRCoef + ppCoef > 0 "
        ElseIf InStr(varAction.ToUpper, "FOLLOWUP") > 0 Then
            QryCust = MyCust.List_CC & " and (CRCoef > 0 or adh='PSP')"
            Me.GrpPmtFollowUp.Left = 0
            Me.GrpPmtFollowUp.Width = 780
        End If
        If MySession.Domain = "EDU" Then
            QryCust = QryCust & " and custid <0"
        End If

        InitSetting()
        GenComboValue()
        Me.CmbCustomer.Focus()
        If InStr(varAction, "Invoicing") > 0 Then
            Me.PadFilter.Enabled = False
            Me.PnlInvoicing.Enabled = True
            Me.BarPaymentDate.Enabled = False
            Me.BarIncludeFullyUsed.Enabled = False
            Me.BarDeleteThisPayment.Enabled = False
            Me.ChotNo.Text = "Invoicing"
            Me.BarPurchaseDate.Checked = True
            If InStr(varAction, "Quick") > 0 Then
                Me.CmdQuickInvUSD.Enabled = True
                Me.CmdQuickInvVND.Enabled = True
            Else
                Me.CmdQuickInvUSD.Enabled = False
                Me.CmdQuickInvVND.Enabled = False
            End If
        ElseIf varAction = "ApplyPayments" Then
            Me.BarPaymentDate.Checked = True
            Me.PnlApplyPayments.Enabled = True
            Me.BarInvoiceDate.Enabled = False
            Me.BarPurchaseDate.Enabled = False
            Me.ThanhToan.Text = "Apply Payments"
            Me.TabControl1.SelectTab("ThanhToan")
            Me.CmbPmtType.Text = "PSP"
        ElseIf InStr(varAction.ToUpper, "FOLLOWUP") > 0 Then
            Me.GrpPmtFollowUp.Visible = True
            Me.TabControl1.SelectTab("TabROE")
            Me.BarCustType_CWT.PerformClick()
        End If
        Me.CmbCNCurr.Text = "VND"
        CheckRightForALLForm(Me)
    End Sub
    Private Sub InitSetting()
        Me.ChotNo.BackColor = pubVarBackColor
        Me.ThanhToan.BackColor = pubVarBackColor
        Me.TabROE.BackColor = pubVarBackColor

        Me.ChotNo.Text = ""
        Me.ThanhToan.Text = ""
        Me.GridDungTienNao.Width = 788
    End Sub
    Private Sub GenComboValue()
        Dim StrSQL As String
        LoadCmb_MSC(Me.CmbInvoiceAL, myStaff.TVA)
        Me.CmbInvoiceAL.Items.Add("ALL")
        Me.CmbInvoiceAL.Text = "ALL"

        LoadCmb_MSC(Me.CmbCNFOP, "select VAL from MISC where CAT='FOP' and  VAL2 like '%CC%' ")

        Me.CmbReceivedBy.Items.Clear()
        Me.CmbReceivedBy.Items.Add("BO-BackOFC")
        Me.CmbReceivedBy.Text = Me.CmbReceivedBy.Items(0).ToString
        StrSQL = "select VAL from MISC where CAT='CURR' order by VAL"
        LoadCmb_MSC(Me.CmbInvoiceCurr, StrSQL)
        LoadCmb_MSC(Me.CmbCNCurr, StrSQL)
        LoadCmb_MSC(Me.Cmb1Curr, StrSQL)
        LoadCmb_MSC(Me.CmbEQnCurr, StrSQL)
        Me.CmbCNCurr.Text = "USD"
        Me.CmbInvoiceCurr.Text = "USD"

        RemoveHandler CmbCustomer.SelectedIndexChanged, AddressOf CmbCustomer_SelectedIndexChanged
        LoadCmb_VAL(Me.CmbCustomer, QryCust)
        AddHandler CmbCustomer.SelectedIndexChanged, AddressOf CmbCustomer_SelectedIndexChanged
        Call CmbCustomer_SelectedIndexChanged(CmbCustomer, System.EventArgs.Empty)
    End Sub
    Private Sub GetPurchases()
        Dim strSQL As String, InvoiceUpto As Date
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        strSQL = String.Format("select RCPNo as OrgDocs, SRV, FstUpdate as OrgDate, Currency as OrgCurr, AMount as OrgTTLAmt, " & _
            " Currency as InvCurr, Amount as InvAmt, 1.00 As InvROE, '{0}' as InvDate, '{1}' as DueDate, CustID, RecID, Nguon" & _
            " from func_CC_PSP ({2},'{3}','PSP') where fstupdate >'31-dec-12'", Format(Me.txtInvoiceDate.Value, "dd-MMM-yy"), _
            Format(Me.txtInvoiceDate.Value, "dd-MMM-yy"), Me.CmbCustomer.SelectedValue, CutOverDatePSP)
        If Me.CmbBooker.Text <> "" Then
            strSQL = String.Format("{0} and rcpid in (select recid from rcp where ca='{1}')", strSQL, Me.CmbBooker.Text)
        End If
        Me.GridNo.DataSource = GetDataTable(strSQL)
        Me.GridNo.EditMode = DataGridViewEditMode.EditOnEnter
        ResizeGridNo()
        InvoiceUpto = Format(Me.txtUpto.Value, "dd-MMM-yy") & " 23:59"
        For i As Int16 = 0 To Me.GridNo.RowCount - 1
            If Me.GridNo.Item("OrgDate", i).Value < InvoiceUpto Then
                Me.GridNo.Item("S", i).Value = True
            End If
        Next
    End Sub
    Private Sub InsertIntoGhiNoKhach(ByVal parInvStatus As String)
        Dim invAmt As Decimal, strSQL As String = "", InvID As Integer, DocNo As String
        Try
            For i As Int16 = 0 To Me.GridNo.Rows.Count - 1
                If Me.GridNo.Item("S", i).Value = True Then
                    invAmt = invAmt + CDec(Me.GridNo.Item("InvAmt", i).Value)
                End If
            Next
            InvID = Insert_GhiNoKhach("ID", Me.CmbInvoiceCurr.Text, Me.txtInvoiceDate.Value, _
                invAmt, parInvStatus, "", Me.txtDueDate.Value.Date, 1, MyCust.CustID)

            For i As Int16 = 0 To Me.GridNo.Rows.Count - 1
                If Me.GridNo.Item("S", i).Value Then
                    DocNo = Me.GridNo.Item("OrgDOcs", i).Value.ToString.Substring(0, 2)
                    If DocNo = "FL" Then
                        strSQL = "update FLX_fop "
                    Else
                        strSQL = "update fop "
                    End If
                    cmd.CommandText = String.Format("{0} set profID={1} where RecID={2}", strSQL, InvID, Me.GridNo.Item("RECID", i).Value)
                    cmd.ExecuteNonQuery()
                End If
            Next
            Me.GridNo.DataSource = Nothing
            Dim c As DataGridViewColumn = Me.GridNo.Columns("S").Clone
            Me.GridNo.Columns.Clear()
            Me.GridNo.Columns.Add(c)
            LoadGridNo()
        Catch ex As Exception
            MsgBox("Error Writing to DataBase", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub
    Private Sub LoadGridNo()
        Dim strSQL As String
        strSQL = String.Format("Select RecID, invCurr, InvAmt, invDate, Paid, Note, CustID, CustShortName, DueDate " & _
            " from GhiNoKhach where custID={0}", Me.CmbCustomer.SelectedValue)
        If Me.BarDueOnly.Checked Then
            strSQL = String.Format("{0}  and ConNo<>0 ", strSQL)
        End If
        Me.GridNo.DataSource = GetDataTable(strSQL)
        Me.GridNo.EditMode = DataGridViewEditMode.EditProgrammatically
        ResizeGridNo()
        Me.GridNo.Columns(1).Visible = False
        Me.BarMarkSelectedAsUnpaid.Enabled = False
    End Sub
    Private Sub ResizeGridNo()

        For i As Int16 = 0 To Me.GridNo.Columns.Count - 1
            If Me.GridNo.Columns(i).Name.ToUpper <> "S" Then
                Me.GridNo.Columns(i).ReadOnly = True
            End If
            If Me.GridNo.Columns(i).Name.ToUpper = "S" Or Me.GridNo.Columns(i).Name.ToUpper = "SRV" Then
                Me.GridNo.Columns(i).Width = 25
            ElseIf Me.GridNo.Columns(i).Name.ToUpper = "ORGDOCS" Then
                Me.GridNo.Columns(i).Width = 95
            ElseIf InStr(Me.GridNo.Columns(i).Name.ToUpper, "DATE") > 0 Then
                Me.GridNo.Columns(i).Width = 60
                Me.GridNo.Columns(i).DefaultCellStyle.Format = "dd-MMM-yy"
            ElseIf InStr(Me.GridNo.Columns(i).Name.ToUpper, "CURR") > 0 Then
                Me.GridNo.Columns(i).Width = 45
            ElseIf Me.GridNo.Columns(i).Name.ToUpper = "NOTE" Then
                Me.GridNo.Columns(i).Width = 100
            Else
                Me.GridNo.Columns(i).Width = 85
            End If
        Next
        For i As Int16 = 3 To Me.GridNo.Columns.Count - 1
            If InStr(Me.GridNo.Columns(i).Name.ToUpper, "AMT") > 0 Or _
                InStr(Me.GridNo.Columns(i).Name.ToUpper, "PAID") > 0 Then
                Me.GridNo.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.GridNo.Columns(i).DefaultCellStyle.Format = "#,##0.00"
            ElseIf InStr(Me.GridNo.Columns(i).Name.ToUpper, "ROE") > 0 Then
                Me.GridNo.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.GridNo.Columns(i).DefaultCellStyle.Format = "#,##0.000000"
            Else
                Me.GridNo.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

            If InStr(Me.GridNo.Columns(i).Name.ToUpper, "ORG") > 0 Then
                Me.GridNo.Columns(i).DefaultCellStyle.BackColor = Color.Azure
            End If
        Next
    End Sub

    Private Sub CmbCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCustomer.SelectedIndexChanged
        MyCust.CustID = Me.CmbCustomer.SelectedValue
        If MyCust.CustID = 0 Then Exit Sub
        Me.txtCustFullName.Text = MyCust.FullName
        iCounter = 0
        Me.GridDungTienNao.Visible = True
        Me.GridTraChoKhoanNao.SelectionMode = DataGridViewSelectionMode.CellSelect
        If InStr(varAction.ToUpper, "INVOIC") > 0 Then
            Me.CmdQuickInvUSD.Enabled = True
            Me.CmdQuickInvVND.Enabled = True
            GetPurchases()
        ElseIf InStr(varAction.ToUpper, "PAYMENT") > 0 Then
            LoadNo()
            LoadPayment(Me.CmbPmtType.Text)
        ElseIf InStr(varAction.ToUpper, "FOLLOWUP") > 0 Then
            LoadGridPendingPmt()
        End If
        If InStr("CS_LC", MyCust.CustType) > 0 Then
            LoadCmb_MSC(Me.CmbBooker, "select distinct CA as VAL from rcp where status='OK' and custid=" & MyCust.CustID & " Union Select ''")
        End If
        Me.txtUpto.Value = XacDinhInvUpTo().AddDays(1).AddHours(-2)
        If MyCust.KyBaoCao = "" Then
            Me.txtDueDate.Enabled = True
        Else
            Me.txtDueDate.Enabled = False
        End If
        If MyCust.DayToCfm = 0 Then
            Me.txtDueDate.Value = XacDinhDueDate(Me.txtUpto.Value.Date)
        Else
            Me.txtDueDate.Value = Me.txtUpto.Value.AddDays(MyCust.DayToCfm)
        End If
    End Sub
    Private Sub LoadPayment(ByVal ParPmtType As String)
        Dim strSQL As String
        Me.LckLblWriteOffTra.Visible = False
        Me.txtWriteOffTra.Visible = False
        If Me.CmbCustomer.SelectedValue Is Nothing Then
            Exit Sub
        End If
        strSQL = String.Format("Select RecID, OrgDocs, OrgDate, OrgCurr, OrgAmt, Used, Note, ReceiveBy, FOP, PmtType, ConLai as Balance, linkID" & _
            " from KhachTra where status='OK' and custID={0}", Me.CmbCustomer.SelectedValue)
        If Not Me.BarIncludeFullyUsed.Checked Then
            strSQL = String.Format("{0} and ConLai>0 ", strSQL)
        End If
        If ParPmtType <> "" And Not Me.BarAllPaymentType.Checked Then
            strSQL = String.Format("{0} and pmtType='{1}'", strSQL, Me.CmbPmtType.Text.Trim)
        End If
        Me.GridDungTienNao.DataSource = GetDataTable(strSQL)
        Me.GridDungTienNao.Columns("LinkID").Visible = False
        Me.LbLSplitPmt.Visible = False
        Me.LckLblWriteOffTra.Visible = False
        Me.txtWriteOffTra.Visible = False
        Me.txtAvailable.Text = 0
        Me.BarDeleteThisPayment.Enabled = False
    End Sub
    Private Sub txtCNAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCNAmount.GotFocus

        Me.CmdApply.Visible = False
        If Me.CmbPmtType.Text = "PSP" Then
            For r As Int16 = 0 To Me.GridTraChoKhoanNao.RowCount - 1
                Me.GridTraChoKhoanNao.Item("S", r).Value = False
            Next
        End If
    End Sub
    Private Sub txtCNAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        txtCNAmount.KeyDown, txtAvailable.KeyDown, txtWriteOffTra.KeyDown, TxtWriteOffNo.KeyDown, TxtNegoROE.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub

    Private Sub txtCNAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles _
        txtCNAmount.KeyPress, txtAvailable.KeyPress, txtWriteOffTra.KeyPress, TxtWriteOffNo.KeyPress, TxtNegoROE.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtCNAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCNAmount.LostFocus
        Dim aa As Double, txt As TextBox = CType(sender, TextBox)

        aa = CDbl(txt.Text)
        txt.Text = Format(aa, "#,##0.00")

        If aa <> 0 Then
            Me.CmdAddPayment.Enabled = True
        Else
            Me.CmdAddPayment.Enabled = False
        End If
        If txt.Name = "txtCNAmount" And Me.CmbPmtType.Text = "PSP" Then
            Me.GridDungTienNao.Visible = True
            Me.GridTraChoKhoanNao.Columns("S").Visible = True
            Me.GridTraChoKhoanNao.SelectionMode = DataGridViewSelectionMode.CellSelect
            Me.txtAvailable.Text = Me.txtCNAmount.Text
        End If
    End Sub
    Private Sub LoadNo()
        Dim LoaiCongNo As String = "", strSQL As String
        If MyCust.CustID = 0 Then Exit Sub
        LoaiCongNo = Me.CmbPmtType.Text
        Me.GridTraChoKhoanNao.Columns.Clear()
        strSQL = String.Format("Select RecID, SelectMe as S, invCurr, InvAmt, invDate, DebType, DueDate, Paid, InvAmt-paid as AmtPayThisTime, " & _
            "Note, bsr from GhiNoKhach where custID={0} and debType='PSP' and status='IV' ", MyCust.CustID)
        If Me.BarDueOnly.Checked Then
            strSQL = String.Format("{0} and conno<>0", strSQL)
        Else
            strSQL = String.Format("{0} and conno=0", strSQL)
        End If
        Me.GridTraChoKhoanNao.DataSource = GetDataTable(strSQL)
        Me.GridTraChoKhoanNao.EditMode = DataGridViewEditMode.EditProgrammatically
        resizeGridTraChoKhoanNao()
        resizeGridDungTienNao()
        Me.LckLblWriteOffNo.Visible = False
        Me.TxtWriteOffNo.Visible = False
    End Sub
    Private Sub resizeGridDungTienNao()

        If Me.GridDungTienNao.Columns.Count = 0 Then Exit Sub
        Me.GridDungTienNao.Columns(0).Visible = False
        For c As Int16 = 1 To Me.GridDungTienNao.Columns.Count - 1
            If Me.GridDungTienNao.Columns(c).Name.ToUpper = "S" Then
                Me.GridDungTienNao.Columns(c).Width = 25
            ElseIf InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "AMT") > 0 Or _
                Me.GridDungTienNao.Columns(c).Name.ToUpper = "BALANCE" Or _
                Me.GridDungTienNao.Columns(c).Name.ToUpper = "USED" Then
                Me.GridDungTienNao.Columns(c).Width = 85
            ElseIf InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "DOCS") > 0 Then
                Me.GridDungTienNao.Columns(c).Width = 95
            ElseIf InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "DATE") > 0 Or _
                InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "STATUS") > 0 Or _
                Me.GridDungTienNao.Columns(c).Name.ToUpper = "RECEIVEBY" Or _
                Me.GridDungTienNao.Columns(c).Name.ToUpper = "PMTTYPE" Then
                Me.GridDungTienNao.Columns(c).Width = 60
                Me.GridDungTienNao.Columns(c).DefaultCellStyle.Format = "dd-MMM-yy"
            ElseIf InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "CURR") > 0 Or _
                Me.GridDungTienNao.Columns(c).Name.ToUpper = "FOP" Then
                Me.GridDungTienNao.Columns(c).Width = 45
            ElseIf Me.GridDungTienNao.Columns(c).Name.ToUpper = "NOTE" Then
                Me.GridDungTienNao.Columns(c).Width = 100
            End If
        Next
        For c As Int16 = 2 To Me.GridDungTienNao.Columns.Count - 1
            If InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "AMT") > 0 Or _
                InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "ROE") > 0 Or _
                InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "USED") > 0 Or _
                Me.GridDungTienNao.Columns(c).Name.ToUpper = "BALANCE" Or _
                InStr(Me.GridDungTienNao.Columns(c).Name.ToUpper, "PAID") > 0 Then
                Me.GridDungTienNao.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.GridDungTienNao.Columns(c).DefaultCellStyle.Format = "#,##0.00"
            Else
                Me.GridDungTienNao.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
        Next
    End Sub
    Private Sub resizeGridTraChoKhoanNao()
        Me.GridTraChoKhoanNao.Columns(0).Visible = False
        For c As Int16 = 1 To Me.GridTraChoKhoanNao.Columns.Count - 1
            If Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper = "S" Then
                Me.GridTraChoKhoanNao.Columns(c).Width = 25
            ElseIf InStr(Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper, "AMT") > 0 Then
                Me.GridTraChoKhoanNao.Columns(c).Width = 85
            ElseIf InStr(Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper, "DOCS") > 0 Then
                Me.GridTraChoKhoanNao.Columns(c).Width = 95
            ElseIf InStr(Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper, "DATE") > 0 Or _
            InStr(Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper, "STATUS") > 0 Then
                Me.GridTraChoKhoanNao.Columns(c).Width = 60
                Me.GridTraChoKhoanNao.Columns(c).DefaultCellStyle.Format = "dd-MMM-yy"
            ElseIf InStr(Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper, "CURR") > 0 Or _
                Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper = "DEBTYPE" Then
                Me.GridTraChoKhoanNao.Columns(c).Width = 45
            ElseIf Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper = "NOTE" Then
                Me.GridTraChoKhoanNao.Columns(c).Width = 100
            End If
        Next
        For c As Int16 = 2 To Me.GridTraChoKhoanNao.Columns.Count - 1
            If InStr(Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper, "AMT") > 0 Or _
                InStr(Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper, "PAID") > 0 Then
                Me.GridTraChoKhoanNao.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.GridTraChoKhoanNao.Columns(c).DefaultCellStyle.Format = "#,##0.00"
            ElseIf InStr(Me.GridTraChoKhoanNao.Columns(c).Name.ToUpper, "ROE") > 0 Then
                Me.GridTraChoKhoanNao.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.GridTraChoKhoanNao.Columns(c).DefaultCellStyle.Format = "#,##0.000000000000"
            Else
                Me.GridTraChoKhoanNao.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If
        Next
    End Sub
    Private Sub GridTraChoKhoanNao_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridTraChoKhoanNao.CellContentClick
        Dim Svalue As Boolean, thisAmt As Decimal, BankRate As Decimal, NgayNao As Date
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = 1 Then
            Me.GridTraChoKhoanNao.Item(1, e.RowIndex).Value = Not Me.GridTraChoKhoanNao.Item(1, e.RowIndex).Value
            If Me.BarPaymentDate.Checked Then
                NgayNao = Me.txtPmtDate.Value
            Else
                NgayNao = Me.GridTraChoKhoanNao.Item("InvDate", e.RowIndex).Value
            End If
            thisAmt = Me.GridTraChoKhoanNao.Item("AmtPayThisTime", e.RowIndex).Value
            Me.GridTraChoKhoanNao.CommitEdit(DataGridViewDataErrorContexts.Commit)
            Svalue = Me.GridTraChoKhoanNao.Item(1, e.RowIndex).Value
            If Svalue = False Then
                thisAmt = -thisAmt
                Me.GridTraChoKhoanNao.Item("AmtPayThisTime", e.RowIndex).Value = Me.GridTraChoKhoanNao.Item("InvAmt", e.RowIndex).Value - Me.GridTraChoKhoanNao.Item("Paid", e.RowIndex).Value
            End If
            BankRate = XDTiGia1(TiGiaGi, Me.GridTraChoKhoanNao.Item("InvCurr", e.RowIndex).Value, Me.CmbCNCurr.Text, NgayNao)
            Me.GridTraChoKhoanNao.Item("bsr", e.RowIndex).Value = BankRate
            If CDec(Me.txtAvailable.Text) - thisAmt * BankRate < 0 Then
                Me.GridTraChoKhoanNao.Item("AmtPayThisTime", e.RowIndex).Value = CDec(Me.txtAvailable.Text) / BankRate
                Me.txtAvailable.Text = "0.00"
            Else
                Me.txtAvailable.Text = Format(CDec(Me.txtAvailable.Text) - thisAmt * BankRate, "#,##0.00")
            End If
        ElseIf e.ColumnIndex = 8 Then
            Me.GridTraChoKhoanNao.BeginEdit(True)
        End If
        Me.BarMarkSelectedAsUnpaid.Enabled = True
        Me.LblListDetail.Visible = True
    End Sub
    Private Sub CmdApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdApply.Click
        If Me.GridDungTienNao.CurrentRow.Cells("pmtType").Value <> "PSP" Then Exit Sub
        Dim RecNo As Integer = Me.GridDungTienNao.CurrentRow.Cells("RecID").Value
        Dim CNCurr As String = Me.GridDungTienNao.CurrentRow.Cells("OrgCurr").Value
        Dim CNDocNo As String = Me.GridDungTienNao.CurrentRow.Cells("OrgDocs").Value
        Me.CmdApply.Visible = False
        ApplyPayment(RecNo, CNCurr, CNDocNo)

        CheckOverDue(Me.CmbCustomer.SelectedValue, Conn)

        If Me.CmbPmtType.Text = "PSP" Then
            RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "Apply: " & CNDocNo, Conn, myStaff.SICode, CnStr)
        End If
    End Sub
    Private Sub CalcAmt()
        Dim BankRate As Decimal, NgayNao As Date, HS As Decimal
        For r As Int16 = 0 To Me.GridNo.Rows.Count - 1
            HS = IIf(Me.GridNo.Item("SRV", r).Value = "R", -1, 1)
            If Me.BarInvoiceDate.Checked = True Then
                NgayNao = Me.txtUpto.Value
                BankRate = XDTiGia1(TiGiaGi, Me.GridNo.Item("OrgCurr", r).Value, Me.CmbInvoiceCurr.Text, NgayNao)
            ElseIf Me.BarPurchaseDate.Checked Then
                If Me.GridNo.Item("OrgCurr", r).Value <> "VND" And Me.CmbInvoiceCurr.Text = "VND" Then
                    BankRate = GetROEatPurchase(Me.GridNo.Item("OrgDocs", r).Value, Me.GridNo.Item("Nguon", r).Value)
                Else
                    NgayNao = Me.GridNo.Item("OrgDate", r).Value
                    BankRate = XDTiGia1(TiGiaGi, Me.GridNo.Item("OrgCurr", r).Value, Me.CmbInvoiceCurr.Text, NgayNao)
                End If
            End If
            Me.GridNo.Item("InvCurr", r).Value = Me.CmbInvoiceCurr.Text
            Me.GridNo.Item("InvROE", r).Value = BankRate
            Me.GridNo.Item("InvAmt", r).Value = BankRate * Me.GridNo.Item("OrgTTLAmt", r).Value * HS
            Me.GridNo.Item("InvDate", r).Value = Format(Me.txtUpto.Value, "dd-MMM-yy")
            Me.GridNo.Item("DueDate", r).Value = Format(Me.txtDueDate.Value, "dd-MMM-yy")
        Next
    End Sub
    Private Function GetROEatPurchase(ByVal TRX As String, ByVal pNguon As String) As Decimal
        Dim KQ As Decimal
        If TRX.Substring(0, 2) = "FL" Then
            KQ = ScalarToDec("FLX_FOP", " ROE", " RCPID=" & TRX.Substring(3) & " and status <>'XX' and fop='PSP'")
        Else
            KQ = ScalarToDec("RCP", "ROE", "RCPNO='" & TRX & "' and status <>'XX'")
        End If
        Return KQ
    End Function
    Private Function XDTiGiaNego(ByVal parFrm1 As String, ByVal ParTo1 As String) As Decimal
        Dim KQ As Decimal = 1
        If Me.Cmb1Curr.Text = parFrm1 And Me.CmbEQnCurr.Text = ParTo1 Then
            Return CDec(Me.TxtNegoROE.Text)
        End If
        If Me.Cmb1Curr.Text = ParTo1 And Me.CmbEQnCurr.Text = parFrm1 And _
            CDec(Me.TxtNegoROE.Text) > 1 Then
            Return 1 / CDec(Me.TxtNegoROE.Text)
        End If
    End Function
    Private Function XDTiGia1(ByVal parTiGiaGi As String, ByVal parFrm As String, ByVal parTo As String, ByVal ParNgayNao As Date) As Decimal
        Dim BankRate2 As Decimal, BankRate1 As Decimal, tmpRate As Decimal, KQ As String
        If parFrm = parTo Then
            Return 1
        End If
        parTiGiaGi = parTiGiaGi.ToUpper
        If InStr(parTiGiaGi, "NEGO") > 0 Then
            tmpRate = XDTiGiaNego(parFrm, parTo)
            Return tmpRate
        Else
            If parFrm <> "VND" And parTo <> "VND" Then
                BankRate1 = ForEX_12(ParNgayNao, parFrm, "BSR", "YY").Amount
                BankRate2 = ForEX_12(ParNgayNao, parTo, "BBR", "YY").Amount
                If BankRate1 = 0 Or BankRate2 = 0 Then
                    KQ = 0
                Else
                    KQ = BankRate1 / BankRate2
                End If
            Else
                KQ = ForEX_12(ParNgayNao, parTo, "BBR", "YY").Amount
                If KQ <> 0 And parTo <> "VND" Then
                    KQ = 1 / KQ
                End If
            End If
        End If
        If KQ = 0 Then
            MsgBox("No ROE Available to Convert " & parFrm & " to " & parTo & ". Please Update or Use Nego Rate", MsgBoxStyle.Critical, msgTitle)
        End If
        Return KQ
    End Function

    Private Sub CmbPmtType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbPmtType.SelectedIndexChanged
        If Me.CmbPmtType.Text = "PSP" Then
            Me.GridTraChoKhoanNao.Enabled = True
            Me.CmdApply.Visible = False
            Me.CmdAddPayment.Enabled = False
        Else
            Me.GridTraChoKhoanNao.Enabled = False
            Me.GridTraChoKhoanNao.Columns.Clear()
        End If
        LoadNo()
        LoadPayment(Me.CmbPmtType.Text)
        Me.BarDeleteThisPayment.Enabled = False
    End Sub

    Private Sub GridNo_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridNo.CellContentClick
        If e.ColumnIndex = 0 Then
            Me.GridNo.BeginEdit(True)
        End If
    End Sub
    Private Sub txtCNDocNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCNDocNo.LostFocus
        Me.txtCNAmount.Text = 0
        Me.txtAvailable.Text = 0
    End Sub
    Private Sub BarInvoiceDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        BarInvoiceDate.Click, BarNego.Click, BarPaymentDate.Click, BarPurchaseDate.Click
        Dim MnuBar As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        ResetPadForEx()
        Me.GrpNegoCurr.Visible = True
        MnuBar.Checked = True
        TiGiaGi = MnuBar.Name
        If MnuBar.Name = "BarNego" Then
            Me.TabControl1.SelectTab("TabROE")
        End If
        If MnuBar.Name = "BarNego" Then
            Me.TabControl1.SelectTab("TabROE")
            Me.TabROE.Text = "ROE"
            Me.GrpNegoCurr.Visible = True
        Else
            Me.GrpNegoCurr.Visible = False
            Me.TabROE.Text = "MISC"
        End If
    End Sub
    Private Sub ResetPadForEx()
        Me.BarInvoiceDate.Checked = False
        Me.BarPurchaseDate.Checked = False
        Me.BarPaymentDate.Checked = False
        Me.BarNego.Checked = False
    End Sub
    Private Sub BarUnlock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarUnlock.Click
        Dim strSQL As String = "", tmpRecID As Integer
        Try
            For r As Int16 = 0 To Me.GridTraChoKhoanNao.Rows.Count - 1
                If Me.GridTraChoKhoanNao.Item("S", r).Value = True Then
                    tmpRecID = ScalarToInt("GhiNoKhach", "RecID", " recid=" & Me.GridTraChoKhoanNao.Item("recID", r).Value & " and InvAMt=ConNo")
                    If tmpRecID = 0 Then
                        MsgBox("This Inv Has Been (Partially) Paid. Mark It As Unpaid First, Before Unlock", MsgBoxStyle.Critical, msgTitle)
                        Exit Sub
                    End If
                    strSQL = strSQL & "; " & ChangeStatus_ByID("GhiNoKhach", "XX", Me.GridTraChoKhoanNao.Item("recID", r).Value)
                    strSQL = strSQL & "; update fop set profid=0 where profid=" & Me.GridTraChoKhoanNao.Item("recID", r).Value
                    strSQL = strSQL & "; update FLX_fop set profid=0 where profid=" & Me.GridTraChoKhoanNao.Item("recID", r).Value
                    cmd.CommandText = strSQL.Substring(1)
                    cmd.ExecuteNonQuery()
                    Exit For ' chi chap nhan unlock tung cai 1
                End If
            Next
            MsgBox("Selected Records Have Been UnLocked", MsgBoxStyle.Critical, msgTitle)
        Catch ex As Exception
            MsgBox(Err.Description & "Error Writing To DataBase", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub

    Private Sub BarDueOnly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarDueOnly.Click
        Me.BarDueOnly.Checked = Not Me.BarDueOnly.Checked
        LoadGridNo()
    End Sub

    Private Sub BarIncludeFullyUsed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarIncludeFullyUsed.Click
        Me.BarIncludeFullyUsed.Checked = Not Me.BarIncludeFullyUsed.Checked
        LoadPayment(Me.CmbPmtType.Text)
    End Sub

    Private Sub GridDungTienNao_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles _
        GridDungTienNao.CellContentClick, GridDungTienNao.CellClick
        If e.RowIndex < 0 Then Exit Sub
        LoadNo()
        Me.txtAvailable.Text = Format(Me.GridDungTienNao.Item("Balance", e.RowIndex).Value, "#,##0.00")
        Me.LckLblWriteOffTra.Visible = True
        Me.txtWriteOffTra.Visible = True
        Me.CmbCNCurr.Text = Me.GridDungTienNao.Item("OrgCurr", e.RowIndex).Value
        Me.CmdApply.Visible = True
        Me.LbLSplitPmt.Visible = True
        Me.BarDeleteThisPayment.Enabled = True
    End Sub

    Private Sub GridDungTienNao_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridDungTienNao.CellContentDoubleClick
        If Me.GridDungTienNao.RowCount = 0 Then Exit Sub
        If e.RowIndex < 0 Then Exit Sub
        Me.BarDeleteThisPayment.Enabled = True
        Dim strSQL As String
        strSQL = String.Format("select RecID, DebDocs as OrgDocs, ApplyDate as PmtDate, AmtInDebCurr, Currency as PmtCurr," & _
            "ROE as PmtROE, AmtInPmtCurr, Note from applypayment where KhachTraID={0}", Me.GridDungTienNao.Item("recID", e.RowIndex).Value)
        Me.GridTraChoKhoanNao.Columns.Clear()
        Me.GridTraChoKhoanNao.DataSource = GetDataTable(strSQL)
        resizeGridTraChoKhoanNao()
    End Sub
    Private Sub BarAllPaymentType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarAllPaymentType.Click
        LoadPayment("")
    End Sub

    Private Sub CmdAddPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdAddPayment.Click
        Dim RecNo As Integer, AutoDocNo As String, BSR As Decimal
        Me.LblPrint.Visible = True
        Me.txtAvailable.Text = CDec(Me.txtCNAmount.Text)
        If Me.CmbReceivedBy.Text.Substring(0, 2) = "BO" Then
            AutoDocNo = "BO" & Format(Now, "ddMMyyHHmmss")
        Else
            AutoDocNo = Me.CmbInvoiceAL.Text & Format(Now, "ddMMyyHHmmss")
        End If
        Me.CmdAddPayment.Enabled = False
        If InStr(MyCust.DelayType & "_" & MyCust.AdhType, Me.CmbPmtType.Text) = 0 Then
            MsgBox("Invalid Payment Type. Plz Check", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If Not (Me.txtCNAmount.Text <> 0 And Me.txtCNDescription.Text <> "" And Me.CmbCNCurr.Text <> "" And Me.CmbCNFOP.Text <> "") Then
            MsgBox("Invalid Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        If Me.CmbCNCurr.Text = "VND" Then
            BSR = 1
        Else
            BSR = ForEX_12(Now, Me.CmbCNCurr.Text, "BBR", Me.CmbInvoiceAL.Text).Amount
        End If
        Try
            RecNo = Insert_KhachTra("E", Me.txtPmtDate.Value, Me.CmbCNCurr.Text, Me.CmbCNFOP.Text, AutoDocNo, CDec(Me.txtCNAmount.Text), Me.txtCNDescription.Text, Me.txtPayer.Text, BSR, Me.txtCNDocNo.Text, Me.CmbPmtType.Text, MyCust.CustID)

            If Me.CmbCNCurr.Text <> "VND" Then
                cmd.CommandText = "Update KhachTra set ROE =" & ForEX_12(Me.txtPmtDate.Value, Me.CmbCNCurr.Text, "BBR", Me.CmbInvoiceAL.Text).Amount _
                    & " Where  RecID =" & RecNo
                cmd.ExecuteNonQuery()
            End If
            RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "After AddPmt" & RecNo.ToString, Conn, myStaff.SICode, CnStr)
            LoadPayment(Me.CmbPmtType.Text)
            Me.txtCNDocNo.Text = ""
            Me.txtCNAmount.Text = 0
        Catch ex As Exception
            MsgBox("Error Writing to Dbase", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub
    Private Sub ApplyPayment(ByVal parRecNo As Integer, ByVal ParCurr As String, ByVal parDocNo As String)
        Dim Amt As Decimal, strSQL As String = ""
        For r As Int16 = 0 To Me.GridTraChoKhoanNao.Rows.Count - 1
            If Me.GridTraChoKhoanNao.Item("S", r).Value = True Then
                Amt = Amt + Me.GridTraChoKhoanNao.Item("AmtPayThisTime", r).Value * Me.GridTraChoKhoanNao.Item("bsr", r).Value
                strSQL = strSQL & ";" & Insert_ApplyPayment("S", Me.GridTraChoKhoanNao.Item("RecID", r).Value, parRecNo, _
                    Me.GridTraChoKhoanNao.Item("AmtPayThisTime", r).Value, ParCurr, _
                    Me.GridTraChoKhoanNao.Item("bsr", r).Value, "", "")
                strSQL = strSQL & "; Update GhiNoKhach set Paid=paid+" & Me.GridTraChoKhoanNao.Item("AmtPayThisTime", r).Value & _
                    " where RecID=" & Me.GridTraChoKhoanNao.Item("recID", r).Value
            End If
        Next
        strSQL = strSQL & "; Update KhachTra set Used=Used+" & Amt & " where recid=" & parRecNo
        Try
            cmd.CommandText = strSQL.Substring(1)
            cmd.ExecuteNonQuery()
            MsgBox("Payment Has Been Applied.", MsgBoxStyle.Information, msgTitle)
            LoadNo()
            LoadPayment(Me.CmbPmtType.Text)
        Catch ex As Exception
            MsgBox("Error Writing to Dbase", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub

    Private Sub LblWriteOffTra_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles _
        LckLblWriteOffTra.LinkClicked
        Dim lnk As LinkLabel = CType(sender, LinkLabel)
        Dim RecNo As Int16 = Me.GridDungTienNao.CurrentRow.Cells("RecID").Value
        Dim Curr As String = Me.GridDungTienNao.CurrentRow.Cells("OrgCurr").Value
        Dim tmpMaxWriteOffValue As Decimal
        tmpMaxWriteOffValue = IIf(Curr = "VND", MaxWriteOffValueVND, MaxWriteOffValueUSD)
        If CDec(Me.txtWriteOffTra.Text) > tmpMaxWriteOffValue Then
            If myStaff.SupOf = "" Then
                MsgBox("Not Enough Right To Write Off Such an Amount", MsgBoxStyle.Information, msgTitle)
                Exit Sub
            End If
        End If
        cmd.CommandText = String.Format("Update KhachTra set used=used + {0} where recid={1},", CDec(Me.txtWriteOffTra.Text), RecNo)
        cmd.ExecuteNonQuery()
        cmd.CommandText = Insert_ApplyPayment("S", 0, RecNo, 0, Curr, 1, "", "WRITEOFF")
        cmd.ExecuteNonQuery()
        LoadPayment("PSP")
    End Sub

    Private Sub LblWriteOffNo_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles _
        LckLblWriteOffNo.LinkClicked
        Dim lnk As LinkLabel = CType(sender, LinkLabel)
        Dim RecNo As Integer = Me.GridTraChoKhoanNao.CurrentRow.Cells("RecID").Value
        Dim Curr As String = Me.GridTraChoKhoanNao.CurrentRow.Cells("InvCurr").Value
        Dim tmpMaxWriteOffValue As Decimal
        tmpMaxWriteOffValue = IIf(Curr = "VND", MaxWriteOffValueVND, MaxWriteOffValueUSD)
        If CDec(Me.TxtWriteOffNo.Text) > tmpMaxWriteOffValue Then
            If myStaff.SupOf = "" Then
                MsgBox("Not Enough Right To Write Off Such an Amount", MsgBoxStyle.Information, msgTitle)
                Exit Sub
            End If
        End If

        cmd.CommandText = String.Format("Update GhiNoKhach set Paid=Paid+{0} where recid={1}", CDec(Me.TxtWriteOffNo.Text), RecNo)
        cmd.ExecuteNonQuery()
        cmd.CommandText = Insert_ApplyPayment("S", RecNo, 0, CDec(Me.TxtWriteOffNo.Text), Curr, 1, "", "WRITEOFF")
        cmd.ExecuteNonQuery()
        CheckOverDue(Me.CmbCustomer.SelectedValue, Conn)
        LoadNo()
    End Sub
    Private Function DaDongBaoCaoNgay() As Boolean
        Dim DocNo As String, tmpRCPID As Integer
        For i As Int16 = 0 To Me.GridNo.Rows.Count - 1
            If Me.GridNo.Item("S", i).Value = True Then
                DocNo = Me.GridNo.Item("OrgDocs", i).Value
                If DocNo.Substring(0, 2) <> "FL" Then
                    tmpRCPID = ScalarToInt("RCP", "RecID", " RcpNO='" & DocNo & "' and RPTNO<>''")
                    If tmpRCPID = 0 Then Return False
                End If
            End If
        Next
        Return True
    End Function
    Private Sub CmdQuickInvUSD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdQuickInvUSD.Click, CmdQuickInvVND.Click
        Dim btn As Button = CType(sender, Button)
        Dim MyAns As Int16
        Me.CmdQuickInvUSD.Enabled = False
        Me.CmdQuickInvVND.Enabled = False
        Me.BarPurchaseDate.PerformClick()
        Me.CmbInvoiceCurr.Text = btn.Tag
        If Not DaDongBaoCaoNgay() Then
            MyAns = MsgBox("Not All TRX Has Been Checked And Authorised by Counter. Wanna Quit?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, msgTitle)
            If MyAns = vbYes Then Exit Sub
        End If
        iCounter = 0
        CalcAmt()
        InsertIntoGhiNoKhach("IV")
        LoadGridNo()
    End Sub
    Private Sub txtUpto_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUpto.ValueChanged
        Dim InvUpto As Date = Format(Me.txtUpto.Value, "dd-MMM-yy") & " 23:59"
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        For i As Int16 = 0 To Me.GridNo.RowCount - 1
            If Me.GridNo.Item("OrgDate", i).Value < InvUpto Then
                Me.GridNo.Item("S", i).Value = True
            Else
                Me.GridNo.Item("S", i).Value = False
            End If
        Next
        If Me.CmbCustomer.SelectedValue Is Nothing Then Exit Sub
        If MyCust.DayToCfm = 0 Then
            Me.txtDueDate.Value = XacDinhDueDate(Me.txtUpto.Value.Date)
        Else
            Me.txtDueDate.Value = InvUpto.AddDays(MyCust.DayToCfm)
        End If
        On Error GoTo 0
        Exit Sub
ErrHandler:
        On Error GoTo 0
        Me.txtDueDate.Value = Now.Date
        MsgBox("Invalid [UpTo] value. Due Date Has Been Set = TODAY", MsgBoxStyle.Critical, msgTitle)
    End Sub
    Private Function XacDinhInvUpTo() As Date
        Dim NgayDauThangNay As Date = DateSerial(Now.Year, Now.Month, 1)
        If MyCust.KyBaoCao.Length = 0 Then Return Now.Date
        If MyCust.KyBaoCao.Length = 3 Then
            For i As Int16 = 1 To 7
                If WeekdayName(Weekday(Now.AddDays(-i))).ToString.Substring(0, 3).ToUpper = MyCust.KyBaoCao Then
                    Return Now.Date.AddDays(-i)
                End If
            Next
        ElseIf MyCust.KyBaoCao.Contains("0/0/0") Then
            Return Now.Date
        Else
            If MyCust.KyBaoCao.Split("/").Length = 1 Then
                Return NgayDauThangNay.AddDays(-1)
            ElseIf Now.Day < CInt(MyCust.KyBaoCao.Split("/")(1)) Then
                Return NgayDauThangNay.AddDays(-1)
            Else
                For i As Int16 = 1 To MyCust.KyBaoCao.Split("/").Length - 1
                    If Now.Day >= CInt(MyCust.KyBaoCao.Split("/")(i)) Then
                        Return DateSerial(Now.Year, Now.Month, CInt(MyCust.KyBaoCao.Split("/")(i)) - 1)
                    End If
                Next
            End If
        End If
    End Function

    Private Sub CheckUnCheckALL(ByVal parVal As Boolean)
        For i As Int16 = 0 To Me.GridNo.RowCount - 1
            Me.GridNo.Item("S", i).Value = parVal
        Next
    End Sub

    Private Sub LblUnCheckALL_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUnCheckALL.LinkClicked
        If Me.LblUnCheckALL.Text = "UnCheckALL" Then
            CheckUnCheckALL(False)
            Me.LblUnCheckALL.Text = "CheckALL"
        Else
            Me.LblUnCheckALL.Text = "UnCheckALL"
            CheckUnCheckALL(True)
        End If
    End Sub

    Private Sub BarDeleteThisPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarDeleteThisPayment.Click
        Dim pmtID As Integer = MsgBox("Are You Sure To Delete This Payment? This Will Set All Related Invoices to UnPaid", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
        Dim DaTra As Decimal, RollBackAmt As Decimal, GhiNoID As Int16, StrSQL As String
        Dim dTable As DataTable
        If pmtID = vbNo Then Exit Sub
        pmtID = Me.GridDungTienNao.CurrentRow.Cells("RecID").Value
        StrSQL = ChangeStatus_ByID("KhachTra", "XX", pmtID) & _
            ";" & ChangeStatus_ByDK("ApplyPayment", "XX", "KhachTraID=" & pmtID)
        dTable = GetDataTable("select GhiNoID, AmtInDebCurr from applypayment where KhachTraID=" & pmtID & " and status='OK'")
        For i As Int16 = 0 To dTable.Rows.Count - 1
            GhiNoID = dTable.Rows(i)("GhiNoID")
            RollBackAmt = dTable.Rows(i)("AmtInDebCurr")
            DaTra = ScalarToDec("GhiNoKhach", "Paid", "RecID=" & GhiNoID)
            StrSQL = String.Format("{0} ; Update GhiNoKhacn set Paid={1} where RecID={2}", StrSQL, DaTra - RollBackAmt, GhiNoID)
        Next
        Try
            cmd.CommandText = StrSQL
            cmd.ExecuteNonQuery()
            MsgBox("Payment Deleted.", MsgBoxStyle.Information, msgTitle)
            RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "Del PmtID " & pmtID.ToString, Conn, myStaff.SICode, CnStr)
        Catch ex As Exception
            MsgBox(Err.Description & vbCrLf & "Error Deleting Payment. Action Aborted" & vbCrLf, MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub

    Private Sub BarMarkSelectedAsUnpaid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarMarkSelectedAsUnpaid.Click
        Dim strSQL As String, RollBackAmt As Decimal, AmtUsed As Decimal, KhachTraID As Integer
        Dim InvID As Integer = MsgBox("Are You Sure To Mark This As UnPaid? This Will Revert Amount of Related Payment", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
        Dim dTable As DataTable
        If InvID = vbNo Then Exit Sub
        InvID = Me.GridTraChoKhoanNao.CurrentRow.Cells("RecID").Value

        strSQL = ChangeStatus_ByDK("ApplyPayment", "XX", "GhiNoID=" & InvID) & _
             String.Format("; Update GhiNoKhach set lstUser='{0}', LstUpdate=getdate(), Note=Note +'{0} |RVRTPMT', paid=0 where recid={1}", _
             myStaff.SICode, InvID)
        dTable = GetDataTable("select KhachTraID, AmtInPmtCurr from applypayment where ghiNoID=" & InvID & " and status='OK'")
        For i As Int16 = 0 To dTable.Rows.Count - 1
            KhachTraID = dTable.Rows(i)("KhachTraID")
            RollBackAmt = dTable.Rows(i)("AmtInPmtCurr")
            AmtUsed = ScalarToDec("KhachTra", "Used", "RecID=" & KhachTraID)
            strSQL = String.Format("{0} ; update KhachTra set Used={1} where recid={2}", strSQL, AmtUsed - RollBackAmt, KhachTraID)
        Next
        Try
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()
            MsgBox("Invoice Has Been Marked as UnPaid.", MsgBoxStyle.Information, msgTitle)
        Catch ex As Exception
            MsgBox(Err.Description & vbCrLf & strSQL & vbCrLf & "Error Reverting Payment. Action Aborted", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub

    Private Sub GridTraChoKhoanNao_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridTraChoKhoanNao.CellContentDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        Me.LckLblWriteOffNo.Visible = True
        Me.TxtWriteOffNo.Visible = True
        If e.ColumnIndex = 2 Then
            Exit Sub
            Me.GridDungTienNao.Visible = False
            Me.GridTraChoKhoanNao.Columns("S").Visible = False
            Me.GridTraChoKhoanNao.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            LoadPayment(Me.GridTraChoKhoanNao.Item("DebType", e.RowIndex).Value)
        End If
    End Sub

    Private Sub txtCNDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCNDocNo.TextChanged
        Me.LblPrint.Visible = False
    End Sub

    Private Sub BarCustType_CWT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarCustType_CWT.Click
        Me.BarCustType_CWT.Checked = True
        Me.BarCustType_NonCWTPPD.Checked = False
        Me.BarCustType_NonCWTPSP.Checked = False

        LoadCmb_VAL(Me.CmbCustomer, QryCust & " and custID in " & MyCust.List_CWT)
    End Sub

    Private Sub BarCustType_NonCWTPPD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarCustType_NonCWTPPD.Click
        Me.BarCustType_CWT.Checked = False
        Me.BarCustType_NonCWTPPD.Checked = True
        Me.BarCustType_NonCWTPSP.Checked = False
        LoadCmb_VAL(Me.CmbCustomer, QryCust & " and ppcoef>0 and custID NOT in " & MyCust.List_CWT)
    End Sub

    Private Sub BarCustType_NonCWT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarCustType_NonCWTPSP.Click
        Me.BarCustType_CWT.Checked = False
        Me.BarCustType_NonCWTPPD.Checked = False
        Me.BarCustType_NonCWTPSP.Checked = True

        LoadCmb_VAL(Me.CmbCustomer, QryCust & " and crcoef>0 and custID NOT in " & MyCust.List_CWT)

    End Sub
    Private Sub LblInBaoCaoCongNo_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblInBaoCaoCongNo.LinkClicked
        Dim fname As String = "BaoCaoCongNo" & MyCust.DelayType & ".xlt"
        Dim strPreview As String = IIf(Me.OptRPTPreview.Checked, "V", IIf(Me.OpRPTtPrint.Checked, "O", "F"))
        InHoaDon(Application.StartupPath, fname, strPreview, "YY", Me.txtRPTFrom.Value.Date, Me.TxtRPTThru.Value.Date, Me.CmbCustomer.SelectedValue, "YY", MySession.Domain)
    End Sub

    Private Sub GridCurr_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.PnlMISC.Visible = Not Me.GrpNegoCurr.Visible
    End Sub

    Private Sub LbLSplitPmt_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LbLSplitPmt.LinkClicked
        Dim LinkID As Integer
        LinkID = MsgBox("Are You Sure to Split UnUsed Amount?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, msgTitle)
        If LinkID = vbNo OrElse Me.GridDungTienNao.CurrentRow.Cells("Used").Value = 0 OrElse _
            Me.GridDungTienNao.CurrentRow.Cells("Balance").Value = 0 Then Exit Sub
        LinkID = Me.GridDungTienNao.CurrentRow.Cells("LinkID").Value
        If LinkID = 0 Then LinkID = Me.GridDungTienNao.CurrentRow.Cells("RecID").Value

        cmd.CommandText = String.Format("insert tt_KhachTra (OrgDocs, OrgCurr, OrgAmt, OrgDate, CustID, CustShortName, CustName, FOP, " & _
            "ReceiveBy, FstUser, Description, Note, PmtType, POS, ROE, LstPmt, LinkID) select OrgDocs, OrgCurr, {0}, OrgDate, " & _
            "CustID, CustShortName, CustName, FOP, ReceiveBy, '{1}', Description, Note, PmtType, POS, ROE, LstPmt,{2} from Khachtra " & _
            "where recid={3}; Update KhachTra set OrgAmt=Used, LinkID={2} where recid={3}", Me.GridDungTienNao.CurrentRow.Cells("Balance").Value, _
            myStaff.SICode, LinkID, Me.GridDungTienNao.CurrentRow.Cells("RecID").Value)
        Try
            cmd.ExecuteNonQuery()
            LoadPayment(Me.CmbPmtType.Text)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub txtDueDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDueDate.ValueChanged
        Me.txtDueDate.Enabled = False
    End Sub
    Private Sub LblListDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblListDetail.LinkClicked
        InHoaDon(Application.StartupPath, "R12_CC_TRXDetailByProfINV.xlt", "F", "YY", Now, Now, Me.GridTraChoKhoanNao.CurrentRow.Cells("RecID").Value, "YY", MySession.Domain)
    End Sub

    Private Sub CmbBooker_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbBooker.SelectedIndexChanged
        GetPurchases()
    End Sub
    Private Sub LblTracing_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblTracing.LinkClicked
        Dim KQ As Decimal, strDK As String = " custID=" & Me.CmbCustomer.SelectedValue & " and status<>'XX'"
        cmd.CommandText = "select sum(invamt) from tt_GhiNoKhach where " & strDK
        KQ = cmd.ExecuteScalar
        Me.TxtDaINV.Text = Format(KQ, "#,##0.0")
        cmd.CommandText = "select sum(orgAmt*ROE) from tt_KhachTra where pmtType='PSP' and " & strDK
        KQ = cmd.ExecuteScalar
        Me.TxtDaTra.Text = Format(KQ, "#,##0.0")
        cmd.CommandText = "select isnull(sum(Amount*ROE), 0) from fop where fop='PSP' and status <>'XX' and profID=0 " & _
            " and rcpid in (select recid from rcp where " & strDK & ")"
        KQ = cmd.ExecuteScalar
        Me.TxtChuaINV.Text = Format(KQ, "#,##0.0")
        KQ = CDec(Me.TxtDaINV.Text) - CDec(Me.TxtDaTra.Text) + CDec(Me.TxtChuaINV.Text)
        Me.TxtConNo.Text = Format(KQ, "#,##0.0")
    End Sub
    Private Sub ChkSelectCustOnly_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSelectCustOnly.CheckedChanged
        LoadGridPendingPmt()
    End Sub
    Private Sub LoadGridPendingPmt()
        Me.LblUpdateCFM_VAT.Visible = False
        Try
            Dim StrSQL As String = "select RecID, InvCurr as Curr, InvAmt, ConNo, InvDate, DueDate, CfmDate, VATDate, CustShortName, CustID, FstUpdate " & _
                "from GhiNoKhach where status<>'XX' and ConNo<>0"
            If Me.ChkSelectCustOnly.Checked Then
                'If CmbCustomer.SelectedValue IsNot Nothing Then
                StrSQL = StrSQL & " and CustID=" & Me.CmbCustomer.SelectedValue
                'End If

            Else
                StrSQL = StrSQL & " and custID in " & MyCust.List_CWT
            End If
            Me.GridPendingINV.DataSource = GetDataTable(StrSQL)
        Catch ex As Exception
            Exit Sub
        End Try

        For r As Int16 = 0 To Me.GridPendingINV.RowCount - 1
            If Me.GridPendingINV.Item("DueDate", r).Value < Now.Date Then
                Me.GridPendingINV.Rows(r).DefaultCellStyle.ForeColor = Color.Red
            End If
        Next
        If GridPendingINV.Columns.Count > 0 Then
            Me.GridPendingINV.Columns(0).Visible = False
            Me.GridPendingINV.Columns("FstUpdate").Visible = False
            Me.GridPendingINV.Columns("Curr").Width = 36
            Me.GridPendingINV.Columns("InvDate").Width = 75
            Me.GridPendingINV.Columns("DueDate").Width = 75
            Me.GridPendingINV.Columns("CfmDate").Width = 75
            Me.GridPendingINV.Columns("VATDate").Width = 75
            Me.GridPendingINV.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.GridPendingINV.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.GridPendingINV.Columns(2).DefaultCellStyle.Format = "#,##0.00"
            Me.GridPendingINV.Columns(3).DefaultCellStyle.Format = "#,##0.00"
        End If

    End Sub

    Private Sub GridPendingINV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridPendingINV.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If Me.ChkSelectCustOnly.Checked Then
            If Me.GridPendingINV.CurrentRow.Cells("CFMDate").Value = Me.GridPendingINV.CurrentRow.Cells("VATDate").Value Then
                Me.Label9.Text = "Cfm Date"
                Me.LblUpdateCFM_VAT.Visible = True
            ElseIf Me.GridPendingINV.CurrentRow.Cells("VATDate").Value < Me.GridPendingINV.CurrentRow.Cells("CFMDate").Value Then
                Me.Label9.Text = "VAT Date"
                Me.LblUpdateCFM_VAT.Visible = True
            End If
        End If
    End Sub
    Private Sub ChkMissDL_CheckedChanged(sender As Object, e As EventArgs)
        LoadGridPendingPmt()
    End Sub
    Private Sub LblUpdateCFM_VAT_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdateCFM_VAT.LinkClicked
        Dim LeadTime As Int16
        If Me.Label9.Text.Substring(0, 3) = "Cfm" Then
            cmd.CommandText = "update GhiNoKhach set CfmDate=@Date where RecID=@RecID"
        Else
            LeadTime = ScalarToInt("KyBaoCao", "daysToVAT", "CustID=" & Me.GridPendingINV.CurrentRow.Cells(0).Value & " and Status='OK'")
            cmd.CommandText = "update GhiNoKhach set VATDate=@Date, DueDate=@DueDate where RecID=@RecID"
        End If
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@Date", SqlDbType.DateTime).Value = Me.TxtCFM_VAT.Value
        cmd.Parameters.Add("@recID", SqlDbType.Int).Value = Me.GridPendingINV.CurrentRow.Cells(0).Value
        If Me.Label9.Text.Substring(0, 3) = "VAT" Then
            cmd.Parameters.Add("@DueDate", SqlDbType.DateTime).Value = Me.TxtCFM_VAT.Value.AddDays(LeadTime)
        End If
        cmd.ExecuteNonQuery()
        LoadGridPendingPmt()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        iCounter = iCounter + 1
        If iCounter = 8 Then Me.Close()
    End Sub
    Private Function XacDinhDueDate(pUpTo As Date) As Date
        If MyCust.KyBaoCao.Length = 1 Then ' 1 thang 1 ky, doi ngay 15
            Return DateSerial(Now.Year, Now.Month, 15).AddHours(23)
        ElseIf MyCust.KyBaoCao.Length = 3 Then ' theo DOW
            Return pUpTo.AddDays(MyCust.DueIn).AddHours(23)
        ElseIf Month(pUpTo) <> Month(Now) Then ' ky cuoi thang truoc, due ngay truoc element 1
            Return DateSerial(Now.Year, Now.Month, CInt(MyCust.KyBaoCao.Split("/")(1)) - 1).AddHours(23)
        Else
            For i As Int16 = 1 To MyCust.KyBaoCao.Split("/").Length - 1
                If pUpTo.Day + 1 = CInt(MyCust.KyBaoCao.Split("/")(i)) Then
                    If i = MyCust.KyBaoCao.Split("/").Length - 1 Then ' ky up chot thang nay, due cuoi thang nay
                        Return DateSerial(Now.Year, Now.Month, 1).AddMonths(1).AddDays(-1).AddHours(23)
                    Else
                        Return DateSerial(Now.Year, Now.Month, CInt(MyCust.KyBaoCao.Split("/")(i + 1)) - 1).AddHours(23)
                    End If
                End If
            Next
        End If
        Return pUpTo.AddDays(MyCust.DueIn).AddHours(23)
    End Function
End Class