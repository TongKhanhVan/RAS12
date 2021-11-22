Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.MySharedFunctionsWzConn
Public Class InvoicePrinting
    Private MyCust As New objCustomer
    Private WhoIs As String = ""
    Private Const CashLimitPerInv As Decimal = 20000 * 1000
    Private SoHoaDonCu As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private TRXAmt_VND As Decimal, AlreadyINV_VND As Decimal
    Public Sub New()
        InitializeComponent()
    End Sub
    Public Sub New(ByVal parWhois As String, parRCPNO As String)
        InitializeComponent()
        WhoIs = parWhois
        Me.TxtDocNo.Text = parRCPNO
    End Sub
    Private Sub LoadTKT(ByVal ParDKRCP As String)
        Dim strSQL As String
        Me.txtTTLofSelectedTKT.Text = 0
        Me.GridTKT2BInv.Rows.Clear()
        ' cac ve ban theo PTA thi ko xuat hoa don vi minh ko Ban ma chi xuat ho ve thoi
        ' Cac MCO ban ra thi ko xuat hoa don vi coi MCO la tien, chi "Doir tien" thoi. 
        ' KHi nao xuat ve dung MCO nay thi xuat HD tren ca gia tri ve thoi, tuc la report dthu 1 lan dung nhu ban chat cua no
        strSQL = "select TKNO, FTKT, t.RCPNo, t.Currency as Curr, Fare, Tax, t.Charge, ChargeTV, "
        strSQL = strSQL & "AgtDisctVAL as Discount, r.ROE, itinerary, RCPID, t.RecID,t.srv,qty, NetToAL "
        strSQL = strSQL & " from TKT T inner join RCP r on t.RCPID=R.RecID"
        strSQL = strSQL & " AND r.custType in " & myStaff.CAccess
        strSQL = strSQL & " where t.status<>'XX' "
        strSQL = strSQL & " and tkno not in (select TKNO from TKTNO_INVNO where status='OK' and rcpid=r.recid)"
        strSQL = strSQL & " and t.RCPNO not in (select RCPNO from FOP where FOP='PTA') and "
        strSQL = strSQL & " t.AL in (select AL from airline where VAT like '%" & MySession.Domain.Substring(0, 1) & "Y%') and "
        strSQL = strSQL & ParDKRCP

        Me.GridTKT.DataSource = GetDataTable(strSQL)
        Me.GridTKT.Columns("TKNO").Width = 95
        Me.GridTKT.Columns("FTKT").Width = 50
        Me.GridTKT.Columns("RCPNO").Width = 85
        Me.GridTKT.Columns("Curr").Width = 35
        Me.GridTKT.Columns("Fare").Width = 85
        Me.GridTKT.Columns("Tax").Width = 75
        Me.GridTKT.Columns("Charge").Width = 75
        Me.GridTKT.Columns("ChargeTV").Width = 75
        Me.GridTKT.Columns("Discount").Width = 75
        Me.GridTKT.Columns("ROE").Visible = False
        Me.GridTKT.Columns("RecID").Visible = False
        Me.GridTKT.Columns("Itinerary").Visible = False
        Me.GridTKT.Columns("SRV").Visible = False
        Me.GridTKT.Columns("Qty").Visible = False
        Me.GridTKT.Columns("RCPID").Visible = False
        Me.GridTKT.Columns("Fare").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("Tax").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("Charge").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("Discount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.GridTKT.Columns("Fare").DefaultCellStyle.Format = "#,##0.00"
        Me.GridTKT.Columns("Tax").DefaultCellStyle.Format = "#,##0.00"
        Me.GridTKT.Columns("Charge").DefaultCellStyle.Format = "#,##0.00"
        Me.GridTKT.Columns("ChargeTV").DefaultCellStyle.Format = "#,##0.00"
        Me.GridTKT.Columns("Discount").DefaultCellStyle.Format = "#,##0.00"

        SoHoaDonCu = ScalarToString("RCP", "AL", " rcpno='" & Me.TxtDocNo.Text & "'")

    End Sub

    Private Sub InvoicePrinting_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MyAL.Domain = MySession.Domain
        MyAL.ALCode = ""
        MyCust.CustID = 0
    End Sub
    Private Sub InvoicePrinting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If HasNewerVersion_R12(Application.ProductVersion) Or SysDateIsWrong(Conn) Then
            Me.Close()
            Me.Dispose()
            End
        End If

        GenCmbValue()
        If WhoIs = "FO" Then
            Me.txtIssDate.Enabled = False
        End If
    End Sub
    Private Sub GenCmbValue()
        LoadCmb_VAL(Me.CmbCust, "select RecID as VAL, CustshortName as DIS from CustomerList")
    End Sub

    Private Sub LblAddToInvoice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAddToInvoice.LinkClicked
        Me.LblPrintInv.Text = "Save and Print"
        AddTKT2Inv(True, 1)
    End Sub
    Private Sub AddTKT2Inv(ByVal parStatus As Boolean, ByVal parNbrOfInv As Int16)
        Dim strTKT As String, strFTKT As String, strCurr As String = "", decFare As Decimal
        Dim decTax As Decimal, decCharge As Decimal, decDiscount As Decimal, decROE As Decimal, strRTG As String
        Dim StrRCPNO As String = "", IntRCPID As Integer, IntRecID As Integer, strTKTj As String
        Dim strSRV As String, intQty As Int16, decNetToAL As Decimal, decChargeTV As Decimal
        Me.GridTKT2BInv.Rows.Clear()
        For i As Int16 = 0 To Me.GridTKT.RowCount - 1
            If Me.GridTKT.Item("S", i).Value = parStatus Then
                strTKT = Me.GridTKT.Item("TKNO", i).Value
                strFTKT = Me.GridTKT.Item("FTKT", i).Value
                strCurr = Me.GridTKT.Item("Curr", i).Value
                decFare = Me.GridTKT.Item("Fare", i).Value / parNbrOfInv
                decTax = Me.GridTKT.Item("Tax", i).Value / parNbrOfInv
                decCharge = Me.GridTKT.Item("Charge", i).Value / parNbrOfInv
                decChargeTV = Me.GridTKT.Item("ChargeTV", i).Value / parNbrOfInv
                decDiscount = Me.GridTKT.Item("Discount", i).Value / parNbrOfInv
                decROE = Me.GridTKT.Item("ROE", i).Value
                strRTG = Me.GridTKT.Item("Itinerary", i).Value
                StrRCPNO = Me.GridTKT.Item("RCPNO", i).Value
                IntRCPID = Me.GridTKT.Item("RCPID", i).Value
                IntRecID = Me.GridTKT.Item("RecID", i).Value
                strSRV = Me.GridTKT.Item("SRV", i).Value
                intQty = Me.GridTKT.Item("Qty", i).Value
                decNetToAL = Me.GridTKT.Item("NetToAL", i).Value
                strTKTj = strTKT
                For j As Int16 = 1 To parNbrOfInv
                    If parNbrOfInv > 1 Then strTKTj = strTKT & "(" & Trim(Str(j)) & ")"
                    Me.GridTKT2BInv.Rows.Add("P", strTKTj, strFTKT, strCurr, decFare, decTax, decCharge, decChargeTV, decDiscount, j, decROE, strRTG, StrRCPNO, IntRCPID, IntRecID, strSRV, intQty, decNetToAL)
                Next
            End If
        Next
    End Sub
    Private Sub GridTKT_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridTKT.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = 0 And _
            (Me.GridTKT.Item(2, e.RowIndex).Value.ToString.Trim = "" OrElse _
            Me.GridTKT.Item(2, e.RowIndex).Value.ToString.Substring(0, 3) = "___") Then
            Me.GridTKT.Item(0, e.RowIndex).Value = Not Me.GridTKT.Item(0, e.RowIndex).Value
            If Me.GridTKT.Item(2, e.RowIndex).Value.ToString.Trim <> "" Then
                CheckAllConjTKT(Me.GridTKT.Item(1, e.RowIndex).Value, Me.GridTKT.Item(0, e.RowIndex).Value)
            End If
            ReCalctxtTTLofSelectedTKT()
        End If
    End Sub
    Private Sub CheckAllConjTKT(ByVal ParTKNO As String, ByVal parStatus As Boolean)
        Dim PhanDau As String, PhanSau As Decimal, VeTiep As String, ConjTiep As String
        Dim tmpPhanSau As String
        VeTiep = ParTKNO
        Do
            PhanDau = VeTiep.Substring(0, 9)
            PhanSau = CDec(VeTiep.Substring(9, 6)) + 1
            tmpPhanSau = "00000" & Trim(Str(PhanSau))
            VeTiep = PhanDau & tmpPhanSau.Substring(tmpPhanSau.Length - 6, 6)
            For r As Int16 = 0 To Me.GridTKT.RowCount - 1
                If Me.GridTKT.Item("TKNO", r).Value = VeTiep Then
                    Me.GridTKT.Item(0, r).Value = parStatus
                    ConjTiep = Me.GridTKT.Item("FTKT", r).Value
                    If ConjTiep.Trim.Length = 3 Then Exit Sub
                End If
            Next
        Loop
    End Sub
    Private Sub LblSplitInvvoice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSplitInvoice.LinkClicked
        Dim j As Int16 = 0, TRX_ChargeDisc As Decimal
        Dim NbrOfInv As Int16 = Int(CDec(Me.txtTTLofSelectedTKT.Text) / CashLimitPerInv) + 1
        TRX_ChargeDisc = ScalarToDec(" RCP ", " sum(charge+discount) ", " RCPNO='" & Me.TxtDocNo.Text & "'")
        If TRX_ChargeDisc <> 0 Then
            MsgBox("This TRX Has Charge and/or Discount and Thus Cant be Splited", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        Me.LblPrintInv.Text = "Save"
        For i As Int16 = 0 To Me.GridTKT.RowCount - 1
            If Me.GridTKT.Item("S", i).Value = True And Me.GridTKT.Item("Fare", i).Value > 0 Then j = j + 1
        Next
        If j > 1 Or CDec(Me.txtTTLofSelectedTKT.Text) < CashLimitPerInv Then
            If j > 1 Then MsgBox("Please Select Only ONE Ticket", MsgBoxStyle.Information, msgTitle)
            Exit Sub
        End If
        AddTKT2Inv(True, NbrOfInv)
    End Sub

    Private Sub LblDocSearch_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDocSearch.LinkClicked
        Dim currMonth As String = Format(Me.txtIssDate.Value, "MMyy")
        If Me.TxtDocNo.Text.Substring(0, 2) = "XX" Then
            MyAL.Domain = "EDU"
        Else
            MyAL.Domain = IIf(Me.TxtDocNo.Text.Substring(0, 2) = "TS", "TVS", "GSA")
        End If
        MyAL.ALCode = Me.TxtDocNo.Text.Substring(0, 2)
        If InStr(myStaff.AAccess, "YY") + InStr(myStaff.AAccess, Me.TxtDocNo.Text.Substring(0, 2)) = 0 Then
            MsgBox("Invalid Input. You May not Have Right to Issue INV for This TRX", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If Me.TxtDocNo.Text.Substring(2, 4) <> currMonth And WhoIs = "FO" Then
            MsgBox("Invalid Input. You Cannot Issue Invoice For Transaction From Other Month", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        DocSearch()
    End Sub
    Private Sub DocSearch()
        Dim strDKTKT As String = "TKT where status<>'XX' and RCPno='" & Me.TxtDocNo.Text & "'"
        Dim strDKRCP As String, tmpINVNo As String, tmpINVID As Integer
        Dim CSH_PmtVND As Decimal, EXC_MCO As Decimal, tmpRCPID As Integer
        tmpRCPID = ScalarToInt("RCP", "RecID", " RCPNO='" & Me.TxtDocNo.Text & "' and CustType in " & myStaff.CAccess & " and city='" & MySession.City & "'")
        If tmpRCPID = 0 Then Exit Sub



        '' DOAN DUOI DAY BO SUNG 26MAY do sau khi trien khai ko in, nhieu TRX ko giu HD thanh cong
        ' check xem da giu so HD cho RCPID nay chua
        tmpINVID = ScalarToInt("INV", "top 1 RecID", "status ='OK' and RCPID=" & tmpRCPID)
        tmpINVNo = ScalarToString("INV", "InvNo ", "RecID=" & tmpINVID)
        If tmpINVID > 0 Then
            'check xem da tao ban ghi TKTNO_INV chua
            Dim CheckBanGhiTKTNO_Exist As Integer = ScalarToInt("TKTNO_INVNO", "top 1 INVID", "status ='OK' and RCPID=" & tmpRCPID)
            Dim tblRCP As DataTable = GetDataTable("select * from rcp where recid=" & tmpRCPID)
            Dim tmpFOP As String = ConnvertFOPtoString(tmpRCPID)
            Dim tmpTTL As Decimal
            If CheckBanGhiTKTNO_Exist = 0 Then
                Dim tmpCounter As String = ScalarToString("RCP", "Counter", "RecID=" & tmpRCPID)
                If tmpCounter = "CWT" Then
                    TaoBanGhiTKTNO_INVNO_Standard(tmpRCPID, tmpINVNo, tmpINVID, "WO")
                    tmpTTL = ScalarToDec("TKTNO_INVNO", "sum(F_VND+T_VND+C_VND)", "INVID=" & tmpINVID & " and status='OK'")
                    Dim dtblTVTR As DataTable = GetDataTable("select VAL, VAL1, VAL2 from MISC where cat='TVCompany' and description='TVT'")
                    cmd.CommandText = Update_INV(pubVarSRV, dtblTVTR.Rows(0)("VAL"), dtblTVTR.Rows(0)("VAL1"), dtblTVTR.Rows(0)("VAL2"), _
                        tmpTTL, tmpFOP, 0, tmpINVID, tblRCP.Rows(0)("CustID"))
                Else
                    TaoBanGhiTKTNO_INVNO_Standard(tmpRCPID, tmpINVNo, tmpINVID, "WZ")
                    tmpTTL = ScalarToDec("TKTNO_INVNO", "sum(F_VND+T_VND+C_VND+CTV_VND)", "INVID=" & tmpINVID & " and status='OK'")
                    cmd.CommandText = Update_INV(pubVarSRV, tblRCP.Rows(0)("CustFullName"), tblRCP.Rows(0)("CustAddress"), tblRCP.Rows(0)("CustTaxCode"), _
                        tmpTTL, tmpFOP, 0, tmpINVID, tblRCP.Rows(0)("CustID"))
                End If
                cmd.ExecuteNonQuery()
            End If
        End If
        '' END BO SUNG 26MAY

        Me.TxtSRV.Text = ScalarToString("RCP", "SRV", "RecID=" & tmpRCPID)
        Me.CmbCust.SelectedValue = ScalarToInt("RCP", "CustID", "RecID=" & tmpRCPID)
        If Me.TxtSRV.Text <> "V" Then
            If tmpINVID > 0 Then
                SoHoaDonCu = tmpINVNo
                MsgBox("Invoice " & tmpINVNo & " Issued. Click Preview for Details", MsgBoxStyle.Information, msgTitle)
                Me.LblPreview.Visible = True
            Else
                MyCust.CustID = Me.CmbCust.SelectedValue
                Me.txtFullName.Text = MyCust.FullName
                Me.txtAddress.Text = MyCust.Addr
                Me.txtTaxCode.Text = MyCust.taxCode
                Me.LblPreview.Visible = False
                strDKRCP = " r.RecID=" & tmpRCPID
                LoadTKT(strDKRCP)
            End If
            TRXAmt_VND = ScalarToDec("RCP", "ttldue * roe", "recid=" & tmpRCPID)
            CSH_PmtVND = ScalarToDec("FOP", "isnull(sum(Amount * roe),0)", " rcpid=" & tmpRCPID & " and status+FOP='OKCSH'")
            EXC_MCO = ScalarToDec("FOP", "isnull(sum(Amount * roe),0)", " rcpid=" & tmpRCPID & " and status='OK' and fop in ('EXC','MCO')")
            If CSH_PmtVND < CashLimitPerInv Then
                CheckAllTKT()
                Me.GridTKT.Columns(0).ReadOnly = True
            Else
                Me.GridTKT.Columns(0).ReadOnly = False
            End If
            Me.LblSplitInvoice.Enabled = Not Me.GridTKT.Columns(0).ReadOnly
            Me.LblSelectAll.Enabled = Not Me.GridTKT.Columns(0).ReadOnly
            AlreadyINV_VND = ScalarToDec("INV", "isnull(sum(Amount ),0)", " rcpid=" & tmpRCPID & " and status='OK'")
            If AlreadyINV_VND >= TRXAmt_VND Then
                Me.GroupBox4.Enabled = False
            Else
                Me.GroupBox4.Enabled = True
            End If
            Me.GroupBox5.Enabled = Me.GroupBox4.Enabled
        End If
    End Sub
    Private Function ConnvertFOPtoString(pRCPID As Integer) As String
        Dim tbl As DataTable = GetDataTable("select * from fop where status <>'XX' and FOP <>'RND' and RCPID=" & pRCPID)
        Dim KQ As String = ""
        For i As Int16 = 0 To tbl.Rows.Count - 1
            If tbl.Rows(i)("Amount") <> 0 Or tbl.Rows(i)("Document") <> "" Then
                KQ = KQ & "|" & tbl.Rows(i)("FOP")
                KQ = KQ & "_" & tbl.Rows(i)("Currency")
                If tbl.Rows(i)("Amount") Is Nothing Then
                    KQ = KQ & "_" & "0"
                Else
                    KQ = KQ & "_" & tbl.Rows(i)("Amount").ToString
                End If
                KQ = KQ & "_" & tbl.Rows(i)("Document")
            End If

        Next
        Return KQ.Substring(1)
    End Function
    Private Sub LblCalc_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCalc.LinkClicked
        If Me.GridTKT2BInv.RowCount = 0 Then Exit Sub
        CalcInvAmt()
    End Sub
    Private Sub CalcInvAmt()
        Dim TTLAmt As Decimal = 0, MyAns As Int16
        For r As Int16 = 0 To Me.GridTKT2BInv.RowCount - 1
            TTLAmt = TTLAmt + Me.GridTKT2BInv.Item("Fare", r).Value
            TTLAmt = TTLAmt + Me.GridTKT2BInv.Item("Tax", r).Value
            TTLAmt = TTLAmt + Me.GridTKT2BInv.Item("Charge", r).Value * Me.GridTKT2BInv.Item("Qty", r).Value
            TTLAmt = TTLAmt + Me.GridTKT2BInv.Item("ChargeTV", r).Value * Me.GridTKT2BInv.Item("Qty", r).Value
            TTLAmt = TTLAmt - Me.GridTKT2BInv.Item("Discount", r).Value
        Next
        TTLAmt = TTLAmt * Me.GridTKT2BInv.Item("ROE", 0).Value
        Me.TxtTTLofInv.Text = Format(TTLAmt, "#,##0.00")
        If TTLAmt - CDec(Me.txtTTLofSelectedTKT.Text) > 0.01 Then
            If WhoIs = "FO" Then
                MsgBox("Invalid Input. Invoice Amout Is Greater Than Amount Of Selected Ticket", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            Else
                MyAns = MsgBox("Invoice Amount Is Greater Than Amount Of Selected Ticket? Wanna Recheck Your Input?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
                If MyAns = vbYes Then Exit Sub
            End If
        ElseIf CDec(Me.txtTTLofSelectedTKT.Text) - TTLAmt > 0.01 Then
            MyAns = MsgBox("Invoice Amount Is Less Than Amount Of Selected Ticket? Wanna Recheck Your Input?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
            If MyAns = vbYes Then Exit Sub
        End If
        Me.LblPrintInv.Enabled = True
    End Sub
    Private Sub LblSelectAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSelectAll.LinkClicked
        CheckAllTKT()
    End Sub
    Private Sub ReCalctxtTTLofSelectedTKT()
        Dim ThisAmt As Decimal
        Me.txtTTLofSelectedTKT.Text = 0
        For i As Int16 = 0 To Me.GridTKT.RowCount - 1
            If Me.GridTKT.Item(0, i).Value = True Then
                ThisAmt = CDec(Me.GridTKT.Item("Fare", i).Value) + CDec(Me.GridTKT.Item("Tax", i).Value)
                ThisAmt = ThisAmt + CDec(Me.GridTKT.Item("Charge", i).Value) * CDec(Me.GridTKT.Item("Qty", i).Value)
                ThisAmt = ThisAmt + CDec(Me.GridTKT.Item("ChargeTV", i).Value) * CDec(Me.GridTKT.Item("Qty", i).Value)
                ThisAmt = ThisAmt - CDec(Me.GridTKT.Item("Discount", i).Value)
                ThisAmt = ThisAmt * CDec(Me.GridTKT.Item("ROE", i).Value)
                txtTTLofSelectedTKT.Text = Format(CDec(txtTTLofSelectedTKT.Text) + ThisAmt, "#,##0.00")
            End If
        Next

        Me.GridTKT2BInv.Rows.Clear()
        Me.TxtTTLofInv.Text = 0
        Me.LblPrintInv.Enabled = False
    End Sub
    Private Sub CheckAllTKT()
        Dim varStatus As Boolean
        If Me.LblSelectAll.Text = "Select All" Then
            Me.LblSelectAll.Text = "DeSelect All"
            varStatus = True
        Else
            Me.LblSelectAll.Text = "Select All"
            varStatus = False
        End If
        For i As Int16 = 0 To Me.GridTKT.RowCount - 1
            Me.GridTKT.Item(0, i).Value = varStatus
        Next
        ReCalctxtTTLofSelectedTKT()

    End Sub

    Private Sub GridTKT2BInv_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridTKT2BInv.CellLeave
        Me.LblPrintInv.Enabled = False
    End Sub

    Private Sub LblPrintInv_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblPrintInv.LinkClicked
        Dim SoHDDaBiHuy As String, LstInv As String = "", InvID As Integer, InvNo As String = ""
        Dim SoHDDaBiDungBoi As Integer, HS As Int16, NgayTaoHDDauTien As Date
        Dim CurrMonth As String = Format(Now, "MMyy"), MyAns As Integer, strSQL As String = ""
        Dim MsgStr As String = "", BaoNhieuHD As Int16, InvDau As String = ""
        MyAL.ALCode = Me.TxtDocNo.Text.Substring(0, 2)
        MyAns = MsgBox("Are You Sure that All Invoices Are Viewed and Correct?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
        If MyAns = vbNo Then Exit Sub
        Me.LblPrintInv.Enabled = False
        For i As Int16 = 0 To Me.GridTKT2BInv.RowCount - 1
            BaoNhieuHD = IIf(Me.GridTKT2BInv.Item("InvNo", i).Value > BaoNhieuHD, Me.GridTKT2BInv.Item("InvNo", i).Value, BaoNhieuHD)
        Next
        HS = IIf(Me.TxtSRV.Text = "R", -1, 1)
        If BaoNhieuHD > 1 Then
            MsgBox("Pls Contact VTH for Help", MsgBoxStyle.Critical, msgTitle)
            ' neu sua cho phan tach hD, xem them vong For j =0
            Exit Sub
        End If
        Try
            Dim ROE As Decimal = Me.GridTKT.Item("ROE", 0).Value

            For i As Int16 = 1 To BaoNhieuHD
                SoHDDaBiDungBoi = -1
                InvNo = ""

                SoHDDaBiHuy = ScalarToString("TKTNO_INVNO", "top 1 INVNO", "RCPID=" & Me.GridTKT2BInv.Item("RCPID", 0).Value & _
                    " and status='XX'" & _
                    " and TKNO='" & IIf(BaoNhieuHD = 1, Me.GridTKT2BInv.Item("TKNO", 0).Value, Me.GridTKT2BInv.Item("TKNO", i - 1).Value) & "'")

                If Not String.IsNullOrEmpty(SoHDDaBiHuy) Then
                    'SoHDDaBiDungBoi = ScalarToInt("INV", " RCPID", " status='OK' and invNo='" & SoHDDaBiHuy & "' ")
                    SoHDDaBiDungBoi = ScalarToInt("INV", "RCPID", " status='OK' and invNo='" & SoHDDaBiHuy & _
                                                  "'  and FstUpdate>=(Select FstUpdate from RCP where RcpId=" _
                                                    & Me.GridTKT.Item("RCPID", 0).Value & ")" & _
                                                  " order by fstupdate desc")

                End If
                If SoHDDaBiDungBoi = 0 Then
                    InvNo = SoHDDaBiHuy

                    'Tu nam 2017 se doi thanh gia tri cu
                    'NgayTaoHDDauTien = ScalarToDate("INV", "Top 1 Fstupdate", "INVNO='" & InvNo & "' order by Fstupdate")
                    NgayTaoHDDauTien = ScalarToDate("INV", "Top 1 Fstupdate", "INVNO='" & InvNo _
                                                    & "' and FstUpdate>=(Select top 1 FstUpdate from RCP where RcpId=" _
                                                    & Me.GridTKT.Item("RCPID", 0).Value & ")" _
                                                    & " order by Fstupdate")

                Else
                    InvNo = GenInvNo_QD153(Me.TxtDocNo.Text, MyAL.VAT_KyHieu)
                    NgayTaoHDDauTien = Now
                End If
                InvID = Insert_INV("E", InvNo, InvNo.Substring(0, 2), Me.GridTKT.Item("RCPID", 0).Value, NgayTaoHDDauTien)
                If i = 1 Then InvDau = InvNo

                For j As Int16 = 0 To Me.GridTKT2BInv.RowCount - 1
                    'dong code sau chi dung voi truong hop ko tach HD
                    cmd.CommandText = " insert TKTNO_INVNO (TKNO, INVNO, INVID, RCPID, FstUser, F_VND, T_VND, C_VND, CTV_VND)" & _
                        " values (@TKNO, @INVNO, @INVID, @RCPID, @FstUser, @F_VND, @T_VND, @C_VND, @CTV_VND)"
                    cmd.Parameters.Clear()
                    cmd.Parameters.Add("@TKNO", SqlDbType.Char).Value = Me.GridTKT2BInv.Item("TKNO", j).Value
                    cmd.Parameters.Add("@INVNO", SqlDbType.Char).Value = InvNo
                    cmd.Parameters.Add("@FstUser", SqlDbType.Char).Value = myStaff.SICode
                    cmd.Parameters.Add("@INVID", SqlDbType.Int).Value = InvID
                    cmd.Parameters.Add("@RCPID", SqlDbType.Int).Value = Me.GridTKT2BInv.Item("RCPID", j).Value
                    cmd.Parameters.Add("@F_VND", SqlDbType.Decimal).Value = (Me.GridTKT2BInv.Item("Fare", j).Value - Me.GridTKT2BInv.Item("Discount", j).Value) * HS * ROE
                    cmd.Parameters.Add("@T_VND", SqlDbType.Decimal).Value = Me.GridTKT2BInv.Item("Tax", j).Value * HS * ROE
                    cmd.Parameters.Add("@C_VND", SqlDbType.Decimal).Value = Me.GridTKT2BInv.Item("Charge", j).Value * ROE
                    cmd.Parameters.Add("@CTV_VND", SqlDbType.Decimal).Value = Me.GridTKT2BInv.Item("ChargeTV", j).Value * ROE
                    cmd.ExecuteNonQuery()
                Next
                cmd.CommandText = UpdateTblInv(InvID, i) & ";" & UpdateTblINVHistory(InvID, InvNo, 0)
                cmd.ExecuteNonQuery()
            Next
            Me.GridTKT2BInv.Rows.Clear()
            Me.GridTKT.DataSource = Nothing
            Me.GridTKT.Rows.Clear()
            InHoaDon(Application.StartupPath, "R12_VATInvoice.xlt", "O", InvDau, Now.Date, Now.Date, 0, Me.TxtDocNo.Text.Substring(0, 2), MySession.Domain)
        Catch ex As Exception
            FileOpen(1, "d:\ErrCreatingINV.txt", OpenMode.Append)
            Print(1, cmd.CommandText)
            FileClose(1)
            MsgBox("Error Creating Invoice. " & vbCrLf & ex.Message & vbCrLf & cmd.CommandText, MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub
    Private Function UpdateTblInv(ByVal ParInvID As Integer, ByVal parHDi As Int16) As String
        Dim tmpFOP As String, TTLAmt As Decimal, TTLFare As Decimal, TTLTax As Decimal
        Dim TTLCharge As Decimal, TTLChargeTV As Decimal
        Dim LclSRV As String = Me.TxtSRV.Text
        Dim tblTKNO_INVNO As DataTable = GetDataTable("Select * from TKTNO_INVNO where status='OK' and INVID=" & ParInvID)
        For i As Int16 = 0 To tblTKNO_INVNO.Rows.Count - 1
            TTLFare = TTLFare + tblTKNO_INVNO.Rows(i)("F_VND")
            TTLTax = TTLTax + tblTKNO_INVNO.Rows(i)("T_VND")
            TTLCharge = TTLCharge + tblTKNO_INVNO.Rows(i)("C_VND")
            TTLChargeTV = TTLChargeTV + tblTKNO_INVNO.Rows(i)("CTV_VND")
        Next
        TTLAmt = TTLFare + TTLTax
        TTLAmt = Math.Abs(TTLAmt)
        If LclSRV = "R" Then
            TTLAmt = TTLAmt - TTLCharge
        Else
            TTLAmt = TTLAmt + TTLCharge
        End If
        If InStr("CS_LC", MyCust.CustType) = 0 Then ' voi khach CWT thi ko in SF
            If Me.GridTKT2BInv.Item("SRV", 0).Value = "R" Then
                TTLAmt = TTLAmt - TTLChargeTV
            Else
                TTLAmt = TTLAmt + TTLChargeTV
            End If
        End If
        tmpFOP = GetFOPdetail()
        If InStr("CS_LC", MyCust.CustType) > 0 Then ' voi khach CWT thi in HD cho TVTR
            Dim dtblTVTR As DataTable = GetDataTable("select VAL, VAL1, VAL2 from MISC where cat='TVCompany' and description='TVT'")
            Return Update_INV(LclSRV, dtblTVTR.Rows(0)("VAL"), dtblTVTR.Rows(0)("VAL1"), _
                dtblTVTR.Rows(0)("VAL2"), TTLAmt, tmpFOP, 1, ParInvID, MyCust.CustID)
        Else
            Return Update_INV(LclSRV, Me.txtFullName.Text.Replace("--", ""), Me.txtAddress.Text.Replace("--", ""), _
                Me.txtTaxCode.Text.Replace("--", ""), TTLAmt, tmpFOP, 1, ParInvID, MyCust.CustID)
        End If
    End Function
    Private Function GetFOPdetail() As String
        Dim KQ As String = ""
        Dim dtbl As DataTable = GetDataTable("select FOP, Currency, Amount, Document, ROE, Status from FOP where status in ('OK','QQ') " & _
                " and fop <>'RND' and rcpid=" & Me.GridTKT.Item("RCPID", 0).Value)
        If dtbl.Rows.Count = 0 Then Return KQ
        For i As Int16 = 0 To dtbl.Rows.Count - 1
            KQ = KQ & "|" & dtbl.Rows(i)("FOP") & "_" & dtbl.Rows(i)("Currency") & "_" & dtbl.Rows(i)("Amount").ToString
            KQ = KQ & "_" & dtbl.Rows(i)("Document")
            If dtbl.Rows(i)("Status") = "QQ" Then
                MsgBox("Credit Card Payment Not Yet Finished. Action Continues but Plz Inform Internal Acct", MsgBoxStyle.Information, msgTitle)
            End If
        Next
        Return KQ.Substring(1)
    End Function
    Private Sub LblGo2InvHandler_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblGo2InvHandler.LinkClicked
        Dim f As New InvHandler("FO")
        f.Show()
    End Sub
    Private Sub LblPreview_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblPreview.LinkClicked
        Dim tmpAL As String
        If SoHoaDonCu.Substring(0, 2) = "TS" Then
            tmpAL = SoHoaDonCu.Substring(7, 2)
        Else
            tmpAL = SoHoaDonCu.Substring(0, 2)
        End If
        InHoaDon(Application.StartupPath, "R12_VATInvoice.xlt", "V", SoHoaDonCu, Now.Date, Now.Date, 0, tmpAL, MySession.Domain)
        Me.LblPreview.Visible = False
    End Sub

    Private Sub LblPreview_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblPreview.VisibleChanged
        Me.txtFullName.Enabled = Not Me.LblPreview.Visible
        Me.txtAddress.Enabled = Me.txtFullName.Enabled
        Me.txtTaxCode.Enabled = Me.txtFullName.Enabled
    End Sub

    Private Sub LblClose_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblClose.LinkClicked
        Me.Close()
        Me.Dispose()
    End Sub
    Private Sub LblAutoFillTVTR_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAutoFillTVTR.LinkClicked
        Dim dTbl As DataTable = GetDataTable("select VAL, VAL1,Val2 from MISC where cat='TVCompany' and description='TVT'")
        Me.txtFullName.Text = dTbl.Rows(0)("VAL")
        Me.txtAddress.Text = dTbl.Rows(0)("VAL1")
        Me.txtTaxCode.Text = dTbl.Rows(0)("VAL2")
    End Sub
End Class