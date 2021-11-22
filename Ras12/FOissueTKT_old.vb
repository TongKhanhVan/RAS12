Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.MySharedFunctionsWzConn
Imports SharedFunctions.crd_ctrl
Public Class FOissueTKT_OLD
    Private MyCust As New objCustomer
    Private CharEntered As Boolean = False
    Private DaTra As Decimal = 0
    Private ESCPressed As Boolean = False
    Private RCPNOtoPrint As String
    Private mRCP As String = "", WhatAction As String = ""
    Private DepositRow As Int16, OldValue As Decimal
    Private DKstatusTKT As String, DKstatusFOP As String
    Private ExcDocDetail As String = ""
    Private ExcDocList As String = ""
    Private QuitNow As Boolean = False
    Private amtCRDPaid As Decimal
    Private DefauSF As Decimal
    Private DK_Status_Validity As String = " and status='OK' and ('" & Now.Date & "' between validFrom and validThru) "
    Private NextTKNO As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private CanEditROE As Boolean = False
    Private NonAirList As String = "('HTL','CAR','INS','VSA','AHC','EVT','TVS')"
    Private SalesInforOfThisTKT As New SalesInfor
    Private isIATARate As Boolean = False, isFromTKTList As Boolean = False
    Private Structure SalesInfor
        Public Comm As Decimal
        Public Rtg As String
        Public Tax As Decimal
        Public Discount As Decimal
        Public Fare As Decimal
        Public NetToAL As Decimal
        Public isEXC As Boolean
    End Structure
    Private Sub FeedBack(ByVal pMsg As String)
        If pMsg.Length = 0 Then
            Me.StatusFeedBack.ForeColor = Color.Blue
        ElseIf pMsg.Substring(0, 3).ToUpper = "ERR" Then
            Me.StatusFeedBack.ForeColor = Color.Red
        Else
            Me.StatusFeedBack.ForeColor = Color.Blue
        End If
        Me.StatusFeedBack.Text = pMsg
        Me.Timer1.Enabled = True
    End Sub
    Public Sub New()
        InitializeComponent()
    End Sub
    Public Sub New(ByVal parAction As String)
        InitializeComponent()
        Dim POS As Int16
        POS = InStr(parAction, "_")
        mRCP = parAction.Substring(POS)
        If POS = 4 Then WhatAction = parAction.Substring(0, 3)
    End Sub

    Private Sub txtRTG_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRTG.Enter
        MyCust.GetSIR(Me.CmbDocType.Text)
        If MyCust.HasSIR Then InitGrpSIR()
    End Sub
    Private Sub txtRTG_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        txtRTG.KeyDown, txtFB.KeyDown
        Dim hotStr As String = "", Fi As String = ""
        Dim txt As TextBox = CType(sender, TextBox)
        If e.KeyValue = 27 Then ESCPressed = Not ESCPressed
        If e.KeyValue > 111 And e.KeyValue < 124 Then
            Fi = "F" & (e.KeyValue - 111).ToString.Trim
            hotStr = getHotStr(Me.CmbAL.Text, Fi, txt.Name)
            txt.Text = hotStr
        End If
    End Sub

    Private Sub txtRTG_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRTG.LostFocus
        Me.CmdAddTKT.Enabled = True
        If InStr(NonAirList, Me.CmbDocType.Text) > 0 Then Exit Sub
        Dim RTG As String, tmpFB As String = "", j As Int16
        If Me.txtRTG.Text.Trim = "" Then Exit Sub
        RTG = Me.txtRTG.Text.Replace(" ", "").Trim
        If RTG.Length > 83 Then
            FeedBack("Error. Invalid Itinerary Length!")
            Me.txtRTG.Focus()
        End If
        If RTG.Length > 23 And _
            (InStr("SO", pubVarSRV) = 0 Or MyAL.isTKTLess) And Me.CmbAL.Text <> "XX" Then
            FeedBack("Error. Invalid Itinerary Length!")
            Me.txtRTG.Focus()
        End If
        If Not CheckRTG(RTG.Trim) Then
            FeedBack("Error. Invalid Itinerary Format!")
            Me.txtRTG.Focus()
        End If
        Me.txtRTG.Text = AddSpace2Rtg(RTG)
        If pubVarSRV = "A" Then Exit Sub
        If pubVarSRV = "R" Then
            Me.txtFB.Enabled = False
            Me.TxtBkgClass.Enabled = False
            Me.CmbPaxType.Enabled = False
            Me.txtDOI.Enabled = False
            Me.txtFltDate.Enabled = False
            Me.TxtPaxName.Enabled = False
            Me.CmbPromoCode.Enabled = False
            Me.TxtTourCode.Enabled = False
            Me.GrpFare.Enabled = True
            Me.txtFare.Enabled = True
            If Me.txtRTG.Text <> SalesInforOfThisTKT.Rtg Then
                Me.TxtBkgClass.Text = Strings.Right(Me.TxtBkgClass.Text, (Me.txtRTG.Text.Trim.Length - 3) / 7)
                j = UBound(Me.txtFB.Text.Split("+"))
                For i As Int16 = Me.TxtBkgClass.Text.Length - 1 To 0 Step -1
                    tmpFB = tmpFB & "+" & Me.txtFB.Text.Split("+")(j - i)
                Next
                Me.txtFB.Text = tmpFB.Substring(1)
            End If
        ElseIf pubVarSRV = "A" Then
            Me.GrpTax.Enabled = False
            Me.GrpComm.Enabled = False
            Me.GrpFare.Enabled = False
        End If
    End Sub
    Private Sub txtALCommVAL_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtALCommVAL.Enter
        Me.txtALCommVAL.Text = Format(CDec(Me.txtFare.Text) - CDec(Me.txtNetToAL.Text), "#,##0.00")
    End Sub
    Private Sub txtFare_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        txtFare.GotFocus, txtNetToAL.GotFocus, TxtAgtDiscPCT.GotFocus, TxtAgtDiscVal.GotFocus, _
        txtALCommVAL.GotFocus, TxtTax.GotFocus, txtRTG.GotFocus
        Me.CmdAddTKT.Enabled = False
        Me.CmdSaveChange.Enabled = False
    End Sub
    Private Sub txtFare_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        txtFare.KeyDown, txtTax1.KeyDown, TxtTax2.KeyDown, TxtTax3.KeyDown, TxtTax4.KeyDown, TxtTax5.KeyDown, _
        TxtTax.KeyDown, TxtCharge2.KeyDown, TxtCharge3.KeyDown, TxtCharge1.KeyDown, _
        txtALCommPCT.KeyDown, txtNetToAL.KeyDown, TxtAgtDiscPCT.KeyDown, TxtAgtDiscVal.KeyDown, TxtCharge4.KeyDown, _
        txtTVCharge1.KeyDown, txtTVCharge2.KeyDown, txtTVDiscount1.KeyDown, txtTVDiscount2.KeyDown, _
        txtShownFare.KeyDown, txtALCommVAL.KeyDown, txtLLF.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub
    Private Sub txtFare_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles _
        txtFare.KeyPress, txtShownFare.KeyPress, txtTax1.KeyPress, TxtTax2.KeyPress, TxtTax3.KeyPress, _
        TxtTax5.KeyPress, TxtTax4.KeyPress, TxtCharge4.KeyPress, txtALCommVAL.KeyPress, _
        TxtTax.KeyPress, TxtCharge2.KeyPress, TxtCharge3.KeyPress, txtTVDiscount2.KeyPress, _
        TxtCharge1.KeyPress, txtALCommPCT.KeyPress, txtNetToAL.KeyPress, TxtAgtDiscPCT.KeyPress, _
        TxtAgtDiscVal.KeyPress, txtTVCharge1.KeyPress, txtTVCharge2.KeyPress, txtTVDiscount1.KeyPress, txtLLF.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub
    Private Sub LoadCmbFareSource_FareType(ByVal pDK_Status_Validity As String, ByVal pFare As Decimal, ByVal SR As String)
        Dim pDK_AL_SBU As String = " and AL+SBU='" & MyAL.ALCode & MySession.Domain & "'"
        Dim strSQL As String = " as VAL from serviceFee where TRXType='" & pubVarSRV & "' " & _
            pDK_Status_Validity & " and " & pFare & " between FareFrom and FareThrough " & _
            " and channel='" & Me.CmbCustType.Text & "' "
        If Me.CmbCustType.Text = "WK" Then
            strSQL = strSQL & " and SBU+AL in (Select SBU+AL "
        Else
            strSQL = strSQL & " and SBU+AL+Channel+CustLevel in (Select SBU+AL+Channel+CustLevel "
        End If
        strSQL = strSQL & " from cust_Channel_level where " & pDK_AL_SBU.Substring(4) & pDK_Status_Validity
        If Me.CmbCustType.Text <> "WK" Then
            strSQL = strSQL & " and Custid =" & Me.CmbChannel.SelectedValue
        End If
        strSQL = strSQL & ")"
        LoadCmb_MSC(Me.CmbGRPFIT, "select distinct G_FIT " & strSQL)
        LoadCmb_MSC(Me.CmbFSource, "select distinct FID_fier " & strSQL)
    End Sub
    Private Sub txtFare_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFare.LostFocus
        If Me.txtFare.Text = "" Then Me.txtFare.Text = "0"
        Dim aa As Double = CDbl(Me.txtFare.Text)
        Me.txtFare.Text = Format(aa, "#,##0.00")
        If aa < 0 Then Me.txtFare.Focus()
        If InStr("MCO_GRP", Me.CmbDocType.Text) > 0 Then
            Me.CmdAddTKT.Enabled = True
            Me.CmdAddTKT.Focus()
            Exit Sub
        End If
        ReCalcALComm()
        If InStr("TA_TO_WK", Me.CmbCustType.Text) > 0 Then
            Me.CmbGRPFIT.Visible = True
            LoadCmbFareSource_FareType(DK_Status_Validity, CDec(Me.txtFare.Text), pubVarSRV)
            CalServiceFee(DK_Status_Validity)
        ElseIf InStr("CS_LC", Me.CmbCustType.Text) > 0 Then
            Me.CmbGRPFIT.Visible = False
            'CalServiceFee_CWT(DK_Status_Validity)
        End If
    End Sub
    Private Sub ToggleMenuFormIssueTKT()
        Dim mnu As Object, MnuI As ToolStripItem
        myStaff.CurrObj = Me.Name
        For Each mnu In Me.MenuStrip2.Items
            If InStr(myStaff.URights, mnu.name.ToString.ToUpper) > 0 Then
                mnu.enabled = False
            Else
                For Each MnuI In mnu.DropDownItems
                    If InStr(myStaff.URights, MnuI.Name.ToString.ToUpper) > 0 Then
                        MnuI.Enabled = False
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub FOissueTKT_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.OptRefund.Enabled = False
        Me.OptSell.Enabled = False
        Me.OptVoid.Enabled = False
        Me.OptRefundOLD.Enabled = False
        Me.OptServicing.Enabled = False
        Me.OptReIss.Enabled = False
        Me.OptPseudoVoid.Enabled = False
        isFromTKTList = False
    End Sub
    Private Sub FOissueTKT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If HasNewerVersion_R12(Application.ProductVersion) Or SysDateIsWrong(Conn) Then
            Me.Close()
            Me.Dispose()
            End
        End If
        If MySession.Domain = "" Or MySession.City = "" Or MySession.Location = "" Then
            MsgBox("You Have to Specify Counter for this Computer. Contact Supervisor for Help", MsgBoxStyle.Critical, msgTitle)
            Me.Close()
            Exit Sub
        End If
        MyCust.GenCustList()

        Me.txtFB.Width = 226
        Me.FOP_FOP.Items.Add("CSH")
        Me.RCPCurrency.Items.Add("VND")
        Me.GrpReasonCode.Top = Me.GrpComm.Top
        Me.GrpReasonCode.Left = Me.GrpComm.Left

        LoadCmbAL(Me.CmbAL)
        LoadCmb_MSC(Me.CmbTVCharge1, "TVCharge")
        LoadCmb_MSC(Me.CmbTVCharge2, "TVCharge")
        Me.StatusCounter.Text = "Domain: " & MySession.Domain & ". Counter: " & MySession.Counter

        LoadCmb_MSC(Me.CmbCustType, "select VAL from MISc where cat ='Channel' and VAL in " & myStaff.CAccess)
        If InStr(myStaff.CAccess, "CS") + InStr(myStaff.CAccess, "LC") > 0 Then
            LoadCmb_MSC(Me.CmbTC, "select SICode as VAL from TblUser where sicode not like '%**' and recid in " & _
                " (select UserID from UserAccess where status='OK' and cat='ChannelAccess' and val like '%CS')")
            Me.CmbTC.Items.Add("")
            Me.CmbTC.Text = ""
        End If
        SetBackGroudColor(Me)
        GenFOP_CURR()
        If WhatAction = "DEL" Or WhatAction = "PRN" Then
            Me.OptPrintInv.Enabled = False
            Me.OptPrintTRX.Enabled = False
            Me.OptPrintTRX.Checked = True
            Me.GrpCharge.Enabled = False
            Me.GrpFare.Enabled = False
            Me.GrpTax.Enabled = False
            Me.GrpGenInfor.Enabled = False
            Me.GrpPaymentDetails.Enabled = False
            DKstatusTKT = " Status <> 'XX'"
            DKstatusFOP = " Status='OK'"
            If WhatAction = "DEL" Then Me.CmdDelete.Visible = True
        ElseIf WhatAction = "FOP" Then
            Me.GrpGenInfor.Enabled = False
            Me.GrpCharge.Enabled = False
            Me.GrpFare.Enabled = False
            Me.GrpTax.Enabled = False
            DKstatusTKT = " Status <>'XX'"
            DKstatusFOP = " Status='OK'"
            Me.CmdChangeFOP.Visible = True
            Me.GridFOP.Enabled = True
        ElseIf WhatAction = "MON" Then
            LoadCmb_MSC(Me.CmbCurr, " select distinct Currency as VAL from forEx where Status='OK' ")
            Me.GrpPaymentDetails.Enabled = False
            Me.GrpGenInfor.Enabled = True
            DKstatusTKT = " Status<>'XX'"
            DKstatusFOP = " Status='OK'"
            Me.CmdSave.Visible = False
            Me.CmdAdjust.Visible = True
        ElseIf WhatAction = "NON" Then
            Me.CmbPromoCode.Enabled = False
            Me.GrpCharge.Enabled = False

            Me.GrpFare.Enabled = False
            Me.GrpTax.Enabled = False
            DKstatusTKT = " StatusAL='OK'"
            DKstatusFOP = " Status='OK'"
            Me.CmdSaveNON.Visible = True
        Else
            DKstatusTKT = " Status='OK'"
            DKstatusFOP = " Status='OK'"
        End If

        If mRCP = "" Then
            pubVarRCPID_BeingEdited = 0
            Me.CmdDelete.Enabled = False
        Else
            pubVarRCPID_BeingCreated = 0
            pubVarRCPID_BeingEdited = ScalarToInt("RCP", "RecID ", " status<>'XX' and RCPNO+SBU='" & mRCP & MySession.Domain & "'")
            Me.TxtTRXNO.Text = mRCP
            RCPNOtoPrint = mRCP
            GenerateComboValue(Me)
            LoadData(pubVarRCPID_BeingEdited)
            Me.GrpSRV.Visible = False
            Me.GrpNewRCP.Visible = True
            Me.CmdSave.Visible = False
            Me.GrpNewRCP.Enabled = False

            Me.CmbCharge1.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCharge2.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCharge3.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCharge4.DropDownStyle = ComboBoxStyle.DropDown
            genDocType()
        End If

        If InStr("EDU_TVS", MySession.Domain) > 0 Then
            Me.txtNetToAL.ReadOnly = False
            Me.txtNetToAL.TabStop = True
            Me.BarALLPendingRQ.Enabled = True
            Me.BarPendingRQ.Enabled = True
            ToggleMenuFormIssueTKT()
        End If

        Me.GrpPaymentDetails.Left = 6
        Me.GrpPaymentDetails.Width = 706
        Me.CmbProduct.Text = "NET"
        If myStaff.SupOf <> "" Then Me.LcktxtROE.ReadOnly = False
    End Sub
    Private Sub CmdAddTKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdAddTKT.Click
        Me.CmbCurr.Enabled = False
        Me.CmbChannel.Enabled = False ' 2 lenh nay de dam bao KO doi vi da co ruot/ve 
        ESCPressed = False
        ExcDocList = ExcDocList & "_" & Me.txtExcDoc.Text
        If Me.txtTKNO.Text.Trim.Length < 13 Or _
            (Me.txtExcDoc.Text.Trim.Length > 0 And Me.txtExcDoc.Text.Trim.Replace(" ", "").Length <> 13) Then
            FeedBack("Error. Invalid Ticket Number!")
            Me.txtTKNO.Focus()
            Exit Sub
        End If
        If Not String.IsNullOrEmpty(Me.txtExcDoc.Text) Then Me.txtExcDoc.Text = AddSpace2TKNO(Me.txtExcDoc.Text)
        If Not LogicCheckOK() Then Exit Sub

        Me.CmdAddTKT.Enabled = False
        addTKTtoGridTKT("ADD")

        ClearInput2(Me)
        Me.CmdSaveChange.Enabled = False
        FeedBack("Updated!")
        Me.CmbChannel.Enabled = False
        Me.CmbCustType.Enabled = False
        Me.CmbCurr.Enabled = False
        Me.txtTKNO.Focus()

    End Sub

    Private Sub addTKTtoGridTKT(ByVal varAction As String)
        Dim TTL As Decimal = 0, NetToAL As Decimal = 0, RowNo As Integer
        Dim TC As String, TKNO As String, Amt2AL As Decimal, TKTRMK As String = ""
        Dim SoVeNoi As Int16, SoChang As Int16, k As Int16, CommOffer As String
        TC = Me.TxtTourCode.Text
        If TC = "TOUR CODE" Then TC = ""
        TKNO = Me.txtTKNO.Text.Trim.Replace(" ", "")
        TKNO = AddSpace2TKNO(TKNO)
        If varAction = "CHANGE" Then
            RowNo = Me.GridTKT.CurrentCell.RowIndex
            If Me.GridTKT.Item(0, RowNo).Style.Alignment = DataGridViewContentAlignment.MiddleRight Then
                ' ko edit ve noi
                Exit Sub
            End If
            If WhatAction = "NON" AndAlso myStaff.SupOf <> "" AndAlso _
                Me.GridTKT.CurrentRow.Cells("FTKT").Value = "" Then
                Me.GridTKT.Item("TKNO", RowNo).Value = Me.txtTKNO.Text
            End If
            Me.GridTKT.Item("IsEdit", RowNo).Value = "Y"
        Else
            Me.GridTKT.Rows.Add()
            RowNo = Me.GridTKT.Rows.Count - 1
            Me.GridTKT.Item("RowID", RowNo).Value = RowNo
            Me.GridTKT.Item("TKNO", RowNo).Value = TKNO
            Me.GridTKT.Item("IsEdit", RowNo).Value = ""
        End If
        If InStr("VSA_CAR", Me.CmbDocType.Text) > 0 Then
            CommOffer = Me.TxtCountry.Text
        Else
            CommOffer = IIf(Me.GrpComm.Visible, Me.CmbCommOffer.Text, "")
        End If
        NetToAL = CDec(Me.txtNetToAL.Text)
        Amt2AL = NetToAL + CDec(Me.TxtTax.Text)
        Amt2AL = Amt2AL + DefineAmtCharge2AL()

        Me.GridTKT.Item("AmountToAL", RowNo).Value = Amt2AL

        Me.GridTKT.Item("SRV", RowNo).Value = pubVarSRV
        Me.GridTKT.Item("Qty", RowNo).Value = 1
        Me.GridTKT.Item("CommOffer", RowNo).Value = CommOffer
        If Me.CmbDocType.Text = "AHC" Then
            Me.GridTKT.Item("DOI", RowNo).Value = Me.txtDOI.Value
        Else
            Me.GridTKT.Item("DOI", RowNo).Value = Me.txtDOI.Value.Date
        End If
        Me.GridTKT.Item("CustType", RowNo).Value = Me.CmbCustType.Text
        Me.GridTKT.Item("DOF", RowNo).Value = Me.txtFltDate.Value.Date
        Me.GridTKT.Item("Routing", RowNo).Value = Me.txtRTG.Text
        Me.GridTKT.Item("BkgClass", RowNo).Value = Me.TxtBkgClass.Text
        Me.GridTKT.Item("FareBasis", RowNo).Value = Me.txtFB.Text
        Me.GridTKT.Item("PaxName", RowNo).Value = Me.TxtPaxName.Text
        Me.GridTKT.Item("DocType", RowNo).Value = Me.CmbDocType.Text
        Me.GridTKT.Item("PaxType", RowNo).Value = Me.CmbPaxType.Text
        Me.GridTKT.Item("StockCtrl", RowNo).Value = AddSpace2TKNO(Me.txtExcDoc.Text)
        Me.GridTKT.Item("TourCode", RowNo).Value = TC
        Me.GridTKT.Item("PromoCode", RowNo).Value = Me.CmbPromoCode.Text
        Me.GridTKT.Item("FareType", RowNo).Value = Me.CmbProduct.Text.Trim
        Me.GridTKT.Item("Fare", RowNo).Value = CDec(Me.txtFare.Text)
        Me.GridTKT.Item("ShownFare", RowNo).Value = CDec(Me.txtShownFare.Text)
        Me.GridTKT.Item("NetToAL", RowNo).Value = NetToAL
        Me.GridTKT.Item("AgtDisctPCT", RowNo).Value = CDec(Me.TxtAgtDiscPCT.Text)
        Me.GridTKT.Item("AgtDisctVAL", RowNo).Value = CDec(Me.TxtAgtDiscVal.Text)
        Me.GridTKT.Item("CommPCT", RowNo).Value = CDec(Me.txtALCommPCT.Text)
        Me.GridTKT.Item("LLF", RowNo).Value = CDec(Me.txtLLF.Text)
        Me.GridTKT.Item("CommVAL", RowNo).Value = CDec(Me.txtALCommVAL.Text)
        Me.GridTKT.Item("Tax", RowNo).Value = CDec(Me.TxtTax.Text)
        Me.GridTKT.Item("Charge_D", RowNo).Value = GomTax_Charge("C")
        Me.GridTKT.Item("Charge", RowNo).Value = CalcCharge(Me.GridTKT.Item("Charge_D", RowNo).Value, "SP", Me.CmbCurr.Text, CDec(Me.LcktxtROE.Text), Me.TxtDOS.Value.Date, MySession.TRXCode)
        Me.GridTKT.Item("ChargeTV", RowNo).Value = CalcCharge(Me.GridTKT.Item("Charge_D", RowNo).Value, "TV", Me.CmbCurr.Text, CDec(Me.LcktxtROE.Text), Me.TxtDOS.Value, MySession.TRXCode)
        Me.GridTKT.Item("Tax_D", RowNo).Value = GomTax_Charge("T")
        If varAction = "ADD" Then Me.txtTKNO.Text = DefineNextTKNO(Me.txtTKNO.Text, MyAL.isTKTLess)
        If InStr("CS_LC", Me.CmbCustType.Text) > 0 Then
            TKTRMK = "BIZ" & Me.ChkBizTrip.Checked.ToString.Substring(0, 1)
            TKTRMK = TKTRMK & "|BKR" & Me.CmbBooker.Text
            If Me.CmbDocType.Text = "HTL" Then
                TKTRMK = TKTRMK & "|MSC" & Me.CmbMSaving_HTL.Text.Split("-")(0)
                TKTRMK = TKTRMK & "|RSC" & Me.cmbRSaving_HTL.Text.Split("-")(0)
            End If
        End If
        Me.GridTKT.Item("TKT_RMK", RowNo).Value = TKTRMK
        If Me.txtRTG.Text.Trim.Length > 32 And varAction = "ADD" And InStr(TKNO, "GRP") + InStr(TKNO, "TV") = 0 Then
            Me.GridTKT.Item("FTKT", RowNo).Value = "___/" & Strings.Right(Me.txtTKNO.Text, 3)
            SoChang = (Me.txtRTG.Text.Trim.Length - 3) / 7
            If SoChang > 12 Then
                SoVeNoi = 3
            ElseIf SoChang > 8 Then
                SoVeNoi = 2
            ElseIf SoChang > 4 Then
                SoVeNoi = 1
            End If
            For i As Int16 = 1 To SoVeNoi
                Me.GridTKT.Rows.Add()
                RowNo = Me.GridTKT.Rows.Count - 1
                Me.GridTKT.Item("RowID", RowNo).Value = RowNo
                Me.GridTKT.Item("TKNO", RowNo).Value = Me.txtTKNO.Text
                TC = Strings.Right(Me.txtTKNO.Text, 3)
                Me.GridTKT.Item("IsEdit", RowNo).Value = ""
                Me.GridTKT.Item("SRV", RowNo).Value = pubVarSRV
                Me.GridTKT.Item("Qty", RowNo).Value = 0
                Me.GridTKT.Item("CommOffer", RowNo).Value = IIf(Me.GrpComm.Visible, Me.CmbCommOffer.Text, "")
                Me.GridTKT.Item("DOI", RowNo).Value = Me.txtDOI.Value.Date
                Me.GridTKT.Item("CustType", RowNo).Value = Me.CmbCustType.Text
                Me.GridTKT.Item("DOF", RowNo).Value = Me.txtFltDate.Text
                Me.GridTKT.Item("PaxName", RowNo).Value = Me.TxtPaxName.Text
                Me.GridTKT.Item("DocType", RowNo).Value = Me.CmbDocType.Text
                Me.GridTKT.Item("PaxType", RowNo).Value = Me.CmbPaxType.Text
                Me.GridTKT.Item("StockCtrl", RowNo).Value = ""
                Me.GridTKT.Item("TourCode", RowNo).Value = ""
                Me.GridTKT.Item("PromoCode", RowNo).Value = ""
                Me.GridTKT.Item("FareType", RowNo).Value = ""
                Me.GridTKT.Item("Fare", RowNo).Value = 0
                Me.GridTKT.Item("ShownFare", RowNo).Value = 0
                Me.GridTKT.Item("LLF", RowNo).Value = 0
                Me.GridTKT.Item("NetToAL", RowNo).Value = 0
                Me.GridTKT.Item("AgtDisctPCT", RowNo).Value = 0
                Me.GridTKT.Item("AgtDisctVAL", RowNo).Value = 0
                Me.GridTKT.Item("CommPCT", RowNo).Value = 0
                Me.GridTKT.Item("CommVAL", RowNo).Value = 0
                Me.GridTKT.Item("Tax", RowNo).Value = 0
                Me.GridTKT.Item("Charge", RowNo).Value = 0
                Me.GridTKT.Item("ChargeTV", RowNo).Value = 0
                Me.GridTKT.Item("Tax_D", RowNo).Value = ""
                Me.GridTKT.Item("Charge_D", RowNo).Value = ""
                If i = SoVeNoi Then
                    Me.GridTKT.Item("FTKT", RowNo).Value = Format(CInt(TC) - 1, "000")
                    Me.GridTKT.Item("Routing", RowNo).Value = Me.txtRTG.Text.Substring(28 * i)
                    Me.GridTKT.Item("BkgClass", RowNo).Value = Me.TxtBkgClass.Text.Substring(4 * i)
                    k = SoChang - 1
                Else
                    Me.GridTKT.Item("FTKT", RowNo).Value = Format(CInt(TC) - 1, "000") & "/" & Format(CInt(TC) + 1, "000")
                    Me.GridTKT.Item("Routing", RowNo).Value = Me.txtRTG.Text.Substring(28 * i, 31)
                    Me.GridTKT.Item("BkgClass", RowNo).Value = Me.TxtBkgClass.Text.Substring(4 * i, 4)
                    k = 4 * i + 3
                End If
                TC = ""
                For j As Int16 = 4 * i To k
                    TC = TC & "+" & Me.txtFB.Text.Split("+")(j)
                Next
                Me.GridTKT.Item("FareBasis", RowNo).Value = TC.Substring(1)
                Me.GridTKT.Item(0, RowNo).Style.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.txtTKNO.Text = DefineNextTKNO(Me.txtTKNO.Text, MyAL.isTKTLess)
            Next
        ElseIf Not Me.GridTKT.CurrentRow.Cells("FTKT").Value Is Nothing Then
            If Me.GridTKT.CurrentRow.Cells("FTKT").Value.ToString.Length > 0 And varAction = "CHANGE" Then
                Dim tmpDOI As Date = Me.GridTKT.CurrentRow.Cells("DOI").Value
                k = Me.GridTKT.CurrentCell.RowIndex
                Do
                    k = k + 1
                    If k < Me.GridTKT.RowCount Then
                        Me.GridTKT.Item("DOI", k).Value = tmpDOI
                        If Me.GridTKT.Item("FTKT", k).Value.ToString.Length = 3 Then Exit Do
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        ChangeColorGridTKT("")
    End Sub
    Private Function GomTax_Charge(ByVal pT_C As String) As String
        Dim KQ As String = "", Txt As String, Lbl As String, Curr As String
        If pT_C = "T" Then
            For i As Int16 = 1 To 5
                Txt = "txtTax" & Trim(Str(i))
                Lbl = "txtTaxLabel" & Trim(Str(i))
                If Me.GrpTax.Controls(Lbl).Visible Then
                    KQ = KQ & "|" & Me.GrpTax.Controls(Lbl).Text & CDec(Me.GrpTax.Controls(Txt).Text)
                End If
            Next
        Else
            For i As Int16 = 1 To 4
                Txt = "txtCharge" & Trim(Str(i))
                Lbl = "CmbCharge" & Trim(Str(i))
                Curr = "CmbCCurr" & Trim(Str(i))
                If Me.GrpCharge.Controls(Txt).Text = "0" OrElse Me.GrpCharge.Controls(Txt).Text = "" Then
                    KQ = KQ & "|" & "...:" & Me.GrpCharge.Controls(Curr).Text & "0"
                Else
                    KQ = KQ & "|" & Me.GrpCharge.Controls(Lbl).Text & ":" & Me.GrpCharge.Controls(Curr).Text & Me.GrpCharge.Controls(Txt).Text
                End If
            Next
        End If
        If KQ <> "" Then KQ = KQ.Substring(1)
        Return KQ
    End Function
    Private Sub txtALCommPCT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtALCommPCT.LostFocus
        Dim PCT As Decimal, VAL As Double, f As Decimal
        f = CDec(Me.txtFare.Text)
        PCT = CDec(Me.txtALCommPCT.Text)
        If PCT > 0 Then Me.txtALCommVAL.Enabled = False
        If PCT > 0 Then
            VAL = (100 - PCT) * f / 100
            Me.txtNetToAL.Text = Format(VAL, "#,##0.00")
            Me.txtALCommPCT.Text = Format(PCT, "#,##0.00")
            Me.txtALCommVAL.Text = Format(f - VAL, "#,##0.00")
        End If
    End Sub

    Private Sub TxtAgtDiscPCT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        TxtAgtDiscPCT.LostFocus, TxtAgtDiscVal.LostFocus
        Dim PCT As Double, VAL As Double, F As Decimal, txt As TextBox = CType(sender, TextBox)
        Dim TVApie As Decimal = CDec(Me.txtFare.Text) - CDec(Me.txtNetToAL.Text)
        F = CDec(Me.txtFare.Text)
        If F = 0 Then Exit Sub
        If txt.Name = "TxtAgtDiscPCT" Then
            PCT = CDbl(txt.Text)
            VAL = PCT * F / 100
        ElseIf txt.Name = "TxtAgtDiscVal" Then
            VAL = CDbl(txt.Text)
            PCT = VAL / F * 100
        End If
        If PCT = 0 Or VAL = 0 Then Exit Sub
        Me.TxtAgtDiscVal.Text = Format(VAL, "#,##0.00")
        Me.TxtAgtDiscPCT.Text = Format(PCT, "#,##0.00")
        If pubVarSRV = "S" Then
            If CDec(Me.TxtAgtDiscVal.Text) > TVApie Then
                txt.Focus()
            End If
        ElseIf pubVarSRV = "R" And CDec(Me.TxtAgtDiscVal.Text) < CDec(SalesInforOfThisTKT.Discount) Then
            Me.TxtAgtDiscVal.Focus()
        End If
    End Sub

    Private Sub CmbTVDiscount1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        CmbTVDiscount1.SelectedIndexChanged, CmbTVDiscount2.SelectedIndexChanged, _
        CmbTVCharge1.SelectedIndexChanged, CmbTVCharge2.SelectedIndexChanged
        Dim cmb As ComboBox = CType(sender, ComboBox), txtName As String
        txtName = "txt" & cmb.Name.Substring("cmb".Length, cmb.Name.Length - "cmb".Length)
        If cmb.Text <> "..." Then
            cmb.Parent.Controls(txtName).Enabled = True
        Else
            cmb.Parent.Controls(txtName).Enabled = False
        End If
    End Sub

    Private Sub txtTVDiscount1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        txtTVDiscount1.LostFocus, txtTVCharge1.LostFocus, txtTVDiscount2.LostFocus, txtTVCharge2.LostFocus
        Dim aa As Double, cbName As String
        Dim txt As TextBox = CType(sender, TextBox)
        aa = CDbl(txt.Text)
        txt.Text = Format(aa, "#,##0.00")

        cbName = "Cmb" & txt.Name.Substring("txt".Length, txt.Name.Length - 1 - "txt".Length) & "2"
        If aa > 0 Then
            txt.Parent.Controls(cbName).Enabled = True
        Else
            txt.Parent.Controls(cbName).Enabled = False
        End If
        CalNettoALtoVND()
    End Sub
    Private Sub GridFOP_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridFOP.CellEnter
        Dim R As Integer, varROE As Decimal, varCurr As String, BasedCurr As String
        Dim TTLVND As Decimal
        BasedCurr = Me.CmbCurr.Text
        R = e.RowIndex
        For i As Int16 = Me.GridFOP.RowCount - 1 To R + 1 Step -1
            Me.GridFOP.Item(3, i).Value = 0
            Me.GridFOP.Item(4, i).Value = ""
        Next
        For i As Int16 = 0 To Me.GridFOP.RowCount - 1 ' Set lai cac o tren dong curr tranh NV sua nguoc
            If InStr("DEP_DDB_PPD_PSP", Me.GridFOP.Item(0, i).Value) > 0 Then
                Me.GridFOP.Item(1, i).ReadOnly = True
                Me.GridFOP.Item(2, i).ReadOnly = True
                If Me.GridFOP.Item(0, i).Value = "DEP" Then
                    Me.GridFOP.Item(3, i).ReadOnly = True
                    Me.GridFOP.Item(4, i).ReadOnly = True
                End If
            ElseIf InStr("CSH_RND", Me.GridFOP.Item(0, i).Value) > 0 Then
                Me.GridFOP.Item(4, i).Value = ""
            End If
        Next
        TTLVND = CDec(Me.txtVNDEquivalent.Text)
        If e.ColumnIndex < 4 Then
            Me.CmdSave.Enabled = False
        End If
        ' phan if nay de set tinh trang readonly voi 1 so FOP dac thu
        If e.ColumnIndex = 1 Then
            If InStr("CRD_UCF_PTA_MCO_EXC", Me.GridFOP.Item(0, R).Value) > 0 Then
                Me.GridFOP.Item(1, R).Value = "USD"
            End If
            If InStr("DEB_PSP_PPD", Me.GridFOP.Item(0, R).Value) > 0 Then
                Me.GridFOP.Item(1, R).ReadOnly = True
                Me.GridFOP.Item(2, R).ReadOnly = True
                If InStr("DEB_PSP", Me.GridFOP.Item(0, R).Value) > 0 Then
                    Me.GridFOP.Item(1, R).Value = Me.CmbCurr.Text
                ElseIf InStr("PPD", Me.GridFOP.Item(0, R).Value) > 0 Then
                    Me.GridFOP.Item(1, R).Value = "VND"
                End If
            Else
                Me.GridFOP.Item(1, R).ReadOnly = False
            End If
        ElseIf e.ColumnIndex = 2 Then
            If InStr("DDB_DEB_PSP_PPD_EXC_PTA_MCO_CRD_UCF", Me.GridFOP.Item(0, R).Value) > 0 Then
                Me.GridFOP.Item(2, R).ReadOnly = True
                Me.GridFOP.Item(4, R).ReadOnly = False
                Me.GridFOP.Item(3, R).ReadOnly = False
            End If
            varCurr = Me.GridFOP.Item(1, R).Value
            If isIATARate Then
                varROE = IIf(varCurr = "VND", 1, ForEX_12(Me.TxtDOS.Value, varCurr, IIf(BasedCurr = "VND", "BBR", "BSR"), "IATA"))
            Else
                varROE = IIf(varCurr = "VND", 1, ForEX_12(Me.TxtDOS.Value, varCurr, IIf(BasedCurr = "VND", "BBR", "BSR"), MySession.TRXCode))
            End If
            If varCurr = Me.CmbCurr.Text Then 'And pubVarSRV = "R" Then
                Me.GridFOP.Item(2, R).Value = CDec(Me.LcktxtROE.Text)
            Else
                Me.GridFOP.Item(2, R).Value = varROE
            End If
        ElseIf e.ColumnIndex = 3 Then
            If InStr("DDB_DEP_DEB_PSP_PPD_EXC_PTA_MCO_CRD_UCF", Me.GridFOP.Item(0, R).Value) > 0 Then
                If InStr("UCF_PTA_MCO", Me.GridFOP.Item(0, R).Value) > 0 And _
                    Me.GridFOP.Item(1, R).Value <> "USD" Then
                    MsgBox("Invalid FOP/Curr", MsgBoxStyle.Critical, msgTitle)
                    Exit Sub
                ElseIf InStr("DEB_PSP_PPD_DDB", Me.GridFOP.Item(0, R).Value) > 0 Then
                    If (InStr("DEB_PSP_DDB_EXC", Me.GridFOP.Item(0, R).Value) > 0 And _
                        Me.GridFOP.Item(1, R).Value <> Me.CmbCurr.Text) Or _
                        (InStr("PPD", Me.GridFOP.Item(0, R).Value) > 0 And Me.GridFOP.Item(1, R).Value <> "VND") Then
                        MsgBox("Invalid FOP/Curr", MsgBoxStyle.Critical, msgTitle)
                        Exit Sub
                    End If
                End If
            End If
        End If


        ' phan if nay de tinh gia tri amount, paid, enable cmdSave
        If e.ColumnIndex = 3 Or e.ColumnIndex = 4 Then
            DaTra = 0
            For i As Int16 = 0 To R - 1 + e.ColumnIndex - 3
                DaTra = DaTra + Me.GridFOP.Item(3, i).Value * Me.GridFOP.Item(2, i).Value
            Next
            If e.ColumnIndex = 3 Then
                Me.GridFOP.Item(3, R).Value = (TTLVND - DaTra) / Me.GridFOP.Item(2, R).Value
            End If
        End If
        Me.TxtPaid.Text = Format(DaTra, "#,##0.00")
        If Me.TxtPaid.Text = Me.txtVNDEquivalent.Text Then Me.CmdSave.Enabled = True
        Me.GridFOP.Columns(2).DefaultCellStyle.Format = "#,##0.00"
        Me.GridFOP.Columns(3).DefaultCellStyle.Format = "#,##0.00"
    End Sub

    Private Sub TxtPromoCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTourCode.Enter, TxtPaxName.Enter
        Dim txt As TextBox = CType(sender, TextBox)
        If txt.Text = "TOUR CODE" Or txt.Text = "PAX NAME" Or txt.Text = "STAY" Then
            txt.Text = ""
            txt.ForeColor = Color.Black
        End If
    End Sub
    Private Sub FillLstTKTinRCP(ByVal varSQL As String)
        Dim dtbl As DataTable = GetDataTable(varSQL)
        For i As Int16 = 0 To dtbl.Rows.Count - 1
            Me.LstTKTinRCP.Items.Add(dtbl.Rows(i)("SRV") & "- " & dtbl.Rows(i)("TKNO"))
        Next
    End Sub
    Private Sub LoadData(ByVal parRCPID As Integer)
        Dim dTable As DataTable
        Dim tmpRCPno As String = "", DK_RVC As String, strSQL As String
        Dim Booker As String = "", DK_S As String
        Dim HasINVID As Integer, AR_RPTNO As String
        Me.CmdRVSelected.Enabled = False
        Me.LstTKTinRCP.Items.Clear()

        If mRCP <> "" And WhatAction <> "PRN" Then
            AR_RPTNO = ScalarToString("RCP", "RPTNO", " recid=" & parRCPID)
            If AR_RPTNO = "" Then
                HasINVID = ScalarToInt("FOP", "ProfID", " RCPID=" & parRCPID)
            End If
            If HasINVID > 0 Or AR_RPTNO <> "" Then
                MsgBox("This Transaction Has Been Either Reported to Accouting or Invoiced To Cust. and Locked. Ask Authorized People to Unlock", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
            Dim EventNonAir As String = ScalarToString("FOP", "Document", "rcpid=" & parRCPID & " and status='OK' and FOP='MCE'")
            If Not String.IsNullOrEmpty(EventNonAir) Then
                EventNonAir = ScalarToString("Dutoan_tour", "status", "Tcode='" & EventNonAir & "'")
                If EventNonAir = "RR" Then
                    MsgBox("This Event Has Been Finalised. Ask Authorized People to Unlock", MsgBoxStyle.Critical, msgTitle)
                    Exit Sub
                End If
            End If
        End If
        dTable = GetDataTable("select * from RCP where RecID=" & parRCPID & " and city='" & MySession.City & _
                              "' and CustType in " & myStaff.CAccess)
        If dTable.Rows.Count = 0 Then
            FeedBack("Error. Transaction Not Found! ")
            Me.CmdPreview.Enabled = False
            Me.CmdDelete.Enabled = False
            Exit Sub
        Else
            Me.CmdPreview.Enabled = True
            Me.CmdDelete.Enabled = True
        End If

        tmpRCPno = dTable.Rows(0)("RCPNO").ToString.Trim
        Me.TxtDOS.Value = dTable.Rows(0)("DOS").ToString.Trim
        Me.txtDiscountedAmount.Text = Format(dTable.Rows(0)("Discount"), "#,##0.00")

        Me.CmbCustType.Text = ("CustType")
        MyAL.ALCode = dTable.Rows(0)("Stock")
        MyCust.CustID = dTable.Rows(0)("CustID")
        If mRCP <> "" Then
            pubVarSRV = dTable.Rows(0)("SRV")
            If pubVarSRV = "O" Then pubVarSRV = "R"
            Me.StatusSRV.Text = "Transaction: " & pubVarSRV
        End If
        Me.CmbAL.Text = MyAL.ALCode
        Me.CmbCurr.Text = dTable.Rows(0)("Currency")
        Me.LcktxtROE.Text = Format(dTable.Rows(0)("ROE"), "#,##0.00")
        GenCmbChargeCurr(Me.CmbCurr.Text)
        Me.GrpComm.Enabled = GenCOC(MyCust.CustID, MyAL.ALCode)

        LoadCmb_VAL(Me.CmbChannel, "select " & MyCust.CustID & " as val,'" & dTable.Rows(0)("CustShortName") & "' as DIS ")
        Me.CmbChannel.SelectedValue = MyCust.CustID
        Me.CmbChannel.Text = dTable.Rows(0)("CustShortName")
        Me.CmbCustType.Text = MyCust.CustType
        Me.txtCustName.Text = dTable.Rows(0)("PrintedCustName")
        Me.txtCustAddrr.Text = dTable.Rows(0)("PrintedCustAddrr")
        Me.txtTaxCode.Text = dTable.Rows(0)("PrintedTaxCode")

        If mRCP = "" Then
            DK_S = " and SRV='S'"                   ' Truong hop tim de hoan huy => lay SRV='S'
            DK_RVC = " where SRV in ('R','V','C') and status <>'XX' and left(rcpno,2)='" & MySession.TRXCode & "')" ' Status theo tinh trang khach
        Else
            DK_S = " and SRV='" & pubVarSRV & "'"   ' Cac truong hop sua => lay SRV = SRV cua receipt
            DK_S = " "   ' sua do ve void deu RCP=R TKT=V thi ko tim thay
            DK_RVC = " where SRV='" & pubVarSRV & "' and " & DKstatusTKT & " and left(rcpno,2)='" & MySession.TRXCode & "')"
        End If
        strSQL = "select *  from TKT where RCPID=" & pubVarRCPID_BeingEdited & DK_S & " and " & DKstatusTKT

        If pubVarSRV = "C" Then
            strSQL = strSQL & " and rptno ='' "
        End If
        If mRCP = "" Then
            strSQL = strSQL & " and TKNO not in (select TKNO from TKT " & DK_RVC & " order by TKNO "
            FillLstTKTinRCP(strSQL)
        Else
            strSQL = strSQL & " order by TKNO "
            FillGridTKT(strSQL)
            If WhatAction <> "FOP" Then FillGridFOP(pubVarRCPID_BeingEdited)
        End If
        If Me.LstTKTinRCP.Items.Count = 0 And mRCP = "" Then
            FeedBack("Error. Ticket Has Already Been Refunded or Reported!")
        End If
    End Sub

    Private Sub FillGridFOP(ByVal parRCPID As Integer)
        Dim dTbl As DataTable = GetDataTable("select * from FOP where RCPID=" & pubVarRCPID_BeingEdited & " and " & DKstatusFOP)
        Dim RowNo As Int16
        If Me.FOP_FOP.Items.Count < 4 Then GenFOP_CURR()
        Me.GridFOP.Rows.Clear()
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            Me.GridFOP.Rows.Add()
            Me.GridFOP.Item("FOP_FOP", RowNo).Value = dTbl.Rows(i)("FOP")
            Me.GridFOP.Item("RCPCurrency", RowNo).Value = dTbl.Rows(i)("Currency")
            Me.GridFOP.Item("RCPROE", RowNo).Value = dTbl.Rows(i)("ROE")
            Me.GridFOP.Item("Amount", RowNo).Value = dTbl.Rows(i)("Amount")
            Me.GridFOP.Item("Document", RowNo).Value = dTbl.Rows(i)("Document")
            RowNo = RowNo + 1
        Next
    End Sub
    Private Sub CmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSearch.Click
        Dim strDK As String, searchWhat As String = Me.CmbSearchWhat.Text.Substring(0, 3)
        Dim SearchValue As String = Me.txtValueToSearch.Text
        Dim tmpAL As String = MyAL.DocToCode(SearchValue.Substring(0, 3))
        If (searchWhat = "TRX" And SearchValue.Substring(0, 2) <> MySession.TRXCode) Or _
           (searchWhat = "TKT" And InStr("AHC_GRP_HTL_INS", SearchValue.Substring(0, 3)) = 0 And tmpAL <> Me.CmbAL.Text) Then
            MsgBox("Invalid Doc No. Plz Check Your Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If searchWhat = "TRX" Then
            strDK = " RCPNO='" & SearchValue & "' and srv='S'"
            strDK = strDK & " and recid in (select rcpid from tkt where doctype not in ('AHC','GRP') and statusAL<>'XX') "
        Else
            If InStr("AHC_", SearchValue.Substring(0, 3)) > 0 Then
                strDK = "SRV='?'" ' cai nay ko dc hoan. dua vao dk deu de ko bao gio tim thay
            ElseIf SearchValue.Substring(0, 3) = "GRP" Then
                strDK = " RCPNO in (select F3 from Actionlog where F2='" & SearchValue & "' and tablename+DoWhat='RQDEPHANDLEOK_R')"
                strDK = strDK & " and srv+sbu='S" & MySession.Domain & "'"
            Else
                strDK = " RecID in (select RCPID from TKT where TKNO='" & SearchValue & "') and srv+sbu='S" & MySession.Domain & "' and status='OK'"
            End If
        End If

        pubVarRCPID_BeingEdited = ScalarToInt("RCP", "RecID", strDK)
        If searchWhat = "TKT" Then
            Me.txtValueToSearch.Text = ScalarToString("RCP", "RCPNO", " Recid=" & pubVarRCPID_BeingEdited)
        End If
        LoadData(pubVarRCPID_BeingEdited)
        If Me.LstTKTinRCP.Items.Count = 0 Then
            FeedBack("Error. Ticket Has Already Been Refunded or Reported!")
        End If
        If Me.OptRefund.Checked Or Me.OptPseudoVoid.Checked Then
            If Me.CmbAL.Text <> "NH" Then amtCRDPaid = calcAmtCrdPaid(pubVarRCPID_BeingEdited)
        End If
        If InStr("CS_LC", MyCust.CustType) > 0 Then
            LoadCmbBooker(MyCust.CustID)
        End If
    End Sub

    Private Function calcAmtCrdPaid(ByVal pRCPID As Integer) As Decimal
        Dim tmpPaid As Decimal, tmpChargeAlready As String = ""
        tmpPaid = ScalarToDec("FOP", "isnull(sum(amount*roe),0)/" & CDec(Me.LcktxtROE.Text), "fop='CRD' and status='OK' and rcpID=" & pRCPID)
        If tmpPaid > 0 Then
            tmpChargeAlready = ScalarToString("RCP", "charge_D", " RecID=" & pRCPID)
            If InStr(tmpChargeAlready, "CRD") > 0 Then tmpPaid = 0
        End If
        Return tmpPaid
    End Function
    Private Sub CmdNextFr1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdNextFr1.Click
        Me.txtTKNO.Enabled = True
        Me.txtTKNO.Text = MyAL.DocCode
        Me.TabControl1.SelectTab("pStep2")
    End Sub

    Private Sub txtNetToAL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNetToAL.LostFocus
        Dim n2al As Decimal, f As Decimal
        n2al = CDec(Me.txtNetToAL.Text)
        f = CDec(Me.txtFare.Text)
        Me.txtNetToAL.Text = Format(n2al, "#,##0.00")
        If InStr("FOC_NET_S+M", Me.CmbProduct.Text.Trim) = 0 Then Me.txtALCommVAL.Text = f - n2al
        Me.txtExcDoc.Enabled = IIf(Me.CmbDocType.Text = "EXC", True, False)
    End Sub

    Private Sub CmbDocType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbDocType.SelectedIndexChanged
        If Me.txtTKNO.Text <> "" Then
            ResetLabel4AIR()
            If Me.CmbDocType.Text = "GRP" Then
                Prepare4GRPdeposit()
            ElseIf Me.CmbDocType.Text = "HTL" Then
                Prepare4HTL()
            ElseIf Me.CmbDocType.Text = "CAR" Then
                Prepare4CAR()
            ElseIf Me.CmbDocType.Text = "AHC" Then
                Prepare4AHC()
            ElseIf Me.CmbDocType.Text = "ATK" Then
                Prepare4AHC()
            ElseIf Me.CmbDocType.Text = "INS" Then
                Prepare4INS()
            ElseIf Me.CmbDocType.Text = "VSA" Then
                Prepare4VSA()
            ElseIf Me.CmbDocType.Text = "TVS" Then
                Prepare4TVS()
            End If
            Me.txtCheckDigit.Enabled = Me.txtTKNO.Enabled
            If Not Me.txtTKNO.Enabled Then
                Me.LblPromoCode.Text = ""
                Me.CmbPromoCode.Items.Add("")
                Me.CmbPromoCode.Items.Add("HAN")
                Me.GrpTax.Enabled = False
                Me.GrpComm.Enabled = False
            End If
        End If
    End Sub
    Private Sub Prepare4TVS()
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, "TV")
        Me.txtTKNO.Enabled = False
        Me.LblSegment.Text = "OrgDoc"
        Me.LblPaxType.Text = "Rsn4ReIss"
        Me.TxtCharge1.Text = 0
        LoadCmb_MSC(Me.CmbPaxType, "Reason4TVS")
    End Sub
    Private Sub Prepare4CAR()
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, "TV")
        Me.txtTKNO.Enabled = False
        Me.LblSegment.Text = "RLOC"
        Me.LblFB.Text = "Task"
        Me.LblDOI.Text = "BDate"
        Me.LblDOF.Text = "PickUp"
        Me.LblPaxType.Text = "Seater"
        Me.LblBkgClass.Text = "Days"
        Me.TxtTourCode.Text = ""
        Me.TxtCharge1.Text = 0
        Me.txtFB.Visible = True
        Me.txtFB.Width = 111
        Me.CmbHTLName.Visible = Not Me.txtFB.Visible
        Me.CmbPaxType.Items.Clear()
        Me.TxtCountry.Width = 116
        Me.TxtCountry.Left = Me.txtTKNO.Left
        Me.TxtCountry.Visible = True
        Me.TxtCountry.Text = "VENDOR"
        Me.TxtCountry.ForeColor = Color.DarkGray
        LoadCmb_MSC(Me.CmbPaxType, "SEATER")
    End Sub

    Private Sub Prepare4VSA()
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, "TV")
        Me.txtTKNO.Enabled = False
        Me.LblSegment.Text = "RLOC"
        Me.LblFB.Text = "Embasy"
        Me.LblDOI.Text = "BDate"
        Me.LblDOF.Text = "DOE"
        Me.LblPaxType.Text = "VisaType"
        Me.LblBkgClass.Text = ""
        Me.TxtTourCode.Text = "STAY"
        Me.TxtCharge1.Text = 0
        Me.txtFB.Visible = False
        Me.CmbHTLName.Visible = Not Me.txtFB.Visible
        Me.CmbHTLName.Left = 45
        Me.CmbHTLName.Width = Me.CmbHTLName.Width + 45
        Me.CmbPaxType.Items.Clear()
        Me.TxtCountry.Width = 116
        Me.TxtCountry.Left = Me.txtTKNO.Left
        Me.TxtCountry.Visible = True
        Me.TxtCountry.Text = "VENDOR"
        Me.TxtCountry.ForeColor = Color.DarkGray
        LoadCmb_MSC(Me.CmbPaxType, "VISATYPE")
        LoadCmb_MSC(Me.CmbHTLName, "DSQ")
    End Sub
    Private Sub ResetLabel4AIR()
        Me.LblSegment.Text = "Segments"
        Me.LblFB.Text = "F.Basis"
        Me.LblDOI.Text = "DOI"
        Me.LblDOF.Text = "DOF"
        Me.LblPaxType.Text = "PaxType"
        Me.LblPromoCode.Text = "PromoCode"
        Me.LblBkgClass.Text = "BkgClass"
        Me.TxtPaxName.Text = "PAX NAME"
        Me.TxtTourCode.Text = "TOUR CODE"
        LoadCmb_MSC(Me.CmbPaxType, "PAXTYPE")
        Me.GrpTax.Enabled = True
        Me.GrpComm.Enabled = True
        Me.txtFB.Width = 226
        Me.txtFB.Visible = True
        Me.txtFB.Text = ""
        Me.txtTKNO.Enabled = True
        Me.CmbHTLName.Visible = False
        Me.TxtCountry.Visible = False
        Me.txtTKNO.Text = MyAL.DocCode
    End Sub
    Private Sub Prepare4INS()
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, "TV")
        Me.txtTKNO.Enabled = False
        Me.LblSegment.Text = "PolicyNo."
        Me.LblFB.Text = "Plan"
        Me.txtFB.Text = "TGD"
        Me.LblDOI.Text = "DOI"
        Me.LblPaxType.Text = "Product"
        Me.LblDOF.Text = "EffDate"
        Me.TxtPaxName.Text = "Insured Name"
        Me.LblBkgClass.Text = "NoOfInsured"
        Me.TxtBkgClass.Text = 1
        Me.TxtCharge1.Text = 0
        Me.CmbPaxType.Items.Clear()
        Me.CmbPaxType.Items.Add("TGD")
        Me.CmbPaxType.Text = "TGD"
    End Sub

    Private Sub Prepare4AHC()
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, "TV")
        Me.txtTKNO.Enabled = False
        Me.LblSegment.Text = "Descr."
        Me.LblFB.Text = "Descr."
        Me.LblDOF.Text = "C.Time"
        Me.LblPromoCode.Text = "Call Type"
        Me.TxtPaxName.Text = "Caller"
        Me.TxtTourCode.Text = "Phone Number"
        Me.TxtCharge1.Text = 0
        LoadCmb_MSC(Me.CmbPromoCode, "AHCType")
    End Sub
    Private Sub Prepare4ATK()
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, "TV")
        Me.txtTKNO.Enabled = False
        Me.LblSegment.Text = "Descr."
        Me.LblFB.Text = "Descr."
        Me.LblDOF.Text = "C.Time"
        Me.LblPromoCode.Text = "Call Type"
        Me.TxtPaxName.Text = "Caller"
        Me.TxtTourCode.Text = "Phone Number"
        Me.TxtCharge1.Text = 0
        LoadCmb_MSC(Me.CmbPromoCode, "AHCType")
    End Sub
    Private Sub Prepare4HTL()
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, "TV")
        Me.txtTKNO.Enabled = False
        Me.LblSegment.Text = "RLOC"
        Me.LblFB.Text = "Name"
        Me.LblDOI.Text = "BDate"
        Me.LblDOF.Text = "ChckIn"
        Me.LblPaxType.Text = "RoomType"
        Me.LblBkgClass.Text = "Room/Nite"
        Me.TxtTourCode.Text = ""
        Me.TxtCharge1.Text = 0
        Me.txtFB.Visible = False
        Me.CmbHTLName.Visible = Not Me.txtFB.Visible
        Me.TxtCountry.Visible = Me.CmbHTLName.Visible
        Me.TxtCountry.Width = 29
        Me.TxtCountry.Text = ""
        LoadCmb_MSC(Me.CmbPaxType, "select TVCode+' - '+ Description as VAL from cwt.dbo.GO_RoomTypeTV where status='OK' order by Description")
        LoadCmb_MSC(Me.CmbHTLName, "select citycode +'-'+HtlName as val from cwt.dbo.GO_HotelListTv where status='OK' ")
        Me.CmbPaxType.Width = 200
        Me.CmbHTLName.Left = 72
        Me.TxtCountry.Left = 45
        Me.CmbHTLName.Width = 196
    End Sub
    Private Sub Prepare4GRPdeposit()
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, "TV")
        Me.OptPrintInv.Enabled = False
        Me.GrpCharge.Enabled = False
        Me.CmbPromoCode.Enabled = False
        Me.txtTKNO.Enabled = False
        Me.TxtPaxName.Text = "No Of Pax"
        Me.TxtTourCode.Text = "DeadLine AdvName"
        Me.LblTKNO.Text = "DocNo"
        Me.LblDOI.Text = "TKTG"
        Me.LblDOI.Left = 474
        If mRCP = "" Then Me.txtTKNO.Text = GenPseudoTKT(Me.CmbDocType.Text, Me.CmbAL.Text)
        Me.txtTKNO.Left = 165
        Me.txtTKNO.Width = 106
        Me.TxtCharge1.Text = 0
    End Sub

    Private Sub OptRefund_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        OptSell.Click, OptRefund.Click, OptVoid.Click, OptPseudoVoid.Click, OptServicing.Click, OptRefundOLD.Click, OptReIss.Click
        Dim opt As RadioButton = CType(sender, RadioButton)
        Me.GrpSRV.Enabled = False
        pubVarSRV = opt.Tag
        PrepareForOptSRV(Me)
        If opt.Name = "OptSell" Then
            Me.GrpNewRCP.Visible = True
            Me.CmbCustType.Focus()
            Me.BarTKTListSV.Enabled = True
            Me.LblAmountPayable.Text = "Amount Payable (1-2+3)"
            Me.lblDiscountedAmount.Visible = False
        ElseIf opt.Name = "OptRefundOLD" Then
            Me.GrpNewRCP.Visible = True
            Me.CmbCustType.Focus()
            Me.LblAmountPayable.Text = "Amount Payable (1-2+3)"
            Me.lblDiscountedAmount.Visible = False
        ElseIf opt.Name = "OptVoid" Or opt.Name = "OptServicing" Then
            Me.GrpNewRCP.Visible = True
            If opt.Name = "OptVoid" Then
                Me.txtRTG.Enabled = False
                Me.txtFB.Enabled = False
                Me.TxtPaxName.Enabled = False
                Me.TxtBkgClass.Enabled = False
                Me.CmbPaxType.Enabled = False
                Me.txtFltDate.Enabled = False
                Me.GrpFare.Enabled = False
            ElseIf opt.Name = "OptServicing" Then
                Me.GrpNewRCP.Enabled = True
                Me.GrpGenInfor.Enabled = True
                Me.GrpFare.Enabled = True
                Me.txtRTG.Enabled = True
                Me.txtFB.Enabled = True
                Me.TxtPaxName.Enabled = True
                Me.TxtBkgClass.Enabled = True
                Me.CmbPaxType.Enabled = True
                Me.txtFltDate.Enabled = True
            End If
            Me.GrpTax.Enabled = False
            Me.GrpComm.Enabled = False
            Me.TxtTourCode.Enabled = False
            Me.CmbPromoCode.Enabled = False
        Else
            Me.CmbChannel.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCharge1.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCharge2.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCharge3.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCharge4.DropDownStyle = ComboBoxStyle.DropDown
            Me.TxtALR_Ch.Visible = True
            Me.TxtTVR_Ch.Visible = True
            Me.CmdNextFr1.Visible = False
            Me.GrpSearch.Visible = True
            Me.CmdRVSelected.Visible = True
            Me.lblDiscountedAmount.Visible = True
            Me.LblAmountPayable.Text = "Amount Payable (1+2-3-4)"
            If opt.Name = "OptRefund" Then
                Me.CmdRVSelected.Text = "Refund Selected Tickets"
            ElseIf opt.Name = "OptPseudoVoid" Then
                Me.CmdRVSelected.Text = "Void Selected Tickets"
                Me.GrpGenInfor.Enabled = False
                Me.GrpTax.Enabled = False
                Me.GrpFare.Enabled = False
            End If

            Me.CmbSearchWhat.Text = Me.CmbSearchWhat.Items(0).ToString
        End If
        Me.txtDiscountedAmount.Visible = Me.lblDiscountedAmount.Visible
        Me.CmbTVDiscount1.Enabled = Not Me.lblDiscountedAmount.Visible
        Me.StatusSRV.Text = "Transaction: " & pubVarSRV
        CanEditROE = CheckEditROE(Me.CmbAL.Text)
        Me.LcktxtROE.ReadOnly = Not CanEditROE
        If MyAL.isMultiCurr Then
            Me.CmbCurr.Enabled = True
        Else
            Me.CmbCurr.Enabled = False
        End If
        Me.CmbCurr.Text = MyAL.DefaultCurr
    End Sub

    Private Sub GridTKT_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridTKT.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Dim GrpFTCstatus As Boolean
        If WhatAction = "MON" Or (mRCP = "" And InStr("SROC", pubVarSRV) > 0) Then
            If Me.GridTKT.Item("FTKT", e.RowIndex).Value Is Nothing OrElse _
                Me.GridTKT.Item("FTKT", e.RowIndex).Value.ToString.Trim = "" OrElse _
                Me.GridTKT.Item("FTKT", e.RowIndex).Value.ToString.Substring(0, 3) = "___" Then
                GrpFTCstatus = True
            Else
                GrpFTCstatus = False
            End If
        Else
            GrpFTCstatus = False
        End If
        Me.GrpFare.Enabled = GrpFTCstatus
        Me.GrpTax.Enabled = GrpFTCstatus
        Me.GrpCharge.Enabled = GrpFTCstatus
        Me.GrpComm.Enabled = GrpFTCstatus

        Me.GrpFTKT2Edit.Visible = False
        refreshTKTcontent(e.RowIndex)
    End Sub
    Private Function XacDinhCoPhaiVeExc(ByVal pTKNo As String) As Boolean
        Dim SoVeTruocDo As String
        SoVeTruocDo = ScalarToString("TKT", "StockCtrl", " TKNO='" & pTKNo & "' and status<>'XX'")
        Return IIf(SoVeTruocDo = "", False, True)
    End Function
    Private Sub refreshTKTcontent(ByVal R As Integer)
        Dim tmpMsg As String, Amt2AL As Decimal

        If WhatAction <> "NON" Then
            Me.CmdSaveChange.Enabled = False
        Else
            Me.CmdSaveChange.Enabled = True
        End If
        Me.CmdSaveChange.Visible = True
        Me.CmdAddTKT.Visible = False
        If R < 0 Then Exit Sub
        If Me.GridTKT.RowCount = 0 Then Exit Sub
        If WhatAction = "" And InStr("SOA", pubVarSRV) > 0 Then
            Me.BarRemoveTKT.Enabled = True
        Else
            Me.BarRemoveTKT.Enabled = False
        End If

        Me.BarCopyTKT.Enabled = Me.BarRemoveTKT.Enabled
        Me.txtTKNO.Enabled = False

        ClearInput2(Me)
        Me.CmdAddTKT.Enabled = False

        Me.txtDOI.Value = Me.GridTKT.Item("DOI", R).Value
        If pubVarSRV = "R" Then
            SalesInforOfThisTKT.Comm = Me.GridTKT.Item("CommVAL", R).Value
            SalesInforOfThisTKT.Rtg = Me.GridTKT.Item("Routing", R).Value.ToString.Trim
            SalesInforOfThisTKT.Tax = Me.GridTKT.Item("Tax", R).Value
            SalesInforOfThisTKT.Discount = Me.GridTKT.Item("AgtDisctVAL", R).Value
            SalesInforOfThisTKT.Fare = Me.GridTKT.Item("Fare", R).Value
            SalesInforOfThisTKT.NetToAL = Me.GridTKT.Item("NetToAL", R).Value
            SalesInforOfThisTKT.isEXC = XacDinhCoPhaiVeExc(Me.GridTKT.Item("TKNO", R).Value)
            If mRCP = "" Then
                Me.txtDOI.Value = Now.Date
            End If
        End If
        Me.txtFltDate.Text = Me.GridTKT.Item("DOF", R).Value
        Me.txtRTG.Text = Me.GridTKT.Item("Routing", R).Value
        Me.TxtBkgClass.Text = Me.GridTKT.Item("BkgClass", R).Value
        Me.txtFB.Text = Me.GridTKT.Item("FareBasis", R).Value
        Me.CmbDocType.Text = Me.GridTKT.Item("DocType", R).Value
        Me.CmbPaxType.Text = Me.GridTKT.Item("PaxType", R).Value
        Me.txtTKNO.Text = Me.GridTKT.Item("TKNO", R).Value.ToString.Trim
        Me.TxtTourCode.Text = Me.GridTKT.Item("Tourcode", R).Value
        Me.CmbPromoCode.Text = Me.GridTKT.Item("PromoCode", R).Value
        Me.txtExcDoc.Text = Me.GridTKT.Item("StockCtrl", R).Value
        Me.CmbProduct.Text = Me.GridTKT.Item("FareType", R).Value
        Me.TxtPaxName.Text = Me.GridTKT.Item("PaxName", R).Value
        Me.ChkBizTrip.Checked = False
        If InStr("CS_LC", Me.CmbCustType.Text) > 0 AndAlso Me.GridTKT.Item("TKT_RMK", R).Value <> "" And Not Me.GridTKT.Item("TKT_RMK", R).Value Is Nothing Then
            For i As Int16 = 0 To UBound(Me.GridTKT.Item("TKT_RMK", R).Value.ToString.Split("|"))
                If Me.GridTKT.Item("TKT_RMK", R).Value.ToString.Split("|")(i) = "BIZT" Then
                    Me.ChkBizTrip.Checked = True
                ElseIf Me.GridTKT.Item("TKT_RMK", R).Value.ToString.Split("|")(i) = "BIZF" Then
                    Me.ChkBizTrip.Checked = False
                End If
                If InStr(Me.GridTKT.Item("TKT_RMK", R).Value.ToString.Split("|")(i), "BKR") Then
                    Me.CmbBooker.Text = Me.GridTKT.Item("TKT_RMK", R).Value.ToString.Split("|")(i).Substring(3)
                End If
            Next
        End If

        If Me.GridTKT.Item(0, R).Style.Alignment = DataGridViewContentAlignment.MiddleRight Then
            Me.GrpGenInfor.Enabled = False
        Else
            Me.GrpGenInfor.Enabled = True
            Me.txtFare.Text = Format(Me.GridTKT.Item("Fare", R).Value, "#,##0.00")
            Me.txtALCommPCT.Text = Format(Me.GridTKT.Item("CommPCT", R).Value, "#,##0.00")
            Me.txtALCommVAL.Text = Format(Me.GridTKT.Item("CommVAL", R).Value, "#,##0.00")
            Me.txtNetToAL.Text = Format(Me.GridTKT.Item("NetToAL", R).Value, "#,##0.00")
            Me.TxtAgtDiscPCT.Text = Format(Me.GridTKT.Item("AgtDisctPCT", R).Value, "#,##0.00")
            Me.TxtAgtDiscVal.Text = Format(Me.GridTKT.Item("AgtDisctVAL", R).Value, "#,##0.00")
            Me.txtShownFare.Text = Format(Me.GridTKT.Item("ShownFare", R).Value, "#,##0.00")
            Me.txtLLF.Text = Format(Me.GridTKT.Item("LLF", R).Value, "#,##0.00")
            Me.TxtTax.Text = Format(Me.GridTKT.Item("Tax", R).Value, "#,##0.00")
            TachTax(Me.GridTKT.Item("Tax_D", R).Value)
            TachCharge(Me.GridTKT.Item("charge_D", R).Value)
            Amt2AL = CDec(Me.txtNetToAL.Text) + CDec(Me.TxtTax.Text)
            If pubVarSRV <> "R" Then
                Amt2AL = Amt2AL + DefineAmtCharge2AL()
            Else
                Amt2AL = Amt2AL - DefineAmtCharge2AL()
            End If
            Me.LblAmt2AL.Text = "TTL Amt To AL: " & Format(Amt2AL, "#,##0.00") & " (VND" & Format(Amt2AL * CDec(Me.LcktxtROE.Text), "#,##0.00") & ")"
        End If
        If pubVarSRV = "R" Or pubVarSRV = "C" Then
            tmpMsg = XacDinhLuuYkhiRefundPromoTKT(Me.CmbPromoCode.Text)
            If tmpMsg <> "" Then MsgBox(tmpMsg, MsgBoxStyle.Information, msgTitle)
        End If
    End Sub
    Private Function DefineAmtCharge2AL() As Decimal
        Dim KQ As Decimal
        For i As Int16 = 1 To 4
            If Me.GrpCharge.Controls("CmbCharge" & i.ToString.Trim).Text.Substring(0, 2) = "SP" Then
                KQ = KQ + CDec(Me.GrpCharge.Controls("TxtCharge" & i.ToString.Trim).Text)
            End If
        Next
        Return KQ
    End Function
    Private Function XacDinhLuuYkhiRefundPromoTKT(ByVal pPromoCode As String) As String
        Return ScalarToString("promocode", "top 1 MSG", "promocode='" & pPromoCode & "' order by recid DESC")
    End Function
    Private Sub TachCharge(ByVal pChargeD As String)
        Dim Txt As String, Cmb As String, Curr As String
        If pChargeD = "..........0" Then Exit Sub
        For i As Int16 = 0 To UBound(pChargeD.Split("|"))
            Cmb = "CmbCharge" & Trim(Str(i + 1))
            Txt = "TxtCharge" & Trim(Str(i + 1))
            Curr = "CmbCCurr" & Trim(Str(i + 1))
            If pChargeD.Split("|")(i) <> "" Then
                If InStr(pChargeD, ":") > 0 Then
                    Me.GrpCharge.Controls(Cmb).Text = pChargeD.Split("|")(i).Split(":")(0)
                    Me.GrpCharge.Controls(Txt).Text = pChargeD.Split("|")(i).Split(":")(1).Substring(3)
                    Me.GrpCharge.Controls(Curr).Text = pChargeD.Split("|")(i).Split(":")(1).Substring(0, 3)
                Else
                    Me.GrpCharge.Controls(Cmb).Text = pChargeD.Split("|")(i).Substring(0, 10)
                    Me.GrpCharge.Controls(Txt).Text = pChargeD.Split("|")(i).Substring(10)
                    Me.GrpCharge.Controls(Curr).Text = Me.CmbCurr.Text
                End If
            End If
        Next
    End Sub
    Private Sub TachTax(ByVal pTaxD As String)
        Dim Lbl As String, Txt As String
        If pTaxD Is Nothing OrElse pTaxD.Trim = "" Then Exit Sub
        For i As Int16 = 1 To 5
            Txt = "TxtTax" & Trim(Str(i))
            Me.GrpTax.Controls(Txt).Text = 0
        Next
        For i As Int16 = 0 To UBound(pTaxD.Split("|"))
            Lbl = "TxtTaxLabel" & Trim(Str(i + 1))
            Txt = "TxtTax" & Trim(Str(i + 1))
            Me.GrpTax.Controls(Lbl).Text = pTaxD.Split("|")(i).Substring(0, 2)
            Me.GrpTax.Controls(Txt).Text = Format(CDec(pTaxD.Split("|")(i).Substring(2)), "#,##0.00")
        Next
    End Sub
    Private Sub FillGridTKT(ByVal parSQL As String)
        Dim dTbl As DataTable = GetDataTable(parSQL)
        Dim RowNo As Integer
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            Me.GridTKT.Rows.Add()
            RowNo = Me.GridTKT.Rows.Count - 1
            If mRCP <> "" Then Me.txtDOI.Value = dTbl.Rows(i)("DOI")
            Me.GridTKT.Item("TKNO", RowNo).Value = dTbl.Rows(i)("TKNO")
            If dTbl.Rows(i)("FTKT") <> "" AndAlso dTbl.Rows(i)("FTKT").ToString.Substring(0, 1) <> "_" Then
                Me.GridTKT.Item(0, RowNo).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
            Me.GridTKT.Item("RecID", RowNo).Value = dTbl.Rows(i)("RecID")
            Me.GridTKT.Item("isEdit", RowNo).Value = ""
            Me.GridTKT.Item("FTKT", RowNo).Value = dTbl.Rows(i)("FTKT")
            Me.GridTKT.Item("SRV", RowNo).Value = dTbl.Rows(i)("SRV")
            Me.GridTKT.Item("DOI", RowNo).Value = dTbl.Rows(i)("DOI")
            Me.GridTKT.Item("DOF", RowNo).Value = dTbl.Rows(i)("DOF")
            Me.GridTKT.Item("Routing", RowNo).Value = dTbl.Rows(i)("Itinerary")
            Me.GridTKT.Item("BkgClass", RowNo).Value = dTbl.Rows(i)("BkgClass")
            Me.GridTKT.Item("fareBasis", RowNo).Value = dTbl.Rows(i)("FareBasis")
            Me.GridTKT.Item("PaxName", RowNo).Value = dTbl.Rows(i)("PaxName")
            Me.GridTKT.Item("DocType", RowNo).Value = dTbl.Rows(i)("DocType")
            Me.GridTKT.Item("PaxType", RowNo).Value = dTbl.Rows(i)("PaxType")
            Me.GridTKT.Item("TourCode", RowNo).Value = dTbl.Rows(i)("TourCode")
            Me.GridTKT.Item("PromoCode", RowNo).Value = dTbl.Rows(i)("PromoCode")
            Me.GridTKT.Item("FareType", RowNo).Value = dTbl.Rows(i)("FareType")
            Me.GridTKT.Item("Fare", RowNo).Value = dTbl.Rows(i)("Fare")
            Me.GridTKT.Item("ShownFare", RowNo).Value = dTbl.Rows(i)("ShownFare")
            Me.GridTKT.Item("NetToAL", RowNo).Value = dTbl.Rows(i)("NetToAL")
            Me.GridTKT.Item("AgtDisctPCT", RowNo).Value = dTbl.Rows(i)("AgtDisctPCT")
            Me.GridTKT.Item("AgtDisctVAL", RowNo).Value = dTbl.Rows(i)("AgtDisctVAL")
            Me.GridTKT.Item("CommPCT", RowNo).Value = dTbl.Rows(i)("CommPCT")
            Me.GridTKT.Item("CommVAL", RowNo).Value = dTbl.Rows(i)("CommVAL")
            Me.GridTKT.Item("Tax", RowNo).Value = dTbl.Rows(i)("Tax")
            Me.GridTKT.Item("Charge", RowNo).Value = dTbl.Rows(i)("Charge")
            Me.GridTKT.Item("ChargeTV", RowNo).Value = dTbl.Rows(i)("ChargeTV")
            Me.GridTKT.Item("Tax_D", RowNo).Value = dTbl.Rows(i)("Tax_D")
            Me.GridTKT.Item("Charge_D", RowNo).Value = dTbl.Rows(i)("Charge_D")
            Me.GridTKT.Item("CommOffer", RowNo).Value = dTbl.Rows(i)("CommOffer")
            Me.GridTKT.Item("StockCtrl", RowNo).Value = dTbl.Rows(i)("StockCtrl")
            Me.GridTKT.Item("TKT_RMK", RowNo).Value = dTbl.Rows(i)("RMK")
            Me.GridTKT.Item("LLF", RowNo).Value = dTbl.Rows(i)("LLF")
        Next
        ChangeColorGridTKT("")
        CalcRefundableCharges()
    End Sub
    Private Sub CmdRVSelected_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdRVSelected.Click
        Dim TKlist As String = "", strSQL As String
        For i As Int16 = 0 To Me.LstTKTinRCP.Items.Count - 1
            If Me.LstTKTinRCP.GetItemCheckState(i) = CheckState.Checked Then
                TKlist = TKlist & ",'" & Me.LstTKTinRCP.Items(i).ToString.Substring(3) & "'"
            End If
        Next
        TKlist = TKlist.Substring(1)
        TKlist = "(" & TKlist & ")"
        Me.GridTKT.Rows.Clear()
        strSQL = "Select * from TKT where rcpID=" & pubVarRCPID_BeingEdited & " and TKNO in " & TKlist
        strSQL = strSQL & " and SRV='S' and status='OK' and tkno not in (select TKNO from tkt where status<>'XX' and srv in('R','V') and left(rcpno,2)='" & MySession.TRXCode & "') "
        strSQL = strSQL & " order by tkno"
        FillGridTKT(strSQL)

        Me.GrpSearch.Visible = False
        Me.GrpNewRCP.Visible = True
        Me.GrpNewRCP.Enabled = False
        Me.GrpFare.Enabled = False
        Me.TabControl1.SelectTab("pStep2")
    End Sub

    Private Sub txtDiscountedAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscountedAmount.LostFocus
        Dim aa As Double = CDbl(txtDiscountedAmount.Text)
        txtDiscountedAmount.Text = Format(aa, "#,##0.00")
    End Sub

    Private Sub LstTKTinRCP_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles LstTKTinRCP.ItemCheck
        Dim i As Int16
        i = e.Index
        If Me.LstTKTinRCP.Items(i).ToString.Substring(0, 1) <> "S" Then
            If Me.OptRefund.Checked Or Me.OptVoid.Checked Then
                e.NewValue = CheckState.Unchecked
                FeedBack("Error. You Can't Refund Twice!  ")
            End If
        End If
    End Sub
    Private Sub txtTKNO_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTKNO.GotFocus
        If Me.txtTKNO.Text.Length = 0 Then Exit Sub
        Me.txtTKNO.Text = Me.txtTKNO.Text.Replace(" ", "")
    End Sub
    Private Sub txtTKNO_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        txtTKNO.KeyDown, TxtFTKT2Edit.KeyDown, txtExcDoc.KeyDown
        If e.KeyValue = 27 Then ESCPressed = Not ESCPressed
        If Not MyAL.isTKTLess Then
            CharEntered = checkCharEntered(e.KeyValue)
        End If
    End Sub
    Private Sub txtTKNO_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles _
        txtTKNO.KeyPress, TxtFTKT2Edit.KeyPress, txtExcDoc.KeyPress
        If Not MyAL.isTKTLess Then
            If CharEntered Then
                e.Handled = True
            End If
        End If
    End Sub
    Private Function isAlreadyInGrid(ByVal vTKNO As String) As Boolean
        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            If vTKNO = Me.GridTKT.Item("TKNO", i).Value.ToString.Trim Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Sub txtTKNO_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTKNO.LostFocus
        Dim TKNO As String, DaSD As Boolean = False, SoPax As Decimal
        If ESCPressed Then Exit Sub
        TKNO = Me.txtTKNO.Text.Trim

        If TKNO.Length <> 13 Then
            FeedBack("Error. Invalid TKNO format!")
            Me.txtTKNO.Focus()
            Exit Sub
        End If
        If Not TKNO.Contains("Z") Then
            If Me.CmbDocType.Text = "TKT" And MySession.Domain <> "EDU" Then
                Me.txtCheckDigit.Enabled = True
            Else
                Me.txtCheckDigit.Enabled = False
            End If
            TKNO = CheckTKTformat(TKNO)
            If TKNO.Substring(0, 3).ToUpper = "ERR" Then
                FeedBack("Error. Invalid Ticket Number! ")
                Me.txtTKNO.Focus()
                Exit Sub
            Else
                Me.txtTKNO.Text = TKNO
            End If
        Else
            Me.txtCheckDigit.Enabled = False
            On Error GoTo ErrorHandler
            SoPax = CDec(TKNO.Substring(TKNO.Length - 3, 3))
            On Error GoTo 0
        End If
        If pubVarSRV <> "A" Then DaSD = isAlreadyInGrid(TKNO)
        If DaSD Then
            FeedBack("Error. This Ticket Has Been In List!")
            Me.txtTKNO.Focus()
            Exit Sub
        End If
        If pubVarSRV = "V" Then Me.CmdAddTKT.Enabled = True
        Exit Sub
ErrorHandler:
        FeedBack("TicketLess Number Format is Booking/Res/Ref Number + 00n ! ")
        Me.txtTKNO.Focus()
    End Sub

    Private Sub txtFB_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFB.LostFocus
        If InStr(NonAirList, Me.CmbDocType.Text) > 0 Or pubVarSRV = "A" Then Exit Sub
        Dim RTG As String = Me.txtRTG.Text.Replace(" ", ""), tmpFB As String
        If ESCPressed Then Exit Sub
        tmpFB = IIf(Me.txtFB.Text = "", "Y", Me.txtFB.Text)
        If RTG.Length < 0 Then Exit Sub
        Me.txtFB.Text = FillFB_byRTG(RTG, tmpFB)
    End Sub

    Private Sub TxtBkgClass_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBkgClass.LostFocus
        If InStr(NonAirList, Me.CmbDocType.Text) > 0 Or pubVarSRV = "A" Then Exit Sub
        Dim RTG As String = Me.txtRTG.Text.Replace(" ", "")
        If Me.TxtBkgClass.Text = "" Then Me.TxtBkgClass.Text = "Y"
        Me.TxtBkgClass.Text = FillRBD_byRTG(RTG, Me.TxtBkgClass.Text)
    End Sub
    Private Sub GenCmbChargeCurr(pCurr As String)
        Dim strQry As String = "select '" & pCurr & "' as VAL"
        If InStr("CS_LC", MyCust.CustType) > 0 Then
            If pCurr <> "VND" Then strQry = strQry & " union select 'VND' as VAL "
            If pCurr <> "USD" Then strQry = strQry & " union select 'USD' as VAL "
        End If
        LoadCmb_MSC(Me.CmbCCurr1, strQry)
        LoadCmb_MSC(Me.CmbCCurr2, strQry)
        LoadCmb_MSC(Me.CmbCCurr3, strQry)
        LoadCmb_MSC(Me.CmbCCurr4, strQry)
    End Sub
    Private Sub CmbCurr_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCurr.LostFocus, CmbCurr.TextChanged
        Dim cmb As Control = CType(sender, ComboBox)
        Dim varROE As Decimal
        GenCmbChargeCurr(cmb.Text)

        If cmb.Text = "VND" Then
            varROE = 1
        Else
            varROE = ForEX_12(Me.TxtDOS.Value, cmb.Text, "BSR", MySession.TRXCode)
        End If
        If cmb.Name = "CmbCurr" Then Me.LcktxtROE.Text = varROE
        Me.CmbCCurr1.Text = cmb.Text
        Me.CmbCCurr3.Text = cmb.Text
        Me.CmbCCurr2.Text = cmb.Text
        Me.CmbCCurr4.Text = cmb.Text
    End Sub
    Private Sub XacDinhLaiTiGia()
        isIATARate = False
        If DefineBSPStock() And pubVarSRV = "S" And Not isFromTKTList And mRCP = "" Then
            isIATARate = True
            Me.LcktxtROE.Text = Format(ForEX_12(Me.TxtDOS.Value, "USD", "BSR", "IATA"), "#,##0.00")
        End If
    End Sub
    Private Sub pStep1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles pStep1.Enter
        ResetMenu()

        If Me.CmbCurr.Text <> "VND" Then XacDinhLaiTiGia() ' Bo sung DEC14 lay ti gia theo IATA voi cac ve xuat bang GDS

        CalNettoALtoVND()
        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            If InStr("GRP_VSA_HTL_AHC_INS_CAR", Me.GridTKT.Item("DocType", i).Value) > 0 Then
                Me.OptPrintInv.Checked = False
                Me.OptPrintInv.Enabled = False
                Exit For
            End If
        Next
    End Sub
    Private Sub CalNettoALtoVND()
        Dim Nett As Decimal, TTLD As Decimal, TTLC As Decimal
        Dim ftc As Decimal = 0, d As Decimal = 0, ROE As Decimal
        TTLC = CDec(Me.txtTVCharge1.Text) + CDec(Me.txtTVCharge2.Text)
        TTLD = CDec(Me.txtTVDiscount1.Text) + CDec(Me.txtTVDiscount2.Text)
        ROE = CDec(Me.LcktxtROE.Text)

        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            ftc = ftc + CDec(Me.GridTKT.Item("Fare", i).Value)
            ftc = ftc + CDec(Me.GridTKT.Item("Tax", i).Value)
            d = d + CDec(Me.GridTKT.Item("AgtDisctVAL", i).Value)
        Next

        If InStr("SVA", pubVarSRV) > 0 Then
            For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
                ftc = ftc + CDec(Me.GridTKT.Item("Charge", i).Value)
                ftc = ftc + CDec(Me.GridTKT.Item("ChargeTV", i).Value)
            Next
            Nett = ftc - d + TTLC - TTLD
        ElseIf InStr("ORC", pubVarSRV) > 0 Or WhatAction = "MON" Or WhatAction = "FOP" Then
            For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
                If Not (Not Me.GridTKT.Item("FTKT", i).Value Is Nothing AndAlso _
                    Me.GridTKT.Item("FTKT", i).Value <> "" AndAlso _
                    Me.GridTKT.Item("FTKT", i).Value.ToString.Substring(0, 3) <> "___") Then
                    ftc = ftc - CDec(Me.GridTKT.Item("Charge", i).Value)
                    ftc = ftc - CDec(Me.GridTKT.Item("ChargeTV", i).Value)
                End If
            Next
            Nett = ftc - d - TTLC + TTLD
        End If
        If pubVarSRV = "R" Then
            Nett = Nett + CDec(Me.TxtALR_Ch.Text) + CDec(Me.TxtTVR_Ch.Text)
        End If
        Me.TxtTTLfare.Text = Format(ftc, "#,##0.00")
        Me.txtTTLDiscount.Text = Format(d, "#,##0.00")
        Me.txtDue.Text = Format(ftc - d, "#,##0.00")
        Me.txtNettPayable.Text = Format(Nett, "#,##0.00")
        Me.txtVNDEquivalent.Text = Format(Nett * ROE, "#,##0.00")
        If Me.txtNettPayable.Text <> Me.TxtPaid.Text And Me.txtNettPayable.Text <> "0" Then
            Me.CmdSave.Enabled = False
            If mRCP = "" Then
                Me.CmdPreview.Enabled = False
            Else
                Me.CmdPreview.Enabled = True
            End If
        End If
        If mRCP = "" Then SetDefaultFOP()
    End Sub

    Private Sub GridFOP_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles GridFOP.CellValidating
        Dim tmp As String
        If QuitNow Then Exit Sub
        If e.ColumnIndex = 3 Or e.ColumnIndex = 2 Then
            Me.GridFOP.Rows(e.RowIndex).ErrorText = ""
            Dim newInteger As Decimal, ErrText As String = "Invalid Input"
            If GridFOP.Rows(e.RowIndex).IsNewRow Then Return
            If Not Decimal.TryParse(e.FormattedValue.ToString(), newInteger) Then
                e.Cancel = True
                Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
            End If
        End If
        If e.ColumnIndex = 4 Then
            Me.GridFOP.Rows(e.RowIndex).ErrorText = ""
            Dim ErrText As String = "Invalid Document Number Format"
            If GridFOP.Rows(e.RowIndex).IsNewRow Then Return
            If (e.FormattedValue.ToString().Length <> 17 And _
                Me.GridFOP.Item(0, e.RowIndex).Value = "UCF") Or _
                (e.FormattedValue.ToString().Length <> 13 And _
                InStr("MCO_PTA", Me.GridFOP.Item(0, e.RowIndex).Value) > 0) Then
                e.Cancel = True
                Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
            End If
            If e.FormattedValue.ToString().Length <> 17 And _
                Me.GridFOP.Item(0, e.RowIndex).Value = "CRD" And MySession.TRXCode = "NH" Then
                e.Cancel = True
                Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
            End If
            If (e.FormattedValue.ToString().Length < 8 And _
                InStr("BTF_CHQ_ITP", Me.GridFOP.Item(0, e.RowIndex).Value) > 0) Then
                e.Cancel = True
                Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
            End If

            If Me.GridFOP.Item(0, e.RowIndex).Value = "EXC" Then
                If e.FormattedValue.ToString().Length <> 13 OrElse InStr(e.FormattedValue.ToString, "GRP") = 0 Then
                    e.Cancel = True
                    Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                    MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
                ElseIf InStr(ExcDocList, e.FormattedValue.ToString) = 0 Then
                    tmp = AddSpace2TKNO(e.FormattedValue.ToString)
                    If Not IsValidExcDoc(tmp, Me.GridFOP.Item(3, e.RowIndex).Value, Me.GridTKT.Item("DOF", 0).Value, pubVarRCPID_BeingEdited + pubVarRCPID_BeingCreated) Then
                        e.Cancel = True
                    End If
                End If
            End If
            If Me.GridFOP.Item(0, e.RowIndex).Value = "MCE" Then
                If InvalidTourCode(e.FormattedValue.ToString, MyCust.CustID, pubVarSRV, Me.GridTKT.Item("TKNO", 0).Value, IIf(mRCP = "", True, False), Me.GridTKT.Item("DOI", 0).Value, Me.GridTKT.Item("StockCtrl", 0).Value) Then
                    e.Cancel = True
                End If
            End If
            If Me.GridFOP.Item(0, e.RowIndex).Value = "3RD" Then
                If Invalid3Rd(e.FormattedValue.ToString, MyCust.CustID) Then
                    e.Cancel = True
                End If
            End If

            If Me.GridFOP.Item(0, e.RowIndex).Value = "VCR" Then
                If inValidVCR(e.FormattedValue.ToString, Me.GridFOP.Item(1, e.RowIndex).Value, Me.GridFOP.Item(3, e.RowIndex).Value) Then
                    e.Cancel = True
                End If
            End If
        End If
    End Sub
    Private Function inValidVCR(ByVal VCRNo As String, ByVal pCurr As String, ByVal pAMt As Decimal) As Boolean
        Dim VCRamt As Decimal, AmtUsed As Decimal
        cmd.CommandText = "select Amount from discountVCR where status+curr ='OK" & pCurr & "' and '" & Now.Date & "'" & _
            " between ValidFrom and ValidThru and VCRNo='" & VCRNo & "'"
        VCRamt = cmd.ExecuteScalar
        If VCRamt = 0 Then Return True
        cmd.CommandText = "select isnull(sum(Amount),0) from VCR_RAS where document ='" & VCRNo & "' and status='OK'"
        AmtUsed = cmd.ExecuteScalar
        cmd.CommandText = "select isnull(sum(Amount),0) from FLX_FOP where document ='" & VCRNo & "' and status='OK'"
        AmtUsed = AmtUsed + cmd.ExecuteScalar
        If pAMt > VCRamt - AmtUsed Then Return True
        Return False
    End Function

    Private Sub GridFOP_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles GridFOP.RowsAdded
        Me.GridFOP.Item(0, e.RowIndex).Value = "CSH"
        Me.GridFOP.Item(1, e.RowIndex).Value = "VND"
        Me.GridFOP.Item(2, e.RowIndex).Value = 1
        Me.GridFOP.Item(3, e.RowIndex).Value = 0
        Me.GridFOP.Item(4, e.RowIndex).Value = ""
    End Sub

    Private Sub CmbCustType_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCustType.LostFocus
        Dim StrDK As String = " and l.recID " & IIf(MySession.Domain = "EDU", "<0 ", ">0 ") & _
            " and l.recid in (select CustID from Cust_Detail where status+CAT='OKAL' and " & _
            " VAL in('YY','" & MySession.TRXCode & "'))"
        If Me.CmbCustType.Text = "WK" Then
            Me.CmbCommOffer.DataSource = Nothing
        End If
        Me.CmbBooker.Visible = False
        If CmbCustType.Text = "TA" Then
            LoadCmb_VAL(Me.CmbChannel, MyCust.List_TA & StrDK)
        ElseIf CmbCustType.Text = "TO" Then
            LoadCmb_VAL(Me.CmbChannel, MyCust.List_TO & StrDK)
        ElseIf CmbCustType.Text = "CA" Then
            LoadCmb_VAL(Me.CmbChannel, MyCust.List_CA & StrDK)
        ElseIf CmbCustType.Text = "LC" Then
            Me.CmbBooker.Visible = True
            LoadCmb_VAL(Me.CmbChannel, MyCust.List_LC & StrDK)
        ElseIf CmbCustType.Text = "CS" Then
            Me.CmbBooker.Visible = True
            LoadCmb_VAL(Me.CmbChannel, MyCust.List_CS & StrDK)
        ElseIf CmbCustType.Text = "WK" Then
            LoadCmb_VAL(Me.CmbChannel, MyCust.List_WK & StrDK)
        End If
    End Sub

    Private Sub CmbCharge1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        CmbCharge1.SelectedIndexChanged, CmbCharge2.SelectedIndexChanged, CmbCharge3.SelectedIndexChanged, _
        CmbCharge4.SelectedIndexChanged
        Try

            Dim cmb As ComboBox = CType(sender, ComboBox)
            Dim cbName As String, txtName As String, ctrlI As String
            Dim i As Int16, f As Decimal
            cbName = cmb.Name
            ctrlI = cbName.Substring("CmbCharge".Length, 1)
            txtName = "TxtCharge" & ctrlI
            If InStr(cmb.Text, "...") = 0 And cmb.Text <> "System.Data.DataRowView" Then
                If cmb.SelectedValue <> "0" Then
                    i = InStr(cmb.SelectedValue, "/S")
                    If i > 0 Then
                        i = cmb.SelectedValue.ToString.Length
                        f = CDec(cmb.SelectedValue.ToString.Substring(0, i - 2))
                        i = (Me.txtRTG.Text.Length - 3) / 7
                        cmb.Parent.Controls(txtName).Text = f * i
                    Else
                        cmb.Parent.Controls(txtName).Text = cmb.SelectedValue
                    End If
                    cmb.Parent.Controls(txtName).Enabled = False
                Else
                    cmb.Parent.Controls(txtName).Text = 0
                    cmb.Parent.Controls(txtName).Enabled = True
                End If
            Else
                cmb.Parent.Controls(txtName).Text = 0
                cmb.Parent.Controls(txtName).Enabled = False
            End If
            If cmb.Text.Length > 4 Then
                If cmb.Text.Substring(2, 1) = "R" Then
                    cmb.Parent.Controls(txtName).ForeColor = Color.Red
                Else
                    cmb.Parent.Controls(txtName).ForeColor = Color.Black
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub TxtCharge1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        TxtCharge1.LostFocus, TxtCharge2.LostFocus, TxtCharge3.LostFocus, TxtCharge4.LostFocus
        Dim txt As TextBox = CType(sender, TextBox)
        Dim aa As Decimal = CDec(txt.Text)
        txt.Text = Format(aa, "#,##0.00")
        If aa <= 0 Then
            If MyCust.CustID = 13234 Then
                MsgBox("Be Prudent. You've Input Negative Charge Amount", MsgBoxStyle.Critical, msgTitle)
                txt.ForeColor = Color.Red
            Else
                txt.Focus()
            End If
        End If
        Me.CmdAddTKT.Enabled = True
        Me.CmdSaveChange.Enabled = True
    End Sub
    Private Function DefineBSPStock() As Boolean
        Dim VeI As String, tmpKQ As Integer
        For i As Int16 = 0 To Me.GridTKT.RowCount - 1
            VeI = Me.GridTKT.Item("TKNO", i).Value.ToString
            VeI = VeI.Substring(4, 4)
            tmpKQ = ScalarToInt("MISC", "count(*)", " cat='BSPSTOCK' and val='" & VeI & "'")
            If tmpKQ = 0 Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub CmdPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdPreview.Click
        Dim fName As String, MyAnswer As Int16, InvToPrint As String = "", BSPStock As Boolean
        If MySession.Domain = "TVS" Then
            BSPStock = DefineBSPStock()
            If BSPStock And MyAL.CanIssVAT Then
                Me.OptPrintInv.Enabled = True
            Else
                Me.OptPrintTRX.Checked = True
                Me.OptPrintInv.Enabled = False
            End If
        End If
        If Not Me.OptPrintTRX.Checked And Not Me.OptPrintInv.Checked Then
            MsgBox("Please Select a Document You Want to Print.", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        fName = "R12_TRX_CFM.xlt"
        InHoaDon(Application.StartupPath, fName, "V", RCPNOtoPrint, Now.Date, Now.Date, 0, Me.CmbAL.Text, MySession.Domain, "")
        MyAnswer = MsgBox("If Input NOT Correct, Click Cancel." & vbCrLf & "If Correct. Do You Want to Print It? [Click Yes/No]", MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel Or MsgBoxStyle.DefaultButton2, msgTitle)
        If MyAnswer = vbCancel Then Exit Sub
        If Me.OptPrintInv.Checked And MyCust.VAT_Company <> "" And MyCust.VAT_Company <> MyAL.TVC Then   ' 1 so dai ly ft chi la khach hang cua 1 ai do. bo sung 3DEC15
            Me.OptPrintTRX.Checked = True ' chi cho in TRX
        End If
        'luon giu so HD, regardless click Y/N
        If Me.OptPrintInv.Enabled And pubVarSRV <> "V" And CDec(Me.txtVNDEquivalent.Text) <> 0 And mRCP = "" Then
            InvToPrint = TaoBanGhiHoaDon(RCPNOtoPrint, 0) 'luon chiem so HD (PDH y/c cho TS 10OCT10 va cho GSA 10APR11)
            If InvToPrint = "" Then
                MsgBox("Cant Create InvoiceNo for This TRX. Plz Use Another Menu or Report Accounting Staff.", MsgBoxStyle.Critical, msgTitle)
                Me.OptPrintTRX.Checked = True
                GoTo ExitHere
            End If
        End If
        If Me.OptPrintInv.Checked And MyCust.VAT_Company <> "" And MyCust.VAT_Company <> MyAL.TVC Then   ' 1 so dai ly ft chi la khach hang cua 1 ai do. bo sung 3DEC15
            Me.OptPrintInv.Enabled = False ' chi cho in TRX
        End If
        If MyAnswer = vbYes Then
            If OptPrintTRX.Checked Then
                InHoaDon(Application.StartupPath, fName, "O", RCPNOtoPrint, Now.Date, Now.Date, 0, Me.CmbAL.Text, MySession.Domain, "")
            ElseIf Me.OptPrintInv.Checked And pubVarSRV <> "V" And CDec(Me.txtVNDEquivalent.Text) <> 0 Then
                If mRCP = "" Then
                    fName = "R12_VATInvoice.xlt"
                    InHoaDon(Application.StartupPath, fName, "O", InvToPrint, Now.Date, Now.Date, 0, Me.CmbAL.Text, MySession.Domain, "")
                    cmd.CommandText = "update INV set printCopy=printCopy +1 where INVNO='" & InvToPrint & "'"
                    cmd.ExecuteNonQuery()
                Else
                    MsgBox("Plz Use Another Menu or Ask Accounting People to RePrint it", MsgBoxStyle.Information, msgTitle)
                End If
            End If
        End If
ExitHere:
        Me.Close()
        Dim f As New FOissueTKT
        f.ShowDialog()
    End Sub
    Private Function TaoBanGhiHoaDon(ByVal parRCPNO As String, ByVal PSoBanIn As Int16) As String
        Dim InvID As Integer, InvNo As String = ""
        Dim tmpFOP As String, MsgStr As String = "Unable To Update Table Invoice"
        InvNo = GenInvNo_QD153(Me.TxtTRXNO.Text, MyAL.VAT_KyHieu)
        InvID = Insert_INV("E", InvNo, InvNo.Substring(0, 2), pubVarRCPID_BeingCreated)
        tmpFOP = GetFOPstring()
        Try
            If myStaff.Counter = "CWT" Then
                TaoBanGhiTKTNO_INVNO_Standard(pubVarRCPID_BeingCreated, InvNo, InvID, "WO")
                Dim TTLAmt As Decimal = GetFareTaxChargeInVND(pubVarSRV, InvID)
                Dim dtblTVTR As DataTable = GetDataTable("select VAL, VAL1, VAL2 from MISC where cat='TVCompany' and description='TVT'")
                cmd.CommandText = Update_INV(pubVarSRV, dtblTVTR.Rows(0)("VAL"), dtblTVTR.Rows(0)("VAL1"), dtblTVTR.Rows(0)("VAL2"), _
                    TTLAmt, tmpFOP, PSoBanIn, InvID, MyCust.CustID)
            Else
                TaoBanGhiTKTNO_INVNO_Standard(pubVarRCPID_BeingCreated, InvNo, InvID, "WZ")
                cmd.CommandText = Update_INV(pubVarSRV, Me.txtCustName.Text, Me.txtCustAddrr.Text, Me.txtTaxCode.Text, _
                    CDec(Me.txtVNDEquivalent.Text), tmpFOP, PSoBanIn, InvID, MyCust.CustID)
            End If

            If PSoBanIn > 0 Then cmd.CommandText = cmd.CommandText & ";" & UpdateTblINVHistory(InvID, InvNo, PSoBanIn)
            cmd.ExecuteNonQuery()
            Return InvNo
        Catch ex As Exception
            MsgBox("Error In Creating Invoice", MsgBoxStyle.Critical, msgTitle)
            FileOpen(0, "d:\LogErrHD.txt", OpenMode.Append)
            Print(0, ex.Message)
            Print(0, cmd.CommandText)
            FileClose(0)
            Return ""
        End Try
    End Function

    Private Function GetFOPstring() As String
        Dim KQ As String = ""
        For i As Int16 = 0 To Me.GridFOP.RowCount - 1
            If CDec(Me.GridFOP.Item("Amount", i).Value) <> 0 Or Me.GridFOP.Item("Document", i).Value <> "" Then
                KQ = KQ & "|" & Me.GridFOP.Item("FOP_FOP", i).Value
                KQ = KQ & "_" & Me.GridFOP.Item("RCPCurrency", i).Value
                If Me.GridFOP.Item("Amount", i).Value Is Nothing Then
                    KQ = KQ & "_" & "0"
                Else
                    KQ = KQ & "_" & Me.GridFOP.Item("Amount", i).Value.ToString
                End If
                KQ = KQ & "_" & Me.GridFOP.Item("Document", i).Value
            End If
        Next
        Return KQ.Substring(1)
    End Function
    Private Sub LstTKTinRCP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LstTKTinRCP.SelectedIndexChanged
        If Me.LstTKTinRCP.CheckedItems.Count > 0 Then
            Me.CmdRVSelected.Enabled = True
        Else
            Me.CmdRVSelected.Enabled = False
        End If
    End Sub

    Private Function FindControl(ByVal parent As Control, ByVal name As String) As Control
        Dim thisctrl As Control = Nothing

        If parent.Name = name Then
            thisctrl = parent
        Else
            Dim ctrl As Control
            For Each ctrl In parent.Controls
                If ctrl.Name = name Then
                    thisctrl = ctrl
                    Exit For
                Else
                    If ctrl.Controls.Count > 0 Then
                        thisctrl = FindControl(ctrl, name)
                        If Not thisctrl Is Nothing Then
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        Return thisctrl
    End Function

    Private Sub txtFltDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFltDate.Enter
        Me.GrpFare.Enabled = False
        Me.TxtPaxName.Enabled = False
    End Sub

    Private Sub txtFltDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFltDate.LostFocus
        If InStr(NonAirList, Me.CmbDocType.Text) = 0 AndAlso CDate(Me.txtFltDate.Text) < Today.Date Then Me.txtFltDate.Focus()
        If WhatAction = "NON" Then Me.CmdSaveChange.Enabled = True
    End Sub

    Private Sub CmbProduct_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbProduct.LostFocus
        Dim HH As Decimal = 0
        Me.txtNetToAL.Enabled = False
        Me.txtALCommVAL.Enabled = False
        Me.txtALCommPCT.Enabled = False

        Me.txtNetToAL.TabStop = False
        Me.txtALCommVAL.TabStop = False
        Me.txtALCommPCT.TabStop = False
        If Me.CmbProduct.Text.Trim = "FOC" Then
            Me.txtALCommPCT.Text = Format(0, "#,##0.00")
            Me.txtALCommVAL.Text = Format(0, "#,##0.00")
            Me.txtNetToAL.Text = Format(0, "#,##0.00")
            Exit Sub
        ElseIf Me.CmbProduct.Text.Trim = "NET" Or Me.CmbProduct.Text.Trim = "EXC" Then
            Me.txtNetToAL.Enabled = True
            Me.txtNetToAL.TabStop = True
            Me.txtNetToAL.Text = 0
            Me.txtALCommPCT.Text = 0
            Me.txtALCommVAL.Text = 0
            Me.txtNetToAL.Focus()
            Exit Sub
        ElseIf Me.CmbProduct.Text.Trim = "S+M" Then 'Selling  co MU nen sua ca net2AL va commVAL 
            Me.txtNetToAL.Enabled = True
            Me.txtNetToAL.TabStop = True
            Me.txtALCommVAL.Enabled = True
            Me.txtALCommVAL.TabStop = True
            Me.txtNetToAL.Text = 0
            Me.txtALCommPCT.Text = 0
            Me.txtALCommVAL.Text = 0
            Me.txtNetToAL.Focus()
            Exit Sub
        ElseIf Me.CmbProduct.Text.Trim = "???" Then 'Hoac gia thuong (vd FIT) nhung ko the tinh trc dc HH, vdu EK 
            Me.txtALCommPCT.Enabled = True
            Me.txtALCommPCT.TabStop = True
            Me.txtALCommVAL.Enabled = True
            Me.txtALCommVAL.TabStop = True
            Me.txtNetToAL.Text = 0
            Me.txtALCommPCT.Text = 0
            Me.txtALCommVAL.Text = 0
            Me.txtALCommVAL.Focus()
            Exit Sub
        Else
            ReCalcALComm()
        End If
    End Sub

    Private Sub txtTax1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtTax.LostFocus, _
        txtTax1.LostFocus, TxtTax2.LostFocus, TxtTax3.LostFocus, TxtTax4.LostFocus, TxtTax5.LostFocus
        Dim txt As TextBox = CType(sender, TextBox)
        Dim taxText As String
        Dim aa As Decimal = CDec(txt.Text)
        txt.Text = Format(aa, "#,##0.00")
        If aa < 0 Then txt.Focus()
        If txt.Name = "TxtTax" Then
            Me.CmdAddTKT.Enabled = True
            Me.CmdSaveChange.Enabled = True
        Else
            If aa > CDec(Me.TxtTax.Text) Then
                For c As Int16 = 1 To 5
                    taxText = "TxtTax" & Trim(Str(c))
                    Me.GrpTax.Controls(taxText).Text = 0
                    Me.TxtTax.Focus()
                Next
            End If
        End If
    End Sub

    Private Sub CmbCommOffer_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCommOffer.LostFocus
        Dim PCT As Decimal = 0, VAL As Decimal, Fare As Decimal = CDec(Me.txtFare.Text)
        Me.TxtAgtDiscPCT.Text = 0
        Me.TxtAgtDiscVal.Text = 0
        If Me.CmbCommOffer.Text.Trim = "" Then Exit Sub
        Dim PQT As String() = CmbCommOffer.SelectedValue.ToString.Split("|")
        PCT = IIf(PQT(0) = "PCT", CDec(PQT(1)), 0)
        VAL = IIf(PQT(0) = "PCT", 0, CDec(PQT(1)))
        If PCT > 0 Then
            Me.TxtAgtDiscPCT.Text = Format(PCT, "#,##0.00")
            Me.TxtAgtDiscVal.Text = Format(PCT * Fare / 100, "#,##0.00")
        Else
            Me.TxtAgtDiscVal.Text = Format(VAL, "#,##0.00")
            Me.TxtAgtDiscPCT.Text = Format(VAL / Fare * 100, "#,##0.00")
        End If
        If Me.CmbCommOffer.Text.Substring(0, 1) = "*" Then
            Me.TxtAgtDiscPCT.Enabled = False
            Me.TxtAgtDiscVal.Enabled = False
        Else
            Me.TxtAgtDiscPCT.Enabled = True
            Me.TxtAgtDiscVal.Enabled = True
            Me.TxtAgtDiscPCT.Focus()
        End If
    End Sub
    Private Sub CmbChannel_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbChannel.LostFocus
        If Me.CmbChannel.Items.Count < 1 Then Exit Sub
        If Me.CmbChannel.SelectedValue Is Nothing Then
            Me.CmbCommOffer.Items.Clear()
            Exit Sub
        End If
        If MySession.Domain = "TVS" And MySession.Counter = "TVS" Then
            Me.CmbProduct.Text = "NET"
            Me.CmbProduct.Enabled = False
            Me.txtNetToAL.Enabled = True
            Me.TxtAgtDiscPCT.Text = 0
            Me.TxtAgtDiscVal.Text = 0
        Else
            Me.CmbProduct.Enabled = True
        End If
        If (InStr("CS_LC", Me.CmbCustType.Text) > 0 Or Me.CmbChannel.SelectedValue < 0) And _
            mRCP = "" And pubVarSRV = "S" Then
        End If
        Me.GrpComm.Enabled = GenCOC(Me.CmbChannel.SelectedValue, Me.CmbAL.Text)
        Me.CmbCurr.Focus()
    End Sub
    Private Sub GenChargeAndFee()
        Dim strSQL As String = "select Amount as VAL, Type as DIS from Charge_comm where cat='CHGE' and currency in ('ALL','" & _
            Me.CmbCurr.Text & "') and '" & Now.Date & "' between ValidFrom and ValidThru and " & _
            " AL in ('YY','" & Me.CmbAL.Text & "') and status='OK' and trx in ('A','S')" & _
            " order by substring(type,3,1) + left(type,2) , substring(type,4,6)"
        LoadCmb_VAL(Me.CmbCharge1, strSQL)
        LoadCmb_VAL(Me.CmbCharge2, strSQL)
        LoadCmb_VAL(Me.CmbCharge3, strSQL)
        LoadCmb_VAL(Me.CmbCharge4, strSQL)
    End Sub
    Private Function GenCOC(ByVal pCustID As Integer, ByVal pAL As String) As Boolean
        If MySession.Domain = "TVS" And InStr(myStaff.CAccess, "CS") + InStr(myStaff.CAccess, "LC") = 0 Then
            Dim Cmd As SqlClient.SqlCommand = Conn.CreateCommand
            Dim IsDiscountable As Boolean
            Cmd.CommandText = "select EnableDiscount from airline where al='" & pAL & "'"
            IsDiscountable = Cmd.ExecuteScalar
            If IsDiscountable = False Then
                LoadCmb_VAL(Me.CmbCommOffer, "select '' as DIS, '0|0|' as VAL")
                Return False
            End If
        End If
        Dim strSQL As String, DK As String = " Status='OK' and '" & Now.Date & "' between "
        strSQL = "select COC as DIS, DType  + '|' + cast(VAL as nvarchar(12))  as VAL from Cust_Discount "
        strSQL = strSQL & " where " & DK & " ValidFrom and ValidThru and SBU+AL+Channel+CustLevel in ("
        strSQL = strSQL & " select sbu+al+channel+CustLevel from cust_channel_level where " & DK
        strSQL = strSQL & " ValidFrom and ValidThru and custID= " & pCustID & " and SBU='" & MySession.Domain & "'"
        strSQL = strSQL & " and al='" & pAL & "')"
        strSQL = strSQL & " union select '' as DIS,  '0|0|' as VAL "
        LoadCmb_VAL(Me.CmbCommOffer, strSQL)
        Return True
    End Function
    Private Sub CmbChannel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbChannel.SelectedIndexChanged
        If Me.CmbChannel.Items.Count < 1 Then Exit Sub
        Dim strDK As String, CMC As String
        Try
            MyCust.CustID = Me.CmbChannel.SelectedValue
            Me.txtCustName.Text = MyCust.FullName
            Me.txtCustAddrr.Text = MyCust.Addr
            Me.txtTaxCode.Text = MyCust.taxCode
            Me.CmbTayBa.Items.Clear()
            If MyCust.TayBa.Length > 2 Then
                For i As Int16 = 0 To UBound(MyCust.TayBa.Split("_"))
                    If MyCust.TayBa.Split("_")(i).Substring(0, 2) = MyAL.ALCode Then
                        Me.CmbTayBa.Items.Add(MyCust.TayBa.Split("_")(i).Substring(2))
                    End If
                Next
            End If
            If Me.CmbTayBa.Items.Count > 0 Then
                Me.CmbTayBa.Items.Add("")
                Me.CmbTayBa.Text = ""
                Me.CmbTayBa.Visible = True
            Else
                Me.CmbTayBa.Visible = False
            End If
            Me.LblCA.Visible = Me.CmbTayBa.Visible
            If InStr("CS_LC", MyCust.CustType) > 0 Then
                LoadCmbBooker(MyCust.CustID)
                Me.txtShownFare.Enabled = True
                Me.txtLLF.Enabled = True

                CMC = ScalarToString("cwt.dbo.go_companyInfo1", "top 1 CMC", "CustID=" & MyCust.CustID & " and status='OK'")
                strDK = "select distinct VAL + '-' + details as VAL from cwt.dbo.go_MISC where cat='RSavingHtl' and RMK='"
                LoadCmb_MSC(Me.cmbRSaving_HTL, strDK & CMC & "'")
                If Me.cmbRSaving_HTL.Items.Count = 0 Then
                    LoadCmb_MSC(Me.cmbRSaving_HTL, strDK & "'")
                End If

                strDK = "select distinct VAL + '-' + Details as VAL from cwt.dbo.go_MISC where cat='MSavingHtl' and RMK='"
                LoadCmb_MSC(Me.CmbMSaving_HTL, strDK & CMC & "'")
                If Me.CmbMSaving_HTL.Items.Count = 0 Then
                    LoadCmb_MSC(Me.CmbMSaving_HTL, strDK & "'")
                End If
                Me.GrpComm.Visible = False
                Me.GrpReasonCode.Visible = True
            ElseIf MySession.Domain = "GSA" And InStr("GA_PG", MyAL.ALCode) > 0 Then
                Me.Label6.Text = "ShownFare"
                Me.txtShownFare.Enabled = True
                Me.GrpComm.Visible = True
                Me.GrpReasonCode.Visible = False
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub LoadCmbBooker(pCustID As Integer)
        Try
            LoadCmb_MSC(Me.CmbBooker, "Select BookerName as VAL from [42.117.5.86].ft.dbo.Cwt_Bookers where Status='OK' and CustId=" & pCustID)
        Catch ex As Exception
            LoadCmb_MSC(Me.CmbBooker, "select distinct fValue as VAL from SIR where fName+status='BOOKEROK' and custID=" & pCustID)
        End Try

    End Sub
    Private Sub CmdSaveChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSaveChange.Click
        Dim RowNo As Int16, MyAnswer As Int16
        RowNo = Me.GridTKT.CurrentCell.RowIndex
        If Not LogicCheckOK() Then Exit Sub
        Me.CmdSaveChange.Enabled = False
        If pubVarSRV = "R" Then
            If Me.GridTKT.Item("FTKT", RowNo).Value.ToString.Trim.Length = 7 Then
                FeedBack("Refund this Conjunction Ticket Will Refund All the Rest")
                On Error Resume Next
                For i As Int16 = RowNo + 1 To RowNo + 3
                    Me.GridTKT.Item("TKNO", i).Style.Font = New Font(Me.GridTKT.DefaultFont, FontStyle.Strikeout)
                    If Me.GridTKT.Item("FTKT", i).Value.ToString.Trim.Length = 3 Then Exit For
                Next
                On Error GoTo 0
            End If
            ' Alert khi refund ve co shopping voucher
            If Me.GridTKT.Item("Promocode", RowNo).Value.ToString.Length > 8 AndAlso _
                Me.GridTKT.Item("Promocode", RowNo).Value.ToString.Substring(2, 2) = "SV" Then
                MyAnswer = MsgBox("Shopping Voucher Was Given Upon TKT Issuance. Deduct That Amount?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
                If MyAnswer = vbYes Then
                    Dim SVamount As Decimal
                    On Error GoTo errHandler
                    SVamount = CDec(Me.GridTKT.Item("Promocode", RowNo).Value.ToString.Substring(4, 3))
                    SVamount = SVamount * 1000 / CDec(Me.LcktxtROE.Text)
                    Me.txtDiscountedAmount.Text = Format(CDec(Me.txtDiscountedAmount.Text) + SVamount, "#,##0.00")
                    Me.txtRMK.Text = Me.txtRMK.Text & "|RFSV"
                    On Error GoTo 0
                Else
                    MsgBox("Remember to Collect Shopping Voucher!", MsgBoxStyle.Critical, msgTitle)
                End If
            End If
            ' tru CC fee neu truoc do thanh toan bang CC
            If amtCRDPaid > 0 Then
                Dim tktAmt As Decimal
                tktAmt = CDec(Me.GridTKT.Item("Fare", RowNo).Value) + CDec(Me.GridTKT.Item("Tax", RowNo).Value) - _
                    CDec(Me.GridTKT.Item("AgtDisctVAL", RowNo).Value)
                If amtCRDPaid > tktAmt Then amtCRDPaid = tktAmt
                amtCRDPaid = amtCRDPaid * 0.03
                Me.CmbTVCharge1.Text = "CRD"
                Me.txtTVCharge1.Text = Format(CDec(Me.txtDiscountedAmount.Text) + amtCRDPaid, "#,##0.00")
            End If
        End If
        addTKTtoGridTKT("CHANGE")
        Exit Sub
errHandler:
        MsgBox("Invalid PromoCode. Cant Define The Amount. Pls Calculate Manually", MsgBoxStyle.Critical, msgTitle)
    End Sub
    Private Sub PadClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PadClose.Click
        QuitNow = True
        ReleaseRCP(pubVarRCPID_BeingCreated)
        Me.Close()
    End Sub
    Private Sub ResetMenu()
        Me.BarRemoveTKT.Enabled = False
        Me.BarCopyTKT.Enabled = False
        Me.BarEditTax.Enabled = False
    End Sub
    Private Sub RemoveOrCopy(ByVal VarRemOrCpy As String)
        Dim rRow As Int16, SoVe As String, soVeNoi As String, Vei As Int16
        Dim VeDau As Int16, VeCuoi As Int16, tmpVeCuoi As String
        Me.BarRemoveTKT.Enabled = False
        Me.BarCopyTKT.Enabled = Me.BarRemoveTKT.Enabled
        rRow = Me.GridTKT.CurrentCell.RowIndex
        SoVe = Me.GridTKT.Item("TKNO", rRow).Value
        soVeNoi = Me.GridTKT.Item("FTKT", rRow).Value
        If soVeNoi = "" Then
            VeDau = rRow
            VeCuoi = rRow
        Else
            Vei = rRow
            Do
                If Me.GridTKT.Item("FTKT", Vei).Value.ToString.Substring(0, 3) = "___" Then Exit Do
                Vei = Vei - 1
            Loop
            VeDau = Vei
            Vei = rRow
            Do While Vei < Me.GridTKT.Rows.Count
                If Me.GridTKT.Item("FTKT", Vei).Value.ToString.Trim.Length = 3 Then
                    If CDec(Strings.Right(Me.GridTKT.Item("TKNO", Vei).Value, 3)) = CDec(Me.GridTKT.Item("FTKT", Vei).Value.ToString.Trim) + 1 Then Exit Do
                End If
                Vei = Vei + 1
            Loop
            VeCuoi = Vei
        End If
        If VarRemOrCpy = "R" Then
            For i As Int16 = VeCuoi To VeDau Step -1
                Me.GridTKT.Rows.RemoveAt(i)
            Next
        Else
            Dim TCP As Int16
            On Error GoTo InvalidTCP
            TCP = InputBox("How Many Copies Do You Want to Make?", msgTitle, 15)
            On Error GoTo 0
            If VeDau <> VeCuoi And rRow < Me.GridTKT.RowCount - 1 Then
                For i As Int16 = rRow To Me.GridTKT.RowCount - 1
                    If Me.GridTKT.Item("FTKT", i).Value = "" Then
                        FeedBack("Must Copy From the Last Ticket of the Range!  ")
                        Exit Sub
                    End If
                Next
            End If
            For i As Int16 = 1 To TCP
                For r As Int16 = VeDau To VeCuoi
                    CloneTKT(r, i, VeCuoi - VeDau + 1)
                Next
            Next
        End If
        If GridTKT.RowCount > 1 Then
            tmpVeCuoi = Me.GridTKT.Item("TKNO", Me.GridTKT.RowCount - 1).Value
            NextTKNO = DefineNextTKNO(tmpVeCuoi, MyAL.isTKTLess)
        End If
        Exit Sub
InvalidTCP:
        FeedBack("Error. Invalid Number Input!  ")
        Exit Sub
    End Sub
    Private Sub pStep2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles pStep2.Enter
        ResetMenu()
        If pubVarSRV = "S" Then Me.BarEditTax.Enabled = True
        If Me.CmbAL.Text = "" Then
            Me.TabControl1.SelectTab("pStep1")
            FeedBack("Error. You Have to Specify an Airline")
        End If
        If Me.txtTKNO.Text = "" Then Me.txtTKNO.Text = MyAL.DocCode
        GenChargeAndFee()
        Me.LcktxtROE.Enabled = False
    End Sub
    Private Sub CloneTKT(ByVal Srow As Int16, ByVal VeI As Int16, ByVal MayVe As Int16)
        Dim Drow As Int16, j As Int16, DaCoTrongGrid As Boolean = False, AA As Int16
        Dim SoVeGoc As Int32, veNoiGoc As String, veNoiMoi As String = "", tmpStr As String
        Dim tmp As String
        If Me.GridTKT.Item(0, Srow).Value.ToString.Trim.Substring(0, 1) = "Z" Then 'Ticketless
            j = 3
            tmpStr = "000"
        Else
            j = 6
            tmpStr = "000000"
        End If
        AA = Me.GridTKT.Item(0, Srow).Value.ToString.Trim.Length - j
        NextTKNO = Me.GridTKT.Item(0, Srow).Value.ToString.Trim.Substring(0, AA)
        SoVeGoc = CDec(Me.GridTKT.Item(0, Srow).Value.ToString.Trim.Substring(AA, j))
        NextTKNO = NextTKNO & Format(SoVeGoc + VeI * MayVe, tmpStr)
        DaCoTrongGrid = ChangeColorGridTKT(NextTKNO)
        If DaCoTrongGrid Then
            FeedBack(NextTKNO & " Has Been Updated and Action Was Aborted! ")

            Exit Sub
        End If
        Me.GridTKT.Rows.Add()
        Drow = Me.GridTKT.Rows.Count - 1
        If MayVe > 1 Then
            veNoiGoc = Me.GridTKT.Item("FTKT", Srow).Value.ToString.Trim
            j = IIf(veNoiGoc.Length > 3, 2, 1)
            veNoiMoi = ""
            For i As Int16 = 1 To j
                tmpStr = veNoiGoc.Substring((i - 1) * 4, 3)
                If tmpStr = "___" Then
                    veNoiMoi = tmpStr
                ElseIf i = 1 Then
                    tmp = Format(CInt(tmpStr) + VeI * MayVe, "000")
                    veNoiMoi = tmp.Substring(tmp.Length - 3, 3)
                Else
                    tmp = Format(CInt(tmpStr) + VeI * MayVe, "000")
                    tmp = tmp.Substring(tmp.Length - 3, 3)
                    veNoiMoi = veNoiMoi & "/" & tmp
                End If
            Next
            Me.GridTKT.Item("FTKT", Drow).Value = veNoiMoi
        End If
        Me.GridTKT.Item(0, Drow).Value = NextTKNO
        If MayVe > 1 Then
            If veNoiMoi.Substring(0, 1) <> "_" Then
                Me.GridTKT.Item(0, Drow).Style.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
        End If
        For i As Int16 = 2 To Me.GridTKT.ColumnCount - 1
            Me.GridTKT.Item(i, Drow).Value = Me.GridTKT.Item(i, Srow).Value
        Next
    End Sub
    Private Function ChangeColorGridTKT(ByVal varTKNO As String) As Boolean
        ChangeColorGridTKT = False
        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            If Me.GridTKT.Item(0, i).Value.ToString.Trim = varTKNO Then
                ChangeColorGridTKT = True
                Exit Function
            End If
            If Me.GridTKT.Item(0, i).Style.Alignment = DataGridViewContentAlignment.MiddleRight Then
                Me.GridTKT.Item(0, i).Style.ForeColor = Color.Gray
                Me.GridTKT.Item(1, i).Style.ForeColor = Color.Gray
            End If
        Next
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.StatusFeedBack.Text = ""
        Me.Timer1.Enabled = False
    End Sub
    Private Sub CheckReservation(pTRXCode As String, pKyHieu As String)
        If myStaff.City <> "SGN" Then Exit Sub
        Dim LstRun As String = ScalarToDate("ActionLog", "F1", "TableName='RESERVATION' and DoWhat='" & pTRXCode & "'")
        If CDate(LstRun).Date = Now.Date Then Exit Sub
        If pTRXCode <> "TS" And Now.Day Mod 5 <> 0 Then Exit Sub
        LstRun = Now.Date.ToString
        Dim INVNo As String = GenInvNo_QD153(pTRXCode & "??" & (Now.Year - 2000).ToString.Trim, pKyHieu)
        Dim InvID As Integer = Insert_INV("E", INVNo, INVNo.Substring(0, 2), -64)
        Dim dtable As DataTable = GetDataTable("Select distinct PaxName from TKT")
        Randomize()
        Dim i As Integer = Rnd() * dtable.Rows.Count
        Dim anyName As String = dtable.Rows(i)("PaxName")
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update ActionLog set F1='" & LstRun & "' where TableName='RESERVATION' and DoWhat='" & pTRXCode & "';" & _
            "update INV set srv='V', custID=697, CustShortName='TEST', Amount=0, CustFullName='" & anyName & "' where recID=" & InvID
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSave.Click
        If SysDateIsWrong(Conn) Then Exit Sub
        If HasNewerVersion_R12(Application.ProductVersion) Then End
        If myStaff.SICode <> "SYS" And (InStr("XXX_ALL", myStaff.Counter) > 0 Or myStaff.Counter = "") Then Exit Sub
        CheckReservation(MySession.TRXCode, MyAL.VAT_KyHieu)
        Dim myAns As Int16, tmp As String
        If MyCust.CustID = 0 Then
            MsgBox("No Customer Has Been Selected. Action Aborted", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If Me.GridTKT.Item("DocType", 0).Value = "GRP" Then
            If pubVarSRV = "R" Then ' check xem co RQ refund deposit ko? refund hay phat nhu nhau tuy amount
                If Not CoRQDepositHandle(Me.txtValueToSearch.Text, "OK-R") Then Exit Sub
            ElseIf pubVarSRV = "S" Then ' check xem co dung exc hay ko? dung tuc la cho doi dat coc thi phai check RQ
                For myAns = 0 To Me.GridFOP.RowCount - 2
                    If Me.GridFOP.Item(0, myAns).Value = "EXC" AndAlso Me.GridFOP.Item(4, myAns).Value.ToString.Substring(0, 3) = "GRP" Then
                        If Not CoRQDepositHandle(Me.GridFOP.Item(4, myAns).Value, "OK-E") Then Exit Sub
                    End If
                Next
            End If
        End If

        If Me.txtVNDEquivalent.Text = 0 Then
            myAns = MsgBox("This Transaction Has No Value. Are You Sure?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
            If myAns = vbNo Then
                Exit Sub
            End If
        End If
        If pubVarSRV <> "V" And CDec(Me.txtNettPayable.Text) <> 0 Then
            If (Me.CmbCurr.Text = "USD" And CDec(Me.txtNettPayable.Text) > 32000) Or _
                (Me.CmbCurr.Text = "VND" And CDec(Me.txtNettPayable.Text) < 32000) Then
                myAns = MsgBox("Illogic Currency/Amount. Plz Check. Wanna Recheck Input?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
                If myAns = vbYes Then Exit Sub
            End If
        End If

        tmp = Trung_EXC_PTA_MCO()
        If tmp <> "" Then
            FeedBack(tmp)
            Exit Sub
        End If
        If CoPSP("PSP") And pubVarSRV = "S" Then
            If MyCust.CrditOnAL <> "YY" AndAlso InStr(MyCust.CrditOnAL, MySession.TRXCode) = 0 Then
                FeedBack("This Customer Doesn't Have Credit Agreement on " & MySession.TRXCode)
                Exit Sub
            End If
        End If
        If InStr("CS_LC", Me.CmbCustType.Text) > 0 Then
            If CoPSP("ANY") And Me.ChkBizTrip.Checked = False Then
                FeedBack("Err: Plz Dont Accept PSP/PPD for Personal Trip")
                Exit Sub
            ElseIf CoPSP("DEB") And Me.ChkBizTrip.Checked And InStr(Me.GridTKT.Item("TKNO", 0).Value.ToString, " TV") = 0 Then
                FeedBack("Err: Plz Dont Accept DEB for Business Trip")
                Exit Sub
            End If
        End If

        If Not CRD_format_Correct() Then
            FeedBack("Err. Invalid CRD No")
            Exit Sub
        End If

        DaTra = 0
        For i As Int16 = 0 To Me.GridFOP.RowCount - 1
            If CDec(Me.GridFOP.Item("Amount", i).Value) <> 0 Or Me.GridFOP.Item("Document", i).Value <> "" Then
                DaTra = DaTra + Me.GridFOP.Item("RCPROE", i).Value * CDec(Me.GridFOP.Item("Amount", i).Value)
            End If
        Next
        If Math.Round(DaTra, 2) <> CDec(Me.txtVNDEquivalent.Text) Then
            FeedBack("Error. Invalid Total Amoutn Paid")
            Exit Sub
        End If
        If Me.TxtTRXNO.Text.Substring(0, 2) = "YY" Then Exit Sub
        Me.CmdSave.Enabled = False
        If SaveInput("OK") Then
            If CoPSP("ANY") Then RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, False, "After " & Me.TxtTRXNO.Text, Conn, myStaff.SICode, CnStr)
            Me.CmdPreview.Enabled = True
            Me.CmbChannel.Enabled = False
            Me.CmbCustType.Enabled = False
        End If
    End Sub
    Private Function CoPSP(ByVal pDelayType As String) As Boolean
        For i As Int16 = 0 To Me.GridFOP.RowCount - 1
            If Me.GridFOP.Item("FOP_FOP", i).Value = pDelayType Or _
                (pDelayType = "ANY" And InStr("PSP_PPD", Me.GridFOP.Item("FOP_FOP", i).Value) > 0) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function Trung_EXC_PTA_MCO() As String
        Dim DocList As String = "", KQ As String = ""
        For i As Int16 = 0 To Me.GridFOP.RowCount - 1
            If InStr("MCO_PTA_EXC_DEP", Me.GridFOP.Item("FOP_FOP", i).Value) > 0 Then
                If InStr(DocList, Me.GridFOP.Item("Document", i).Value) = 0 Then
                    DocList = DocList & "_" & Me.GridFOP.Item("Document", i).Value
                Else
                    Return "Duplicate Document Numbers"
                End If
            End If
        Next
        Return KQ
    End Function
    Private Function XD_DEB_PPD_PSP_Amt() As String
        Dim FOPList As String = "PPD_PSP_DDB_DEB_ITP_", j As Int16 = 0
        Dim VND_Amt(8) As Decimal, soFOP As Int16 = FOPList.Length / 4
        Dim tmpAmt As Decimal, PPDOnly As Decimal
        For r As Int16 = 0 To Me.GridFOP.RowCount - 1
            If CDec(Me.GridFOP.Item("Amount", r).Value) <> 0 Then
                tmpAmt = Me.GridFOP.Item("Amount", r).Value * Me.GridFOP.Item("RCPROE", r).Value
                For i As Int16 = 1 To soFOP
                    If Me.GridFOP.Item("FOP_FOP", r).Value = FOPList.Substring(4 * (i - 1), 3) Then
                        VND_Amt(i) = VND_Amt(i) + tmpAmt
                        If j = 0 Then j = i
                    End If
                Next
                If Me.GridFOP.Item("FOP_FOP", r).Value = "PPD" Then
                    PPDOnly = PPDOnly + tmpAmt
                End If
            End If
        Next
        If j = 0 Then Return "OK00"
        tmpAmt = 0
        For i As Int16 = 1 To soFOP
            If i <> j And VND_Amt(i) > 0 Then
                Return "MIX" & PPDOnly.ToString
            End If
        Next
        Return FOPList.Substring(4 * (j - 1), 3) & VND_Amt(j).ToString
    End Function
    Private Function CRD_format_Correct() As Boolean

        If InStr("UA_NH", Me.CmbAL.Text) = 0 Or MySession.Domain <> "GSA" Then Return True
        For i As Int16 = 0 To Me.GridFOP.RowCount - 1
            If InStr("UCF_CRD", Me.GridFOP.Item("FOP_FOP", i).Value) > 0 AndAlso _
                    (InStr("ACDIJMTV", Me.GridFOP.Item("Document", i).Value.ToString.Substring(0, 1)) = 0 OrElse _
                    InStr(Me.GridFOP.Item("Document", i).Value.ToString.ToUpper, "XXXXX") = 0 OrElse _
                    CDec(Me.GridFOP.Item("Document", i).Value.ToString.Substring(2, 3)) = 0 OrElse _
                    CDec(Me.GridFOP.Item("Document", i).Value.ToString.Substring(13, 4)) = 0) Then
                If CDec(Me.GridFOP.Item("Document", i).Value.ToString.Substring(13, 4)) = 0 And _
                    Me.GridFOP.Item("Document", i).Value.ToString.Substring(13, 4) = "0000" Then
                Else
                    MsgBox("Invalid Document Number." & vbCrLf & " - 1st Character Must Be One of A D J M T V" & vbCrLf & " - It Must Contain 4 1st and 4 Last Digit of Card Number" & vbCrLf & " - The Rest Must Be XXXX..", MsgBoxStyle.Critical, msgTitle)
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    Private Function NoOK() As Boolean
        Dim VND_DelayPmt As String, VND_DelayAmt As Decimal, VND_Avail As Decimal = 0
        Dim tmpCutOverDate As Date, DelayedType As String, tmpTKNo As String
        Dim AmtApproved As Decimal, CustID As Integer = Me.CmbChannel.SelectedValue
        Dim tmpOverDue As String = "", HoldCredit As Decimal

        VND_DelayPmt = XD_DEB_PPD_PSP_Amt()
        DelayedType = VND_DelayPmt.Substring(0, 3)
        VND_DelayAmt = CDec(VND_DelayPmt.Substring(3))

        If InStr("CRO", pubVarSRV) > 0 AndAlso DelayedType = "ITP" Then Return True ' xu ly ngoai le MU cho phep ITP de ve NKKN hoan tien

        If DelayedType = "OK0" Then Return True
        If DelayedType = "MIX" Then
            If InStr("CS_LC", Me.CmbCustType.Text) > 0 And MyCust.DelayType = "PPD" And MyCust.AdhType = "PSP" Then
                DelayedType = "PPD"
            Else
                MsgBox("Mixed Delayed Payment Not Allowed", MsgBoxStyle.Critical, msgTitle)
                Return False
            End If
        End If

        If InStr("CS_LC", MyCust.CustType) > 0 Then
            If Me.ChkBizTrip.Checked Then
                If MyCust.DelayType <> DelayedType AndAlso DelayedType <> MyCust.AdhType Then
                    MsgBox("Customer Type and Delayed Payment Type Mismatch", MsgBoxStyle.Critical, msgTitle)
                    Return False
                End If
            Else
                If InStr("PPD_PSP", DelayedType) > 0 Then
                    MsgBox("Cant Accept Inv FOP for Personal Trips", MsgBoxStyle.Critical, msgTitle)
                    Return False
                End If
            End If
        Else
            If MyCust.DelayType <> DelayedType And DelayedType <> "ITP" Then
                MsgBox("Customer Type and Delayed Payment Type Mismatch", MsgBoxStyle.Critical, msgTitle)
                Return False
            End If
        End If

        If InStr("CRO", pubVarSRV) > 0 Then Return True

        If InStr("PSP_PPD", DelayedType) > 0 Then
            tmpCutOverDate = CutOverDatePPD
            If DelayedType = "PSP" Then
                tmpCutOverDate = CutOverDatePSP
                If MyCust.DelayType <> "PSP" And MyCust.AdhType <> "PSP" Then
                    MsgBox("No PostPaid Agreement With This Customer", MsgBoxStyle.Critical, msgTitle)
                    Return False
                End If
                If InStr("CS_LC", MyCust.CustType) = 0 Then tmpOverDue = CheckOverDue(Me.CmbChannel.SelectedValue, Conn)
            End If
            VND_Avail = defineVND_Avail(CustID, DelayedType, tmpCutOverDate, MyCust.LstReconcile, "B4 " & Me.TxtTRXNO.Text, Conn, myStaff.SICode, CnStr)
            If VND_DelayAmt > VND_Avail Then HoldCredit = XacDinhAmountHold()
        ElseIf DelayedType = "DEB" Then
            VND_Avail = DefineDEB_Avail(CustID)
        End If

        If (VND_DelayAmt > VND_Avail And HoldCredit = 0) Or tmpOverDue <> "" Or InStr("DEB_ITP", DelayedType) > 0 Then
            tmpTKNo = Me.GridTKT.Item("TKNO", 0).Value.ToString.Replace(" ", "")
            tmpTKNo = tmpTKNo.Substring(3)
            If DelayedType = MyCust.AdhType And InStr("CS_LC", MyCust.CustType) > 0 Then
                Return True
            Else
                AmtApproved = GetApproval(Me.TxtTRXNO.Text, DelayedType, CustID, tmpTKNo)
                If VND_DelayAmt > AmtApproved Then
                    Me.BarRQ4CRDExt.Enabled = True
                    MsgBox("No or Over Credit Line" & IIf(tmpOverDue <> "", " or Has OverDue Amt", "") & ". Need Supervisor Approval", MsgBoxStyle.Critical, msgTitle)
                    Return False
                End If
            End If
        End If
        Return True
    End Function
    Private Function DefineDEB_Avail(ByVal pCustID As Integer) As Decimal
        Dim Cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim DEBQuota As Decimal, PendingDEB As Decimal
        Dim strDK1 As String = " fop+Status='DEBOK' and RCPid in "
        Dim StrDK2 As String = " custid=" & pCustID & ")"
        DEBQuota = ScalarToDec("Cust_detail", "VAL", "cat+Status='DEBOK' and custid=" & pCustID)

        'PendingDEB = ScalarToDec("FOP_TVS", " isnull(sum(roe*Amount),0) ", strDK1 & "(select recid from rcp_TVS where" & StrDK2)
        'PendingDEB = PendingDEB + ScalarToDec("FOP_GSA", " isnull(sum(roe*Amount),0) ", strDK1 & "(select recid from rcp_GSA where" & StrDK2)

        Return DEBQuota - PendingDEB
    End Function

    Private Function SaveInput(ByVal varStatus As String) As Boolean
        Dim ROEID As Integer, strSQL As String = "", tmpRMK As String
        Dim LclVarSRV As String = pubVarSRV, TKTList As String, tmpDocNo As String = "", RMK As String = ""
        Me.BarRQ4CRDExt.Enabled = False

        If LclVarSRV = "C" Or LclVarSRV = "O" Or Me.GridTKT.Item("DocType", 0).Value = "ACM" Then LclVarSRV = "R"
        If pubVarSRV = "I" Then LclVarSRV = "S"
        RCPNOtoPrint = Me.TxtTRXNO.Text

        If TKTAlreadyExist(pubVarRCPID_BeingCreated, pubVarSRV) Then Return False

        If Not NoOK() Then Return False ' Doan Nay kiem tra xem PPD va DEB co valid ko

        If DefineBSPStock() Then
            ROEID = ForEX_12(Me.TxtDOS.Value, "USD", "RECID", "IATA")
        Else
            ROEID = ForEX_12(Me.TxtDOS.Value, "USD", "RECID", MySession.TRXCode)
        End If

        tmpRMK = Me.txtRMK.Text & ExcDocDetail
        If Me.CmbTC.Text <> "" Then tmpRMK = tmpRMK & "|TC" & Me.CmbTC.Text
        tmpRMK = tmpRMK.Replace("'", "")

        Dim t As SqlClient.SqlTransaction = Conn.BeginTransaction
        cmd.Transaction = t
        Try
            cmd.CommandText = "update RCP set Custid=@Custid, CustShortName=@CustShortName, CustType=@CustType, " & _
                "City=@City, Location=@Location, Counter=@Counter, Status=@Status, Currency=@Currency, ROE=@ROE," & _
                "TTLDue=@TTLDue, Discount=@Discount, Charge=@Charge, PrintedCustName=@PrintedCustName, PrintedCustAddrr=@PrintedCustAddrr, " & _
                "PrintedTaxCode=@PrintedTaxCode, CA=@CA, Charge_D=@Charge_D, Discount_D=@Discount_D, " & _
                "SRV=@SRV, DOS=@DOS, Stock=@Stock, ROEID=@ROEID, RMK=@RMK where RecID=@RecID"
            cmd.Parameters.Clear()
            cmd.Parameters.Add("@CustID", SqlDbType.Int).Value = MyCust.CustID
            cmd.Parameters.Add("@CustShortName", SqlDbType.VarChar).Value = MyCust.ShortName
            cmd.Parameters.Add("@CustType", SqlDbType.VarChar).Value = MyCust.CustType
            cmd.Parameters.Add("@City", SqlDbType.VarChar).Value = MySession.City
            cmd.Parameters.Add("@Location", SqlDbType.VarChar).Value = MySession.Location
            cmd.Parameters.Add("@Counter", SqlDbType.VarChar).Value = MySession.Counter
            cmd.Parameters.Add("@Status", SqlDbType.VarChar).Value = varStatus
            cmd.Parameters.Add("@Currency", SqlDbType.VarChar).Value = Me.CmbCurr.Text

            cmd.Parameters.Add("@ROE", SqlDbType.Decimal).Value = CDec(Me.LcktxtROE.Text)
            cmd.Parameters.Add("@TTLDue", SqlDbType.Decimal).Value = CDec(Me.txtNettPayable.Text)
            cmd.Parameters.Add("@Discount", SqlDbType.Decimal).Value = CDec(Me.txtTVDiscount1.Text) + CDec(Me.txtTVDiscount2.Text)
            cmd.Parameters.Add("@Charge", SqlDbType.Decimal).Value = CDec(Me.txtTVCharge1.Text) + CDec(Me.txtTVCharge2.Text) + CDec(Me.txtDiscountedAmount.Text)
            cmd.Parameters.Add("@PrintedCustName", SqlDbType.NVarChar).Value = Me.txtCustName.Text
            cmd.Parameters.Add("@PrintedCustAddrr", SqlDbType.NVarChar).Value = Me.txtCustAddrr.Text
            cmd.Parameters.Add("@PrintedTaxCode", SqlDbType.VarChar).Value = Me.txtTaxCode.Text

            cmd.Parameters.Add("@CA", SqlDbType.VarChar).Value = Me.CmbTayBa.Text
            cmd.Parameters.Add("@Charge_D", SqlDbType.VarChar).Value = Me.CmbTVCharge1.Text & ":" & Me.txtTVCharge1.Text & "|" & Me.CmbTVCharge2.Text & ":" & Me.txtTVCharge2.Text
            cmd.Parameters.Add("@Discount_D", SqlDbType.VarChar).Value = Me.CmbTVDiscount1.Text & ":" & Me.txtTVDiscount1.Text & "|" & Me.CmbTVDiscount2.Text & ":" & Me.txtTVDiscount2.Text
            cmd.Parameters.Add("@SRV", SqlDbType.VarChar).Value = LclVarSRV
            cmd.Parameters.Add("@DOS", SqlDbType.DateTime).Value = Me.TxtDOS.Value.Date
            cmd.Parameters.Add("@Stock", SqlDbType.VarChar).Value = Me.CmbAL.Text
            cmd.Parameters.Add("@ROEID", SqlDbType.Int).Value = ROEID
            cmd.Parameters.Add("@RMK", SqlDbType.VarChar).Value = tmpRMK
            cmd.Parameters.Add("@RECID", SqlDbType.Int).Value = pubVarRCPID_BeingCreated
            cmd.ExecuteNonQuery()

            cmd.Parameters.Clear()
            cmd.CommandText = getSQL_GridTKT_to_TableTKT(pubVarRCPID_BeingCreated, pubVarSRV)
            cmd.ExecuteNonQuery()


            For i As Int16 = 0 To Me.GridFOP.RowCount - 1
                RMK = Me.GridFOP.Item("RMK", i).Value
                If String.IsNullOrEmpty(RMK) Then RMK = ""
                RMK = RMK.ToUpper
                If CDec(Me.GridFOP.Item("Amount", i).Value) <> 0 Or Me.GridFOP.Item("Document", i).Value <> "" Then
                    If InStr("MCO_EXC_PTA", Me.GridFOP.Item("FOP_FOP", i).Value) > 0 Then
                        tmpDocNo = AddSpace2TKNO(Me.GridFOP.Item("Document", i).Value)
                        If tmpDocNo.Substring(0, 3) <> "GRP" AndAlso CDec(Me.GridFOP.Item("Amount", i).Value) <> 0 Then
                            t.Rollback()
                            Return False
                        End If
                    Else
                        tmpDocNo = Me.GridFOP.Item("Document", i).Value
                    End If
                    If String.IsNullOrEmpty(tmpDocNo) Then tmpDocNo = ""
                    cmd.CommandText = Insert_FOP(pubVarRCPID_BeingCreated, Me.TxtTRXNO.Text, Me.GridFOP.Item("FOP_FOP", i).Value, _
                        Me.GridFOP.Item("RCPCurrency", i).Value, Me.GridFOP.Item("RCPROE", i).Value, _
                        CDec(Me.GridFOP.Item("Amount", i).Value), tmpDocNo.ToUpper, RMK, MyCust.CustID)
                    cmd.ExecuteNonQuery()
                    If Me.GridFOP.Item("FOP_FOP", i).Value = "RND" And Me.GridFOP.Item("Amount", i).Value * Me.GridFOP.Item("RCPROE", i).Value > 16000 Then
                        t.Rollback()
                        FeedBack("Err. Illogic RND FOP and Amount")
                        Return False
                    End If
                End If
            Next
            t.Commit()
            If CoPSP("ANY") Then
                cmd.Parameters.Clear()
                TKTList = ListOfTKTinRCP("INCLAUSE")
                cmd.CommandText = ChangeStatus_ByDK("TKT_1A_" & myStaff.City, "RE", " where status in ('OK') and TKNO in " & TKTList & " and SRV ='" & pubVarSRV & "'")
                cmd.ExecuteNonQuery()
            End If
            'update status=EX neu co ve doi
        Catch ex As Exception
            t.Rollback()
            MsgBox(Err.Description & vbCrLf & strSQL, MsgBoxStyle.Critical, msgTitle)
            Return False
        End Try
        cmd.CommandText = "Update TKT set status='EX' where status='OK' and tkno in (select stockCtrl from tkt where stockCtrl <>'' and " & _
            "rcpid=" & pubVarRCPID_BeingCreated & ")" & _
            "; Update TKT set status='EX' where status='OK' and tkno in (select Document from FOP where fop='EXC' and status<>'XX' " & _
                " and rcpid=" & pubVarRCPID_BeingCreated & ")"
        If CoVeDoi() Then cmd.ExecuteNonQuery()
        Me.CmdPreview.Enabled = True
        FeedBack("Transaction Saved ")
        Me.CmdSave.Visible = False
        Return True
    End Function
    Private Function CoVeDoi() As Boolean
        For r As Int16 = 0 To Me.GridFOP.RowCount - 1
            If Me.GridFOP.Item("FOP_FOP", r).Value = "EXC" Then
                Return True
            End If
        Next
        For r As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            If Me.GridTKT.Item("StockCtrl", r).Value <> "" Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Sub TxtPaxName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaxName.LostFocus
        If Me.CmbCustType.Text = "WK" Then
            Me.txtCustName.Text = Me.TxtPaxName.Text
        End If
        Me.CmdSaveChange.Enabled = True
    End Sub
    Private Function CoRQDepositHandle(ByVal pRCP_DocNo As String, ByVal pAction As String) As Boolean
        Dim KQ As Boolean, RQID As Integer, strDK As String
        strDK = "where tablename='RQDEPHANDLE' and doWhat='" & pAction & "' and "
        strDK = strDK & "F3='" & pRCP_DocNo & "'"
        RQID = ScalarToInt("ActionLog", "RecID", strDK)
        KQ = IIf(RQID = 0, False, True)
        Return KQ
    End Function
    Private Sub CmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDelete.Click
        Dim MyAns As Int16, Msg As String
        Msg = "Do You Really Want to Delete This Transaction? PLz Make Sure That It Was Issued Today"
        Msg = Msg & vbCrLf & "If You Choose YES, Remember to Attache All Copies of This Transaction to DailySalesReport."
        MyAns = MsgBox(Msg, MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
        If MyAns = vbNo Then Exit Sub
        If Me.GridTKT.RowCount > 0 AndAlso Me.GridTKT.Item("DocType", 0).Value = "GRP" Then
            If Not CoRQDepositHandle(mRCP, "OK-D") Then Exit Sub
        End If
        Msg = XXRCP(Me.TxtTRXNO.Text, myStaff.SICode)
        Me.CmdDelete.Enabled = False
        FeedBack(Msg)
        If Msg.Substring(0, 5) <> "Error" Then
            If pubVarSRV = "S" Then
                ResetEXstatusIfAny(pubVarRCPID_BeingEdited)
                Msg = "Do You Want to Create New Transaction Based on This One?"
                MyAns = MsgBox(Msg, MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
                If MyAns = vbNo Then
                    Me.Close()
                    Dim f As New FOissueTKT
                    f.ShowDialog()
                Else
                    Me.CmdPreview.Enabled = False
                    Me.TxtTRXNO.Text = GenRCPNo(MySession.TRXCode, MySession.POSCode)
                    mRCP = ""
                    Me.GrpGenInfor.Enabled = True
                    Me.GrpFare.Enabled = True
                    Me.GrpTax.Enabled = True
                    Me.GrpCharge.Enabled = True
                    Me.GrpPaymentDetails.Enabled = True
                    Me.CmdSave.Visible = True
                End If
            ElseIf pubVarSRV = "R" Then
                ResetVoidRecordedIfAny(pubVarRCPID_BeingEdited)
            End If
        End If
    End Sub
    Private Sub ResetVoidRecordedIfAny(ByVal pRCPID As Integer)
        Dim ChkID As Integer = ScalarToInt("TKT", "RecID", "RCPID=" & pRCPID & " and SRV='V'") ' vua R vua V cung RCP tuc la voidRecorded
        If ChkID = 0 Then Exit Sub
        Dim TKTList As String = ListOfTKTinRCP("INCLAUSE")
        If TKTList = "" Then Exit Sub
        cmd.CommandText = "update TKT set statusAL='OK' where statusAL+SRV+status='XXSOK' and tkt in " & TKTList
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub ResetEXstatusIfAny(ByVal pRCPID As Integer)
        Dim strDK As String, SubQryDoc As String, SubQryStockCtrl As String
        strDK = " where status='EX' and TKNO in "
        SubQryDoc = "(select document from FOP where fop='EXC' and len(rtrim(ltrim(document)))=15 and RCPID=" & pRCPID & ")"
        SubQryStockCtrl = "(select StockCtrl from TKT where RCPID=" & pRCPID & " and srv='S')"
        cmd.CommandText = ChangeStatus_ByDK("TKT", "OK", strDK & SubQryDoc) & "; " & ChangeStatus_ByDK("TKT", "OK", strDK & SubQryStockCtrl)
        cmd.CommandText = cmd.CommandText & _
            ";update tkt_1A_" & myStaff.City & " set status='OK' where status='RE' and tkno in (select tkno from tkt where rcpid=" & pRCPID & ")"
        cmd.ExecuteNonQuery()
    End Sub
    Private Function XXRCP(ByVal varRCPNo As String, ByVal strUser As String) As String
        Dim KQ As String, tmpInt As Int16, StrMSG As String, TKTList As String, FstUpdate As Date
        Dim strSQL As String, InvID As Integer, RecCount As Integer = 0, TravelID As Integer
        KQ = varRCPNo & " Has Been Deleted and Will Go to Daily Sales Report"
        If VeCongNoDaInv(pubVarRCPID_BeingEdited) Then
            KQ = "Error. TRX Has Been Invoiced To Client"
            MsgBox(KQ, MsgBoxStyle.Critical, msgTitle)
            Return KQ
        End If

        InvID = ScalarToInt("TKTNO_INVNO", "top 1 InvID", " RCPID=" & pubVarRCPID_BeingEdited & " and status='OK'")
        If InvID <> 0 Then
            tmpInt = ScalarToInt("INV", "PrintCopy", " RecID=" & InvID)
            If tmpInt > 0 Then
                KQ = "Error. INV Has Been Issued for This Transaction. Ask Your Supervisor or Acc. People to Delete Invoice First"
                MsgBox(KQ, MsgBoxStyle.Critical, msgTitle)
                Return KQ
            End If
        End If

        StrMSG = ScalarToString("RCP", "RPTNO", "RCPNO='" & varRCPNo & "'")
        If StrMSG <> "" Then
            KQ = "Error. This Transaction Has Been Reported and Cant Be Deleted"
            MsgBox(KQ, MsgBoxStyle.Critical, msgTitle)
            Return KQ
        End If
        If pubVarSRV = "S" Then
            TKTList = ListOfTKTinRCP("INCLAUSE")
            If TKTList <> "" Then RecCount = ScalarToInt("TKT", "count(*)", " Status <>'XX' and srv in ('R','V') and tkno in " & TKTList)
        End If

        FstUpdate = ScalarToDate("RCP", "FstUpdate", "recID=" & pubVarRCPID_BeingEdited)
        If FstUpdate.Date <> Now.Date Or RecCount > 0 Then   ' xoa khac ngay hoac co ve RV lien quan
            If myStaff.SupOf = "" Then ' ko co quyen cao thi ko xoa dc" 
                KQ = "Error. You Cant Delete This Transaction. Ask Your Manager for Help"
                MsgBox(KQ, MsgBoxStyle.Critical, msgTitle)
                Return KQ
            End If
        End If
        If InStr("CWT_N-A", MySession.Counter) > 0 Then
            TravelID = ScalarToInt("CWT.dbo.go_travel", "RecID", "dossier=" & pubVarRCPID_BeingEdited)
        End If
        Dim t As SqlClient.SqlTransaction = Conn.BeginTransaction
        cmd.Transaction = t
        Try
            strSQL = ChangeStatus_ByDK("INV", "X0", "status='OK' and recid=" & InvID)
            strSQL = strSQL & "; update TKTNO_INVNO set status='XX' where status='OK' and INVID=" & InvID
            strSQL = strSQL & "; update RCP set status='XX', RMK=RMK+'|XXby" & myStaff.SICode & Now & _
                "' where RecID=" & pubVarRCPID_BeingEdited
            strSQL = strSQL & ";" & ChangeStatus_ByDK("TKT", "XX", "RCPID=" & pubVarRCPID_BeingEdited, "XX")
            strSQL = strSQL & ";" & ChangeStatus_ByDK("FOP", "XX", "RCPID=" & pubVarRCPID_BeingEdited)
            If TravelID > 0 Then
                strSQL = strSQL & "; update cwt.dbo.go_travel set status='XX' where recid=" & TravelID
                strSQL = strSQL & "; delete from cwt.dbo.go_air where travelid=" & TravelID
                strSQL = strSQL & "; delete from cwt.dbo.go_hotel where travelid=" & TravelID
            End If
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()
            t.Commit()
            Dim LstUser As String = ScalarToString("TKT", "LstUser", "RCPID=" & pubVarRCPID_BeingEdited)
            If LstUser = "AUT" Then
                cmd.CommandText = "update tkt_1a_" & myStaff.City & " set status='XX', Lstupdate=getdate(), LstUser='" & _
                    myStaff.SICode & "' where status='RE' and AutoRas=2 and TKNO in (" & _
                    "select TKNO from tkt where rcpID=" & pubVarRCPID_BeingEdited & ")"
                cmd.ExecuteNonQuery()
            End If
        Catch ex As Exception
            t.Rollback()
            KQ = " Error Deleting TRX"
            MsgBox(KQ, MsgBoxStyle.Critical, msgTitle)
        End Try
        Return KQ
    End Function
    Private Sub txtALCommVAL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtALCommVAL.LostFocus
        If Me.CmbProduct.Text = "EXC" Then
            Me.txtExcDoc.Enabled = True
        Else
            Me.txtExcDoc.Enabled = False
        End If
        If Me.CmbProduct.Text = "???" Then
            Me.txtALCommPCT.Enabled = False
            Me.txtNetToAL.Text = Format(CDec(Me.txtFare.Text) - CDec(Me.txtALCommVAL.Text), "#,##0.00")
            Me.txtALCommPCT.Text = Format(CDec(Me.txtALCommVAL.Text) / CDec(Me.txtFare.Text) * 100, "0.00")
        End If
    End Sub
    Private Sub TxtPaid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPaid.TextChanged
        If CDec(Me.TxtPaid.Text) = CDec(Me.txtVNDEquivalent.Text) Then
            Me.CmdSave.Enabled = True
            Me.CmdChangeFOP.Enabled = True
        Else
            Me.CmdSave.Enabled = False
            Me.CmdChangeFOP.Enabled = False
            Me.CmdPreview.Enabled = False
        End If
    End Sub

    Private Sub CmdChangeFOP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdChangeFOP.Click
        Dim myAns As Int16, StrQ As String, tmpCustID As Integer
        Dim strSQL As String, RPTNO As String
        If myStaff.SICode <> "SYS" And (InStr("XXX_ALL", myStaff.Counter) > 0 Or myStaff.Counter = "") Then Exit Sub
        If VeCongNoDaInv(pubVarRCPID_BeingEdited) Then
            MsgBox("This TRX has Been Invoiced And Cant Be Changed", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        RPTNO = ScalarToString("RCP", "RPTNO", "Recid=" & pubVarRCPID_BeingEdited)
        tmpCustID = ScalarToInt("RCP", "CustID", "RecID=" & pubVarRCPID_BeingEdited)
        If RPTNO <> "" Then
            MsgBox("TRX Has Been Reported or Invoiced.", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If CoPSP("PSP") Then
            StrQ = CheckOverDue(Me.CmbChannel.SelectedValue, Conn)
            If StrQ <> "" Then
                MsgBox(StrQ, MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        End If

        If Not CRD_format_Correct() Then
            FeedBack("Err. Invalid CRD No")
            Exit Sub
        End If

        If Not NoOK() Then Exit Sub

        StrQ = "PLz Make Sure All Inputs Are Correct Before Clicking OK"
        myAns = MsgBox(StrQ, MsgBoxStyle.OkCancel + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2, msgTitle)
        If myAns = vbCancel Then Exit Sub

        Me.CmdChangeFOP.Enabled = False
        Dim t As SqlClient.SqlTransaction = Conn.BeginTransaction
        Dim DocNo As String = "", RMK As String = ""
        Try
            strSQL = ChangeStatus_ByDK("FOP", "XX", "status in ('OK','QQ') and RCPID=" & pubVarRCPID_BeingEdited)
            For r As Int16 = 0 To Me.GridFOP.RowCount - 1
                DocNo = Me.GridFOP.Item("Document", r).Value
                RMK = Me.GridFOP.Item("RMK", r).Value
                If String.IsNullOrEmpty(DocNo) Then DocNo = ""
                If String.IsNullOrEmpty(RMK) Then RMK = ""
                If CDec(Me.GridFOP.Item("Amount", r).Value) <> 0 Or Me.GridFOP.Item("Document", r).Value <> "" Then
                    strSQL = strSQL & "; " & Insert_FOP(pubVarRCPID_BeingEdited, Me.TxtTRXNO.Text, Me.GridFOP.Item("FOP_FOP", r).Value, _
                        Me.GridFOP.Item("RCPCurrency", r).Value, Me.GridFOP.Item("RCPROE", r).Value, _
                        CDec(Me.GridFOP.Item("Amount", r).Value), DocNo.ToUpper, RMK.ToUpper, tmpCustID)
                End If
                If Me.GridFOP.Item("FOP_FOP", r).Value = "RND" And Me.GridFOP.Item("Amount", r).Value * Me.GridFOP.Item("RCPROE", r).Value > 16000 Then
                    FeedBack("Err. Illogic RND FOP and Amount")
                    Exit Sub
                End If
            Next
            cmd.Transaction = t
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()
            t.Commit()
            Me.CmdChangeFOP.Visible = False
            FeedBack("FOP Has Been Updated")
            If CoPSP("ANY") Then RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, False, "After CHG FOP " & Me.TxtTRXNO.Text, Conn, myStaff.SICode, CnStr)
        Catch ex As Exception
            t.Rollback()
            MsgBox("Error Writing into DataBase. Try Again", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub

    Private Sub CmdSaveNON_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSaveNON.Click
        If myStaff.SICode <> "SYS" And (InStr("XXX_ALL", myStaff.Counter) > 0 Or myStaff.Counter = "") Then Exit Sub
        Dim myAns As Int16, RecNo As Integer, strSQL As String
        Me.CmdSaveNON.Enabled = False
        strSQL = "Are You Sure that All Changes You Have Made Are Not Money Related?"
        myAns = MsgBox(strSQL, MsgBoxStyle.YesNo + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton2, msgTitle)
        If myAns = vbYes Then
            Dim t As SqlClient.SqlTransaction = Conn.BeginTransaction
            Try
                strSQL = ""
                For r As Int16 = 0 To Me.GridTKT.RowCount - 1
                    If Me.GridTKT.Item("isEdit", r).Value = "Y" Then
                        RecNo = Me.GridTKT.Item("recID", r).Value
                        strSQL = UpdateLogFile("TKT", "RecID:" & RecNo.ToString, _
                            "DOI:" & Me.GridTKT.Item("DOI", r).Value & "_DOF" & Me.GridTKT.Item("DOF", r).Value, _
                            "RBD:" & Me.GridTKT.Item("BkgClass", r).Value & "_" & Me.GridTKT.Item("DocType", r).Value, _
                            Me.GridTKT.Item("PaxType", r).Value & "_" & Me.GridTKT.Item("Paxname", r).Value, _
                            "TC:" & Me.GridTKT.Item("TourCode", r).Value & "_PC:" & Me.GridTKT.Item("promocode", r).Value, _
                            Me.GridTKT.Item("Routing", r).Value, Me.GridTKT.Item("FareBasis", r).Value, _
                            Me.GridTKT.Item("TKNO", r).Value & "_Conj" & Me.GridTKT.Item("FTKT", r).Value, "")
                        strSQL = strSQL & "; update TKT Set "
                        strSQL = strSQL & " DOF='" & Me.GridTKT.Item("DOF", r).Value & "',"
                        strSQL = strSQL & " DOI='" & Me.GridTKT.Item("DOI", r).Value & "',"
                        strSQL = strSQL & " BkgClass='" & Me.GridTKT.Item("BKGCLASS", r).Value.Replace("--", "") & "',"
                        strSQL = strSQL & " DocType='" & Me.GridTKT.Item("DocType", r).Value & "',"
                        strSQL = strSQL & " PaxType='" & Me.GridTKT.Item("PaxType", r).Value.ToString.Substring(0, 3) & "',"
                        strSQL = strSQL & " PaxName='" & Me.GridTKT.Item("PaxName", r).Value.ToString.Replace("--", "") & "',"
                        strSQL = strSQL & " TourCode='" & Me.GridTKT.Item("tourCode", r).Value.ToString.Replace("--", "") & "',"
                        strSQL = strSQL & " PromoCode='" & Me.GridTKT.Item("Promocode", r).Value & "',"
                        strSQL = strSQL & " Itinerary='" & Me.GridTKT.Item("Routing", r).Value.ToString.Replace("--", "") & "',"
                        strSQL = strSQL & " FareBasis='" & Me.GridTKT.Item("Farebasis", r).Value.ToString.Replace("--", "") & "',"
                        strSQL = strSQL & " Tax_D='" & Me.GridTKT.Item("Tax_D", r).Value & "',"
                        strSQL = strSQL & " Currency='" & Me.CmbCurr.Text & "',"
                        strSQL = strSQL & " StockCtrl='" & Me.GridTKT.Item("StockCtrl", r).Value.ToString.Replace("--", "") & "',"
                        strSQL = strSQL & " FareType='" & Me.GridTKT.Item("FareType", r).Value & "',"
                        If WhatAction = "NON" AndAlso myStaff.SupOf <> "" Then
                            strSQL = strSQL & " TKNO='" & Me.GridTKT.Item("TKNO", r).Value & "',"
                            strSQL = strSQL & " FTKT='" & Me.GridTKT.Item("FTKT", r).Value & "',"
                        End If
                        strSQL = strSQL & " RMK='" & Me.GridTKT.Item("TKT_RMK", r).Value.ToString.Replace("--", "") & "',"
                        strSQL = strSQL & " LstUser='" & myStaff.SICode & "',"
                        strSQL = strSQL & " LstUpdate=getdate()"
                        strSQL = strSQL & " where recID=" & RecNo
                    End If
                Next
                If strSQL.Trim.Substring(0, 1) = ";" Then strSQL = strSQL.Trim.Substring(1)
                cmd.Transaction = t
                cmd.CommandText = strSQL
                cmd.ExecuteNonQuery()
                t.Commit()
                FeedBack("Changes Updated! ")
            Catch ex As Exception
                t.Rollback()
                MsgBox("Error Writing into DataBase. Try Again", MsgBoxStyle.Critical, msgTitle)
            End Try
        End If
    End Sub

    Private Sub CmdAdjust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdAdjust.Click
        If SysDateIsWrong(Conn) Then Exit Sub
        If myStaff.SICode <> "SYS" And (InStr("XXX_ALL", myStaff.Counter) > 0 Or myStaff.Counter = "") Then Exit Sub
        Dim myAns As Int16, strSQL As String, ChenhLech As Decimal, oldFOP As String = "", OldTTLDue As Decimal
        Dim tmpCurr As String = "", tmpAmt As Decimal, tmpROE As Decimal, tmpCustID As Integer
        strSQL = "This May Auto Edit Payment Infor. Plz Be Sure of Your Input. Abort?"
        myAns = MsgBox(strSQL, MsgBoxStyle.YesNo + MsgBoxStyle.Critical + MsgBoxStyle.DefaultButton1, msgTitle)
        If myAns = vbYes Then Exit Sub
        Me.CmdAdjust.Enabled = False

        If VeCongNoDaInv(pubVarRCPID_BeingEdited) Then
            MsgBox("This TRX has Been Invoiced And Cant Be Changed", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        ' Tinh da tra/ ConPhaiTra

        OldTTLDue = ScalarToDec("RCP", "TTLDue", " RecID=" & pubVarRCPID_BeingEdited)
        tmpCustID = ScalarToInt("RCP", "CustID", "RecID=" & pubVarRCPID_BeingEdited)
        ChenhLech = CDec(Me.txtNettPayable.Text) - OldTTLDue

        If ChenhLech <> 0 Then
            Dim INVID As Integer = ScalarToInt("INV", "REcID", "RcpID=" & pubVarRCPID_BeingEdited & " and status='OK'")
            If INVID > 0 Then
                Dim invCopy As Int16 = ScalarToInt("INV", "PrintCopy", "RecID=" & INVID)
                If invCopy > 0 Then
                    MsgBox("Invoice Has Been Issued. Please Contact Acc. and/or Delete it First", MsgBoxStyle.Critical, msgTitle)
                    Exit Sub
                Else
                    cmd.CommandText = ChangeStatus_ByID("INV", "X0", INVID) & ";" & _
                        ChangeStatus_ByDK("TKTNO_INVNO", "XX", "INVID=" & INVID)
                    cmd.ExecuteNonQuery()
                End If
            End If
            ' XacDinh FOP can update vao FOP
            oldFOP = ScalarToString("FOP", "top 1 FOP", "fop in ('PSP','PPD') and status <>'XX' and RCPID=" & pubVarRCPID_BeingEdited & " order by FOP")
            If oldFOP = "" OrElse oldFOP Is Nothing Then oldFOP = "DEB"
            If oldFOP = "PPD" Then
                tmpCurr = "VND"
                tmpROE = 1
                tmpAmt = ChenhLech * CDec(Me.LcktxtROE.Text)
            Else
                tmpCurr = Me.CmbCurr.Text
                tmpROE = CDec(Me.LcktxtROE.Text)
                tmpAmt = ChenhLech
            End If
        End If
        'Luu gia tri moi trong TKT
        strSQL = " update TKT set status='XX', statusAL='XX' where  RCPID=" & pubVarRCPID_BeingEdited
        strSQL = strSQL & ";" & getSQL_GridTKT_to_TableTKT(pubVarRCPID_BeingEdited, pubVarSRV)
        Dim t As SqlClient.SqlTransaction = Conn.BeginTransaction
        Try
            If ChenhLech <> 0 Then
                ' Luu ban ghi RCP hien tai + update TTLDue moi
                strSQL = strSQL & ";" & UpdateLogFile("RCP", "EDITMON", "OldTTLDue:" & OldTTLDue, "NewTTLDue" & Me.txtNettPayable.Text, "", "", "", "", "", "")
                strSQL = strSQL & "; Update RCP set TTLDue=" & CDec(Me.txtNettPayable.Text)
                strSQL = strSQL & ", Discount=" & CDec(Me.txtTVDiscount1.Text) + CDec(Me.txtTVDiscount2.Text)
                strSQL = strSQL & ", Charge=" & CDec(Me.txtTVCharge1.Text) + CDec(Me.txtTVCharge2.Text) + CDec(Me.txtDiscountedAmount.Text)
                strSQL = strSQL & ", Charge_D='" & Me.CmbTVCharge1.Text & ":" & Me.txtTVCharge1.Text & "|"
                strSQL = strSQL & Me.CmbTVCharge2.Text & ":" & Me.txtTVCharge2.Text & "'"
                strSQL = strSQL & ", Discount_D='" & Me.CmbTVDiscount1.Text & ":" & Me.txtTVDiscount1.Text & "|"
                strSQL = strSQL & Me.CmbTVDiscount2.Text & ":" & Me.txtTVDiscount2.Text & "'"
                strSQL = strSQL & " where recid=" & pubVarRCPID_BeingEdited
                ' Insert chenhlech vao FOP
                strSQL = strSQL & ";" & Insert_FOP(pubVarRCPID_BeingEdited, Me.TxtTRXNO.Text, oldFOP, tmpCurr, tmpROE, tmpAmt, "", "AUTO ADJ AMT", tmpCustID)
                FeedBack("Changes Updated!")
                If oldFOP = "DEB" Then
                    MsgBox("The Difference Has Been Recorded as DEB. Plz Go to Clear Pending Payment to Clear This Amount", MsgBoxStyle.Exclamation, msgTitle)
                End If
            End If
            cmd.Transaction = t
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()
            t.Commit()

            If ChenhLech <> 0 And InStr("PSP_PPD", MyCust.DelayType) > 0 Then RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, False, "After Adj " & Me.TxtTRXNO.Text, Conn, myStaff.SICode, CnStr)
        Catch ex As Exception
            t.Rollback()
            MsgBox("Error Writing into DataBase. Try Again", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub
    Private Function ListOfTKTinRCP(ByVal pReturnType As String) As String
        Dim KQ As String = "", KQ_inClause As String = ""
        For i As Int16 = 0 To Me.GridTKT.RowCount - 1
            If pReturnType <> "SHORT" Then
                KQ = KQ & "_" & Me.GridTKT.Item("TKNO", i).Value
            ElseIf Me.GridTKT.Item("Qty", i).Value <> 0 Then
                KQ = KQ & "_" & Me.GridTKT.Item("TKNO", i).Value.ToString.Substring(4).Replace(" ", "")
            End If
        Next
        If KQ.Length > 2 Then
            KQ = KQ.Substring(1)
            KQ_inClause = KQ.Replace("_", "','")
            KQ_inClause = "('" & KQ_inClause & "')"
        End If
        If pReturnType = "INCLAUSE" Then
            Return KQ_inClause
        Else
            Return KQ
        End If
    End Function
    Private Sub BarRQ4CRDExt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarRQ4CRDExt.Click
        Dim strDK As String = "", VND_DelayedPmt As String, VND_Delayed_Amt As Decimal
        Dim ExistRQ As Integer, TKTinRCP As String = ListOfTKTinRCP("SHORT")
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim StaffEmail As String = ""
        If Me.TxtTRXNO.Text.Substring(0, 2) = "TS" Then
            If InStr("CA_TA_TO", MyCust.CustType) > 0 Then
                If MyAL.ALCode = "VN" Then
                    StaffEmail = "tocsvn"
                Else
                    StaffEmail = "tocs"
                End If
            ElseIf MyCust.CustType = "WK" Then
                StaffEmail = "travelshop"
            End If
        Else
            StaffEmail = Me.TxtTRXNO.Text.Substring(0, 2)
        End If
        VND_DelayedPmt = XD_DEB_PPD_PSP_Amt()
        VND_Delayed_Amt = 1 + CDec(VND_DelayedPmt.Substring(3))
        strDK = " tableName='CREXT' and Dowhat <>'XX' and F1='" & Me.TxtTRXNO.Text & "' and F2=" & Me.CmbChannel.SelectedValue
        strDK = strDK & " and F6='" & VND_DelayedPmt.Substring(0, 3) & "' And F5 = " & VND_Delayed_Amt
        ExistRQ = ScalarToInt("ActionLog", "RecID", strDK)
        If ExistRQ = 0 Then
            cmd.CommandText = UpdateLogFile("CREXT", "QQ", Me.TxtTRXNO.Text, Me.CmbChannel.SelectedValue, Me.CmbChannel.Text, _
                    TKTinRCP, VND_Delayed_Amt, VND_DelayedPmt.Substring(0, 3), myStaff.SICode, "", StaffEmail, MyCust.CustType, pubVarRCPID_BeingCreated)
            cmd.ExecuteNonQuery()
        End If
        Me.BarRQ4CRDExt.Enabled = False
        MsgBox("This Transaction Has Been Q for Approval", MsgBoxStyle.Information, msgTitle)
    End Sub
    Private Sub LblTKNO_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblTKNO.DoubleClick
        Dim tmpINV As String
        If WhatAction = "NON" AndAlso myStaff.SupOf <> "" AndAlso _
            Me.GridTKT.CurrentRow.Cells("FTKT").Value = "" AndAlso InStr("SV", pubVarSRV) > 0 Then
            tmpINV = ScalarToString("TKTNO_INVNO", "INVNO", " status='OK' and TKNO='" & Me.txtTKNO.Text & "'")
            If tmpINV <> "" Then
                MsgBox("Invoice Has Been Issued For This TKT", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
            Me.txtTKNO.Enabled = True
        End If
        If pubVarSRV = "S" And mRCP = "" Then
            Me.txtTKNO.Enabled = True
        End If
    End Sub
    Private Sub txtCheckDigit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCheckDigit.LostFocus
        If Me.txtTKNO.Enabled = False Then Exit Sub
        Dim KQ As Integer, tmp As String = ""
        tmp = Me.txtTKNO.Text.Replace(" ", "")
        If Not MyAL.isTKTLess Then
            KQ = CDec(tmp) Mod 7
            If Me.txtCheckDigit.Enabled And (Me.txtCheckDigit.Text = "" OrElse CInt(Me.txtCheckDigit.Text) <> KQ) Then
                MsgBox("Invalid Ticket Number", MsgBoxStyle.Critical, msgTitle)
                Me.txtCheckDigit.Text = ""
                Me.txtTKNO.Focus()
            End If
        End If
    End Sub

    Private Sub GridTKT_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridTKT.CellContentDoubleClick
        If e.RowIndex < 0 Then Exit Sub
        Dim tmpINV As String, strSQL As String
        If e.ColumnIndex = 1 AndAlso myStaff.SupOf <> "" AndAlso _
            WhatAction = "NON" AndAlso InStr("SV", pubVarSRV) > 0 Then
            tmpINV = ScalarToString("TKTNO_INVNO", "INVNO", " status='OK' and TKNO='" & Me.txtTKNO.Text & "'")
            If tmpINV <> "" Then
                MsgBox("Invoice Has Been Issued For This TKT", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
            On Error GoTo ErrHandler
            If Me.GridTKT.CurrentCell.Value.ToString.Substring(0, 4) = "___/" Then
                Me.GrpFTKT2Edit.Visible = True
            End If
            On Error GoTo 0
        End If
ErrHandler:
    End Sub

    Private Sub CmdFTKT2EditCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdFTKT2EditCancel.Click
        Me.GrpFTKT2Edit.Visible = False
    End Sub

    Private Sub TxtFTKT2Edit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFTKT2Edit.LostFocus
        Dim TKNO As String, DaSD As Boolean = False
        If ESCPressed Then Exit Sub
        TKNO = Me.TxtFTKT2Edit.Text
        If Not MyAL.isTKTLess Then
            TKNO = CheckTKTformat(TKNO)
            If TKNO.Substring(0, 3).ToUpper = "ERR" Then
                FeedBack("Error. Invalid Ticket Number!")

                Me.TxtFTKT2Edit.Focus()
                Exit Sub
            Else
                Me.TxtFTKT2Edit.Text = TKNO
            End If
        End If
        If pubVarSRV <> "A" Then DaSD = isAlreadyInGrid(TKNO)
        If DaSD Then
            FeedBack("Error. This Ticket Has Been In List!")
            Me.TxtFTKT2Edit.Focus()
            Exit Sub
        End If
        Exit Sub
    End Sub
    Private Sub CmdFTKT2EditOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LckCmdFTKT2EditOK.Click
        Dim RowNo As Int16 = Me.GridTKT.CurrentCell.RowIndex
        Dim myAns As Int16, VeNay As String
        Dim Conj As String = "", tmpPhanSau As String, tmpConjTruoc As String, tmpConjSau As String
        Dim ArrTKT(4) As String, ArrConj(4) As String
        Dim ListOfTKT As String = ""
        Dim PhanDau As String, PhanSau As Decimal, i As Int16
        Dim MsgStr As String
        PhanDau = Me.TxtFTKT2Edit.Text.Substring(0, 9)
        PhanSau = CDec(Strings.Right(Me.TxtFTKT2Edit.Text, 6))
        i = 0
        Do
            tmpPhanSau = "00000" & Trim(Str(PhanSau + i))
            tmpPhanSau = tmpPhanSau.Substring(tmpPhanSau.Length - 6, 6)
            VeNay = PhanDau & tmpPhanSau
            tmpConjSau = "00000" & Trim(Str(PhanSau + i + 1))
            tmpConjSau = tmpConjSau.Substring(tmpConjSau.Length - 6, 6)
            tmpConjTruoc = "00000" & Trim(Str(PhanSau + i - 1))
            tmpConjTruoc = tmpConjTruoc.Substring(tmpConjTruoc.Length - 6, 6)
            If i = 0 Then
                Conj = "___/" & tmpConjSau.Substring(3, 3)
            ElseIf Me.GridTKT.Item("FTKT", RowNo).Value.ToString.Trim.Length = 3 Then
                Conj = tmpConjTruoc.Substring(3, 3)
            Else
                Conj = tmpConjTruoc.Substring(3, 3) & "/" & tmpConjSau.Substring(3, 3)
            End If
            ArrTKT(i) = VeNay
            ArrConj(i) = Conj
            ListOfTKT = ListOfTKT & VeNay & "|" & Conj & vbCrLf
            If Me.GridTKT.Item("FTKT", RowNo).Value.ToString.Trim.Length = 3 Then Exit Do
            RowNo = RowNo + 1
            i = i + 1
        Loop
        MsgStr = "Change First TKNO Will Change All the Conjuncted TKTs as Followings:"
        MsgStr = MsgStr & vbCrLf & ListOfTKT
        MsgStr = MsgStr & vbCrLf & "Agree?"
        MsgStr = MsgStr & vbCrLf & "WARNING: Be Careful Before Clicking YES!!! "
        myAns = MsgBox(MsgStr, MsgBoxStyle.Critical Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
        If myAns = vbNo Then GoTo ThoatODay

        RowNo = Me.GridTKT.CurrentCell.RowIndex
        i = 0
        Do
            Me.GridTKT.Item("FTKT", RowNo).Value = ArrConj(i)
            Me.GridTKT.Item("TKNO", RowNo).Value = ArrTKT(i)
            Me.GridTKT.Item("isEdit", RowNo).Value = "Y"
            If Me.GridTKT.Item("FTKT", RowNo).Value.ToString.Trim.Length = 3 Then Exit Do
            i = i + 1
            RowNo = RowNo + 1
        Loop
ThoatODay:
        Me.TxtFTKT2Edit.Text = ""
        Me.GrpFTKT2Edit.Visible = False
    End Sub

    Private Sub LblEXCdoc_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LblEXCdoc.DoubleClick, Label47.DoubleClick
        If pubVarSRV = "S" Then
            Me.txtExcDoc.Enabled = True
            Me.txtExcDoc.Focus()
        End If
    End Sub

    Private Sub txtExcDoc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExcDoc.LostFocus
        Dim tmp As String = Me.txtExcDoc.Text
        tmp = AddSpace2TKNO(tmp)
        Me.txtExcDoc.Text = tmp
        If tmp <> "" And Not IsValidExcDoc(tmp, CDec(Me.txtEXCValue.Text), Me.txtFltDate.Value.Date, pubVarRCPID_BeingEdited + pubVarRCPID_BeingCreated) Then
            Me.txtExcDoc.Focus()
        End If
    End Sub
    Private Function IsValidExcDoc(ByVal pDocNo As String, ByVal pAmt As Decimal, pDOF As Date, pRCPID As Integer) As Boolean
        Dim tmp As String
        Dim StrDKOK As String = " tkno ='" & pDocNo & "' and Status='OK'", strDKEX As String

        Dim MyAns As Int16, tmpExcDetail As String, strSQL As String
        Dim IssAL As String = "", IssToCust As Integer, ValidUntil As Date, IssRCPID As Integer
        If QuitNow Then Return True
        If pDocNo.Length = 0 Then Return True
        If pDocNo.Substring(0, 1) <> "Z" AndAlso pDocNo.Length <> 15 Then
            FeedBack("Error. Invalid Document format/length!")
            Return False
        End If
        tmp = ScalarToString("TKT", "TKNO", "tkno ='" & pDocNo & "'")
        strDKEX = " tkno ='" & pDocNo & "' and srv='S' and tkno not in (select TKNO from tkt where status='OK' and SRV in ('R','V')) " & _
                    " and (FTKT='' or left(ftkt,3)='___') and (status='OK' or Rcpid <>" & pRCPID & ")"
        tmp = ScalarToString("TKT", "TKNO", strDKEX)
        If tmp <> "" Then
            IssAL = ScalarToString("TKT", "AL", StrDKOK)
            ValidUntil = ScalarToDate("TKT", "DOI", StrDKOK)
            ValidUntil = ValidUntil.AddDays(365)
            IssRCPID = ScalarToInt("TKT", "RCPID", StrDKOK)
            IssToCust = ScalarToInt("RCP", "CustID", "Recid=" & IssRCPID)
            If IssAL <> Me.CmbAL.Text Or IssToCust <> MyCust.CustID Or ValidUntil < Now.Date Then
                FeedBack("Error. Invalid AL/CustID of Exc Doc or validity expired!")
                Return False
            End If
            If pDocNo.Substring(0, 3) = "GRP" Then
                Dim AmtDepo As Decimal = ScalarToDec("TKT", "Fare", StrDKOK)
                If AmtDepo < pAmt Then Return False
                Dim DOF As Date = ScalarToDate("TKT", "DOF", StrDKOK)
                If DOF <> pDOF Then Return False ' phai cho cung ngay bay. neu roi coc thi phai issue cai khac
            End If
            Return True
        Else
            If pDocNo.Substring(0, 3) = "GRP" Then Return False
            strSQL = "Document " & pDocNo & " Could Not Be Found or Has Already Been Exchanged/Refunded"
            strSQL = strSQL & vbCrLf & vbCrLf & " - Make Sure that You Have Rights to EXC TKT that Is not in Our Dbase, or"
            strSQL = strSQL & vbCrLf & " - If It Was Conjuncted TKT, Make Sure that You Input The First One, or"
            strSQL = strSQL & vbCrLf & " - Wanna Correct the Input?"
            MyAns = MsgBox(strSQL, MsgBoxStyle.Critical Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1, msgTitle)
            If MyAns = vbYes Or myStaff.SupOf = "" Then
                Return False
            Else
KoTimThay:
                MyAns = MsgBox("Has " & pDocNo & " Been Issued By TransViet?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1, msgTitle)
                If MyAns = vbYes Then
                    MyAns = MsgBox("Has " & pDocNo & " Been Recorded into Old Version of RAS?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1, msgTitle)
                    If MyAns = vbYes Then
                        tmpExcDetail = InputBox("Enter the Original Receipt Of " & pDocNo, msgTitle)
                        If tmpExcDetail = "" Then
                            Return False
                        Else
                            Me.txtExcDocDetail.Text = tmpExcDetail.ToUpper
                            ExcDocDetail = ExcDocDetail & "|TRX" & tmpExcDetail.ToUpper
                            Me.txtExcDocDetail.Visible = True
                            Return True
                        End If
                    Else
                        Return False
                    End If
                Else
                    tmpExcDetail = InputBox("Enter the IATA Code for the Place of Issue of " & pDocNo, msgTitle)
                    If tmpExcDetail = "" Then
                        Return False
                    Else
                        Me.txtExcDocDetail.Text = tmpExcDetail.ToUpper
                        ExcDocDetail = ExcDocDetail & "|POI" & tmpExcDetail.ToUpper
                        Me.txtExcDocDetail.Visible = True
                        Return True
                    End If
                End If
            End If
        End If
    End Function
    Private Sub GrpFare_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles GrpFare.Enter
        If Me.CmbProduct.Text = "" Then
            Me.CmbProduct.Text = "FIT"
        End If
    End Sub
    Private Sub GrpFare_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles GrpFare.Leave
        Dim f As Decimal, n2al As Decimal, PCT As Decimal, VAL As Decimal
        If Me.txtNetToAL.Enabled Or Me.txtALCommVAL.Enabled Or Me.txtALCommPCT.Enabled Then
            f = CDec(Me.txtFare.Text)
            n2al = CDec(Me.txtNetToAL.Text)
            PCT = CDec(Me.txtALCommPCT.Text)
            VAL = CDec(Me.txtALCommVAL.Text)
            If (f = n2al + VAL And Me.CmbProduct.Text = "S+M") Or _
                (n2al + VAL = 0 And Me.CmbProduct.Text = "S+M") Then
                MsgBox("Check Your Input!", MsgBoxStyle.Critical, msgTitle)
                Me.txtFare.Focus()
            ElseIf VAL < 0 And pubVarSRV = "S" Then 'Lost Sales
                If myStaff.SupOf <> "" Then ' Duoc Phep Ban lost sales
                    MsgBox("Be Carefull. Lost Sales", MsgBoxStyle.Critical, msgTitle)
                Else
                    MsgBox("Check Your Input!", MsgBoxStyle.Critical, msgTitle)
                    Me.txtFare.Focus()
                End If
            End If
        End If
    End Sub
    Private Sub ReCalcALComm()
        Dim f As Decimal, hh As Decimal
        Me.TxtAgtDiscPCT.Text = 0
        Me.TxtAgtDiscVal.Text = 0
        If InStr("FOC_NET_EXC_S+M_???", Me.CmbProduct.Text.Trim) > 0 Then Exit Sub
        f = CDec(Me.txtFare.Text)
        hh = XDHoaHongGSA(Me.CmbAL.Text, Me.CmbProduct.Text, Me.CmbCurr.Text)
        If hh > 1 Then
            Me.txtALCommVAL.Text = Format(hh, "#,##0.00")
            Me.txtNetToAL.Text = Format(f - hh, "#,##0.00")
        ElseIf hh > 0 And hh < 1 Then
            Me.txtALCommPCT.Text = Format(hh * 100, "0.00")
            Me.txtNetToAL.Text = Format((1 - hh) * f, "#,##0.00")
            Me.txtALCommVAL.Text = Format(f - CDec(Me.txtNetToAL.Text), "#,##0.00")
        ElseIf hh = 0 Then
            Me.txtALCommPCT.Text = Format(0, "0.00")
            Me.txtNetToAL.Text = Format(f, "#,##0.00")
            Me.txtALCommVAL.Text = Format(0, "0.00")
        End If
    End Sub
    Private Sub TxtPaxName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPaxName.TextChanged
        If Me.TxtPaxName.Text <> "" Then Me.GrpFare.Enabled = True
    End Sub
    Private Sub txtFltDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFltDate.TextChanged
        Me.TxtPaxName.Enabled = True
    End Sub
    Private Sub LblAutoTKTClose_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAutoTKTClose.LinkClicked
        Me.PnlAutoTKT.Visible = False
        Me.MenuStrip2.Enabled = Not Me.PnlAutoTKT.Visible
    End Sub

    Private Sub BarTKTList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarTKTListSV.Click
        Dim strSQL As String
        If Me.PnlAutoTKT.Visible Then Exit Sub
        strSQL = "select distinct TKNO, FTKT, fullRTG as Itinerary, PaxName, Fare*qty as fare, Tax*qty as tax, Charge*qty " & _
            "as charge, BkgClass, Qty, CommVAL*qty as CommVal, DOI, DOF, FareBasis, TourCode, ROE, Booker," & _
            "NetToAL*qty as NetToAL, PaxType, left(StockCtrl,13) as StockCtrl, TaxDetail, ShownFare*Qty as ShownFare, SRV, DocType," & _
            "Currency, svcfee, 0 as LLF, TinhTrang, ChargeDetail, svcfeeCur from tkt_1a where len(tkno)>=8 and srv='" & pubVarSRV & "' " & _
            " and CustID=" & Me.CmbChannel.SelectedValue & " and doi >'01-Jan-15'"
        If MySession.Domain = "GSA" Then
            strSQL = strSQL & " and len(counter)=2 "
        Else
            strSQL = strSQL & " and len(counter)=3 "
        End If
        If MyAL.DocCode <> "ZXX" Then
            If MyAL.ALCode = "01" Then
                strSQL = strSQL & "and AL='01'"
            Else
                strSQL = strSQL & " and charindex(left(tkno,3),'" & MyAL.ValidDocCode & "')<>0"
            End If
        End If
        If pubVarSRV = "S" Then
            If MyAL.ALCode = "01" Then
                strSQL = strSQL & " and tkno not in (select dependent from TKT where status <>'XX') " & _
                    " and substring(tkno,5,10) not in (select substring(dependent,4,10) from TKT where dependent <>'' and status <>'XX') "
            Else
                strSQL = strSQL & " and tkno not in (select TKNO from TKT where status <>'XX') "
            End If
        End If
        strSQL = strSQL & " order by doi desc"
        Me.GridAutoTKT.DataSource = GetDataTable(strSQL)
        Me.GridAutoTKT.Columns("FTKT").Width = 56
        Me.GridAutoTKT.Columns("PaxType").Width = 32
        Me.GridAutoTKT.Columns("BkgClass").Width = 32
        Me.GridAutoTKT.Columns("Fare").Width = 56
        Me.GridAutoTKT.Columns("Tax").Width = 56
        Me.GridAutoTKT.Columns("Charge").Width = 56
        Me.GridAutoTKT.Columns("Qty").Width = 32
        Me.GridAutoTKT.Columns("Currency").Width = 32
        Me.GridAutoTKT.Columns("DOI").Width = 64
        Me.GridAutoTKT.Columns("DOF").Width = 64
        Me.GridAutoTKT.Columns("ROE").Width = 64
        Me.GridAutoTKT.Columns("Itinerary").Width = 196
        Me.GridAutoTKT.Columns("PaxName").Width = 196
        Me.GridAutoTKT.Columns("FareBasis").Width = 196
        Me.GridAutoTKT.Columns("Qty").Visible = False
        Me.PnlAutoTKT.Visible = True
        Me.PnlAutoTKT.Top = 0
        Me.PnlAutoTKT.Left = 0
        Me.PnlAutoTKT.Width = 797
        Me.PnlAutoTKT.Height = 373
        Me.MenuStrip2.Enabled = Not Me.PnlAutoTKT.Visible
    End Sub
    Private Function NewInsurance(pVcrNo As String) As Boolean
        Dim tmpRec As Integer = ScalarToInt("TKT", "RecID", "dependent='" & pVcrNo & "' and status='OK'")
        If tmpRec > 0 Then Return False
        Return True
    End Function
    Private Sub LblAutoTKTOK_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAutoTKTOK.LinkClicked
        Dim RowNo As Integer, tmpTKNO As String, tmpRTG As String, MyAns As Int16
        Dim tmpCommPCT As Decimal, tmpCommVAL As Decimal
        Dim tmpROE As Decimal, tmpQty As Int16, NikeVAT As Decimal, tmpChargeD As String
        MyAns = MsgBox("Are You Sure with This Selection? REMEMBER, You Cant Undo This Action", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, msgTitle)
        If MyAns = vbNo Then Exit Sub
        isFromTKTList = True
        For r As Int16 = 0 To Me.GridAutoTKT.RowCount - 1
            If Me.GridAutoTKT.Item(0, r).Value = True Then
                If tmpCommVAL = 0 Then tmpCommVAL = Me.GridAutoTKT.Item("ROE", r).Value
                If tmpCommVAL <> Me.GridAutoTKT.Item("ROE", r).Value Then
                    MsgBox("Different ROE Applied For Selected TKTs and Cant Be Combined in 1 TRX", MsgBoxStyle.Critical, msgTitle)
                    Exit Sub
                End If
            End If
        Next
        For r As Int16 = 0 To Me.GridAutoTKT.RowCount - 1
            If Me.GridAutoTKT.Item(0, r).Value = True Then
                If tmpROE = 0 Then
                    Me.CmbCurr.Text = Me.GridAutoTKT.Item("Currency", r).Value
                    Me.LcktxtROE.Text = Me.GridAutoTKT.Item("ROE", r).Value
                    tmpROE = Me.GridAutoTKT.Item("ROE", r).Value
                    Me.LcktxtROE.Enabled = False
                    Me.CmbCurr.Enabled = False
                End If
                tmpTKNO = Me.GridAutoTKT.Item("TKNO", r).Value.ToString.Trim
                If tmpTKNO.Length = 13 Then
                    tmpTKNO = AddSpace2TKNO(tmpTKNO)
                End If
                If (MyAL.ALCode <> "01" And Not isAlreadyInGrid(tmpTKNO)) Or (MyAL.ALCode = "01" And NewInsurance(tmpTKNO)) Then
                    Me.GridTKT.Rows.Add()
                    RowNo = Me.GridTKT.Rows.Count - 1
                    If MyAL.ALCode = "01" Then
                        tmpRTG = Me.GridAutoTKT.Item("TKNO", r).Value.trim
                        If tmpRTG.Substring(0, 4) = "WNVT" Or tmpRTG.Length = 8 Then
                            Try
                                tmpCommVAL = CDec(tmpRTG.Substring(4))
                            Catch ex As Exception
                                tmpRTG = "840" & tmpRTG.Substring(4)
                            End Try
                            tmpTKNO = GenPseudoTKT("INS", "TV")
                            tmpTKNO = AddSpace2TKNO(tmpTKNO)
                        End If
                    ElseIf Me.GridAutoTKT.Item("DocType", r).Value <> "MCO" Then
                        tmpRTG = Me.GridAutoTKT.Item("Itinerary", r).Value
                        If InStr(tmpRTG, " ") = 0 Then
                            tmpRTG = AddSpace2Rtg(tmpRTG)
                        End If
                    Else
                        tmpRTG = ""
                    End If
                    tmpCommVAL = Me.GridAutoTKT.Item("Fare", r).Value - Me.GridAutoTKT.Item("NetToAL", r).Value
                    If Me.GridAutoTKT.Item("Fare", r).Value > 0 Then
                        tmpCommPCT = tmpCommVAL / Me.GridAutoTKT.Item("Fare", r).Value * 100
                    Else
                        tmpCommPCT = 0
                    End If
                    Me.GridTKT.Item("RowID", RowNo).Value = RowNo
                    Me.GridTKT.Item("TKNO", RowNo).Value = tmpTKNO
                    If Me.GridAutoTKT.Item("FTKT", r).Value.ToString.Length = 4 Then
                        Me.GridTKT.Item("FTKT", RowNo).Value = Me.GridAutoTKT.Item("FTKT", r).Value.ToString.Substring(0, 3)
                    Else
                        Me.GridTKT.Item("FTKT", RowNo).Value = Me.GridAutoTKT.Item("FTKT", r).Value
                    End If
                    tmpQty = Me.GridAutoTKT.Item("Qty", r).Value
                    If Me.GridAutoTKT.Item("SRV", r).Value = "V" Then tmpQty = 0

                    Me.GridTKT.Item("FTKT", RowNo).Value = Me.GridAutoTKT.Item("FTKT", r).Value
                    Me.GridTKT.Item("SRV", RowNo).Value = Me.GridAutoTKT.Item("SRV", r).Value
                    Me.GridTKT.Item("Qty", RowNo).Value = tmpQty
                    Me.GridTKT.Item("CommOffer", RowNo).Value = "0"
                    Me.GridTKT.Item("DOI", RowNo).Value = Me.GridAutoTKT.Item("DOI", r).Value
                    Me.GridTKT.Item("DocType", RowNo).Value = Me.GridAutoTKT.Item("DocType", r).Value
                    Me.GridTKT.Item("CustType", RowNo).Value = Me.CmbCustType.Text
                    Me.GridTKT.Item("DOF", RowNo).Value = Me.GridAutoTKT.Item("DOF", r).Value
                    Me.GridTKT.Item("Routing", RowNo).Value = tmpRTG
                    If MyAL.ALCode <> "01" Then
                        Me.GridTKT.Item("BkgClass", RowNo).Value = Me.GridAutoTKT.Item("BkgClass", r).Value
                        Me.GridTKT.Item("PaxType", RowNo).Value = Me.GridAutoTKT.Item("PaxType", r).Value
                    Else
                        If tmpRTG.Length = 8 Or tmpRTG.Substring(0, 4) = "WNVT" Or (MyAL.ALCode = "01" AndAlso tmpRTG.Substring(0, 3) = "840") Then
                            Me.GridTKT.Item("BkgClass", RowNo).Value = "1"
                            Me.GridTKT.Item("DocType", RowNo).Value = "INS"
                            Me.GridTKT.Item("PaxType", RowNo).Value = "TGD"
                        End If
                    End If
                    Me.GridTKT.Item("FareBasis", RowNo).Value = Me.GridAutoTKT.Item("FareBasis", r).Value
                    Me.GridTKT.Item("StockCtrl", RowNo).Value = Me.GridAutoTKT.Item("StockCtrl", r).Value
                    Me.GridTKT.Item("PaxName", RowNo).Value = Me.GridAutoTKT.Item("PaxName", r).Value
                    Me.GridTKT.Item("TourCode", RowNo).Value = Me.GridAutoTKT.Item("TourCode", r).Value
                    Me.GridTKT.Item("PromoCode", RowNo).Value = "..."
                    Me.GridTKT.Item("FareType", RowNo).Value = "..."
                    Me.GridTKT.Item("Fare", RowNo).Value = Me.GridAutoTKT.Item("Fare", r).Value * tmpQty
                    Me.GridTKT.Item("ShownFare", RowNo).Value = Me.GridAutoTKT.Item("ShownFare", r).Value * tmpQty
                    Me.GridTKT.Item("LLF", RowNo).Value = Me.GridAutoTKT.Item("LLF", r).Value * tmpQty
                    Me.GridTKT.Item("NetToAL", RowNo).Value = Me.GridAutoTKT.Item("NetToAL", r).Value * tmpQty
                    Me.GridTKT.Item("AgtDisctPCT", RowNo).Value = 0
                    Me.GridTKT.Item("AgtDisctVAL", RowNo).Value = 0
                    Me.GridTKT.Item("CommVAL", RowNo).Value = 0
                    Me.GridTKT.Item("CommPCT", RowNo).Value = 0
                    Me.GridTKT.Item("Tax", RowNo).Value = Me.GridAutoTKT.Item("Tax", r).Value * tmpQty
                    Me.GridTKT.Item("Charge_D", RowNo).Value = ""
                    If Me.GridAutoTKT.Item("Booker", r).Value = "" Then
                        Me.GridTKT.Item("TKT_RMK", RowNo).Value = ""
                    ElseIf Me.GridAutoTKT.Item("Booker", r).Value.ToString.Contains("ZPERSONAL") Then
                        Me.GridTKT.Item("TKT_RMK", RowNo).Value = "BIZF|" & Me.GridAutoTKT.Item("Booker", r).Value
                        Me.CmbBooker.Text = Me.GridAutoTKT.Item("Booker", r).Value
                    Else
                        Me.GridTKT.Item("TKT_RMK", RowNo).Value = "BIZT|" & Me.GridAutoTKT.Item("Booker", r).Value
                        Me.CmbBooker.Text = Me.GridAutoTKT.Item("Booker", r).Value
                    End If
                    NikeVAT = 0
                    If tmpQty <> 0 Then
                        Me.GridTKT.Item("Tax_D", RowNo).Value = Me.GridAutoTKT.Item("TaxDetail", r).Value
                        If Me.GridAutoTKT.Item("Svcfee", r).Value <> 0 Then
                            Me.GridTKT.Item("Charge_D", RowNo).Value = "TVN-OTHER:" & Me.GridAutoTKT.Item("svcfeeCur", r).Value & Me.GridAutoTKT.Item("Svcfee", r).Value
                        End If
                        tmpChargeD = Me.GridAutoTKT.Item("ChargeDetail", r).Value.ToString.Replace(" ", "")
                        If String.IsNullOrEmpty(tmpChargeD) Then tmpChargeD = ""
                        If tmpChargeD.Contains("TVR-VAT") Then
                            NikeVAT = CDec(tmpChargeD.Substring(11))
                            Me.GridTKT.Item("Charge_D", RowNo).Value = Me.GridTKT.Item("Charge_D", RowNo).Value & "|" & tmpChargeD
                        End If
                        If Me.GridAutoTKT.Item("ChargeDetail", r).Value = "" Then
                            If Me.GridAutoTKT.Item("Charge", r).Value <> 0 Then
                                Me.GridTKT.Item("Charge_D", RowNo).Value = "SPN-OTHER:" & Me.GridAutoTKT.Item("Currency", r).Value & _
                                    Me.GridAutoTKT.Item("Charge", r).Value & "|"
                            End If
                        End If
                    Else
                        Me.GridTKT.Item("Tax_D", RowNo).Value = "YQ00"
                    End If
                    Me.GridTKT.Item("Charge_D", RowNo).Value = tmpChargeD
                    If Me.GridTKT.Item("Charge_D", RowNo).Value = "" Then Me.GridTKT.Item("Charge_D", RowNo).Value = "TVN-OTHER:VND0.00"
                    Me.GridTKT.Item("Charge", RowNo).Value = Me.GridAutoTKT.Item("Charge", r).Value * tmpQty
                    If String.IsNullOrEmpty(Me.GridAutoTKT.Item("SvcFeeCur", r).Value) Then
                        Me.GridTKT.Item("ChargeTV", RowNo).Value = (Me.GridAutoTKT.Item("Svcfee", r).Value + NikeVAT) * tmpQty
                    ElseIf Me.GridAutoTKT.Item("Currency", r).Value = Me.GridAutoTKT.Item("SvcFeeCur", r).Value Then
                        Me.GridTKT.Item("ChargeTV", RowNo).Value = (Me.GridAutoTKT.Item("Svcfee", r).Value + NikeVAT) * tmpQty
                    ElseIf Me.GridAutoTKT.Item("SvcFeeCur", r).Value = "VND" Then
                        Me.GridTKT.Item("ChargeTV", RowNo).Value = (Me.GridAutoTKT.Item("Svcfee", r).Value / Me.GridAutoTKT.Item("ROE", r).Value + NikeVAT) * tmpQty
                    ElseIf Me.GridAutoTKT.Item("SvcFeeCur", r).Value <> "VND" Then
                        Me.GridTKT.Item("ChargeTV", RowNo).Value = (Me.GridAutoTKT.Item("Svcfee", r).Value * Me.GridAutoTKT.Item("ROE", r).Value + NikeVAT) * tmpQty
                    End If
                    Me.GridTKT.Item("isEdit", RowNo).Value = ""
                    Me.GridTKT.Item("AmountToAL", RowNo).Value = 0
                End If
            End If
        Next
        Me.PnlAutoTKT.Visible = False
        Me.MenuStrip2.Enabled = Not Me.PnlAutoTKT.Visible
        refreshTKTcontent(0)
        If Me.GridTKT.RowCount > 0 Then
            Me.GridTKT.Item("TKNO", 0).Selected = True
        End If
    End Sub
    Private Function XacDinhAmountHold() As Decimal
        Dim tmpTKList As String, KQ As Decimal
        tmpTKList = ListOfTKTinRCP("INCLAUSE")
        If tmpTKList = "()" Then Return 0
        If tmpTKList <> "" Then
            KQ = ScalarToDec("TKT_1A", "isNull(Sum(CreditAmt),0)", "SRV='S' and TKNO in " & tmpTKList)
        End If
        Return KQ
    End Function
    Private Sub genDocType()
        Dim strSQL As String
        strSQL = "select distinct VAL from MISC where city='" & MySession.City & "' and CAT='DOCTYPE' "
        strSQL = strSQL & " and (VAL1 like '%YY%' or VAL1 like '%" & Me.CmbAL.Text & "%' or VAL1 like '%" & MySession.TRXCode & "%')"
        If Me.CmbAL.Text = "01" Then
            strSQL = strSQL & " and VAL in " & NonAirList
            If InStr("CWT_TVS", myStaff.Counter) > 0 Then
                strSQL = strSQL & " and VAL2 like '%" & myStaff.Counter & "%'"
            End If
        End If
        LoadCmb_MSC(Me.CmbDocType, strSQL)
        Me.CmbDocType.Text = "ETK"
    End Sub

    Private Sub CmbAL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbAL.LostFocus
        GenFOP_CURR()
        MyAL.ALCode = Me.CmbAL.Text
        Me.StatusCounter.Text = " Counter: " & MySession.Domain & "/" & MyAL.ALCode
    End Sub

    Private Sub CmbAL_MouseClick(sender As Object, e As MouseEventArgs) Handles CmbAL.MouseClick
        Me.OptServicing.Enabled = True
        Me.OptSell.Enabled = True
        Me.OptRefund.Enabled = True
        Me.OptVoid.Enabled = True
        Me.OptPseudoVoid.Enabled = True
        Me.OptReIss.Enabled = True
    End Sub

    Private Sub CmAL_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbAL.SelectedIndexChanged
        Dim needServicing As String
        MyAL.Domain = MySession.Domain
        MyAL.ALCode = Me.CmbAL.Text
        If Me.CmbAL.Text = "" Then
            Me.OptServicing.Enabled = False
            Me.OptSell.Enabled = False
            Me.OptRefund.Enabled = False
            Me.OptVoid.Enabled = False
            Me.OptPseudoVoid.Enabled = False
            Me.OptRefundOLD.Enabled = False
            Me.OptReIss.Enabled = False
            Exit Sub
        End If
        On Error Resume Next
        GenerateComboValue(Me)
        If MyAL.CanIssVAT Then
            Me.OptPrintTRX.Checked = False
            Me.OptPrintInv.Enabled = True
            Me.OptPrintTRX.Enabled = True
        Else
            Me.OptPrintTRX.Checked = True
            Me.OptPrintInv.Enabled = False
            Me.OptPrintTRX.Enabled = False

        End If
        Me.CmbProduct.Text = "NET"
        If MySession.Domain = "GSA" Then
            MySession.TRXCode = Me.CmbAL.Text
            Me.CmbProduct.Text = "FIT"
        End If
        DefauSF = DefineDefauSF(Me.CmbAL.Text)
        Me.StatusCounter.Text = " Counter: " & MySession.Domain & "/" & MyAL.ALCode
        Me.LcktxtROE.Text = Format(ForEX_12(Me.TxtDOS.Value, "USD", "BSR", MySession.TRXCode), "#,##0.00")
        On Error GoTo 0
        If myStaff.SupOf <> "" Then
            Me.OptRefundOLD.Enabled = True
        End If
    End Sub
    Private Sub OptSell_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        OptSell.CheckedChanged, OptRefund.CheckedChanged, OptPseudoVoid.CheckedChanged, _
        OptRefundOLD.CheckedChanged, OptServicing.CheckedChanged, OptVoid.CheckedChanged
        If mRCP = "" Then
            Me.TxtTRXNO.Text = GenRCPNo(MySession.TRXCode, MySession.POSCode)
            If Me.TxtTRXNO.Text = "" Or Me.TxtTRXNO.Text.Substring(0, 2) = "YY" Or pubVarRCPID_BeingCreated = 0 Then
                MsgBox("Cant Generate TRX No. Please Contact System Admin for Help", MsgBoxStyle.Critical, msgTitle)
                Me.Close()
                Exit Sub
            End If
        End If
        genDocType()
    End Sub
    Private Sub SetDefaultFOP()
        If MyCust.DelayType = "PPD" Or MyCust.DelayType = "PSP" Then
            Me.GridFOP.Item("FOP_FOP", 0).Value = MyCust.DelayType
            Me.GridFOP.Item("Amount", 0).Value = 0
            Me.GridFOP.Item("RCPCurrency", 0).ReadOnly = True
            If MyCust.DelayType = "PPD" Then
                Me.GridFOP.Item("RCPCurrency", 0).Value = "VND"
                Me.GridFOP.Item("RCPROE", 0).Value = 1
            ElseIf MyCust.DelayType = "PSP" Then
                Me.GridFOP.Item("RCPCurrency", 0).Value = Me.CmbCurr.Text
                Me.GridFOP.Item("RCPROE", 0).Value = CDec(Me.LcktxtROE.Text)
            End If
        End If
        Me.CmdSave.Enabled = False
    End Sub

    Private Sub BarPendingRQ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarPendingRQ.Click, BarALLPendingRQ.Click
        Dim bBar As ToolStripItem = CType(sender, ToolStripItem)
        Dim strSQL As String
        Dim btag As String = bBar.Tag
        Dim DKien As String
        Dim UptoDate As String = Format(Now.Date.AddDays(-8), "dd-MMM-yy")
        If Me.GridPendingRQ.Visible Then Exit Sub
        If MySession.Domain = "EDU" Then
            DKien = " t.custid <0 "
        Else
            DKien = " t.custid >0 "
        End If
        If btag <> "ALL" Then
            DKien = DKien & " and t.custid in (select custId from cust_Detail where status+Cat='OKChannel' and "
            DKien = DKien & " VAL in " & myStaff.CAccess & ") "
        End If

        strSQL = "select distinct t.CustID, a.AL, SRV, RLOC from TKT_1a t inner join airline a on a.doccode=left(t.tkno,3) "
        strSQL = strSQL & " where tkno <>'' and " & DKien

        strSQL = strSQL & "and t.recid in (select recID from funcTKT_1AnotUpdated2RAS_V ('" & UptoDate & "'))" & _
            " UNION  " & strSQL & " and t.recid in (select recID from funcTKT_1AnotUpdated2RAS_wzparam ('" & UptoDate & "'))"

        Me.GridPendingRQ.DataSource = GetDataTable(strSQL)
        Me.GridPendingRQ.Columns("CustID").Width = 48
        Me.GridPendingRQ.Columns("AL").Width = 32
        Me.GridPendingRQ.Columns("SRV").Width = 32
        Me.GridPendingRQ.Columns("RLOC").Width = 64
        Me.GridPendingRQ.Visible = True
        Me.GridPendingRQ.Top = 85
        Me.GridPendingRQ.Left = 6
        Me.GridPendingRQ.Width = 200
        Me.GridPendingRQ.Height = 256
    End Sub
    Private Sub GridPendingRQ_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridPendingRQ.CellContentDoubleClick
        Dim tmpCustChannel As String
        Me.GridPendingRQ.Visible = False
        Me.CmbAL.Text = Me.GridPendingRQ.CurrentRow.Cells("AL").Value
        If Me.GridPendingRQ.CurrentRow.Cells("SRV").Value = "S" Then
            Me.OptSell.PerformClick()
        ElseIf Me.GridPendingRQ.CurrentRow.Cells("SRV").Value = "V" Then
            MsgBox("Sorry, cant support Void Ticket. Plz start transaction as normal", MsgBoxStyle.Information, msgTitle)
            Exit Sub
        End If
        tmpCustChannel = ScalarToString("cust_Detail", "VAL", "status+Cat='OKChannel' and Custid=" & Me.GridPendingRQ.CurrentRow.Cells("CustID").Value)
        Me.CmbCustType.Text = tmpCustChannel
        MyAL.ALCode = Me.CmbAL.Text
        Me.CmbAL.Focus()
        Me.OptSell.Enabled = True
        Me.OptSell.PerformClick()
        Me.GrpNewRCP.Visible = True
        Me.CmbCustType.Focus()
        Me.CmbChannel.Focus()
        Me.CmbChannel.SelectedValue = Me.GridPendingRQ.CurrentRow.Cells("CustID").Value
    End Sub
    Private Sub CalServiceFee_CWT(ByVal pDK_Status_Validity As String)
        Dim regCountries As String, rtgType As String = "", SF As Decimal, curr As String
        Dim Svc As String = IIf(InStr("AHC_INS_HTL_CAR_VSA", Me.CmbDocType.Text) > 0, Me.CmbDocType.Text, "AIR")
        Dim SVCFeeID As Integer, strDK As String = "custID=" & MyCust.CustID & " and SvcType='" & Svc & "'"
        strDK = strDK & " and TRX='" & pubVarSRV & "'"
        SVCFeeID = ScalarToInt("CWT_SF", "recID", strDK)
        If SVCFeeID = 0 Then Exit Sub ' ko dc khai SF cho cust nay
        If Svc = "AIR" Then
            regCountries = ScalarToString("CWT_SF", "Countries", " recid=" & SVCFeeID)
            rtgType = DefineRTGType_SF_CWT(regCountries)
        End If
        curr = ScalarToString("CWT_SF", "Curr", " recid=" & SVCFeeID)
        SF = ScalarToDec("CWT_SF", rtgType, " recid=" & SVCFeeID)
        Me.CmbCharge1.Text = "TVN-AUTOSF"
        If curr = Me.CmbCurr.Text Then
            Me.TxtCharge1.Text = SF
        ElseIf curr = "VND" Then
            Me.TxtCharge1.Text = SF / CDec(Me.LcktxtROE.Text)
        ElseIf Me.CmbCurr.Text = "VND" Then
            Me.TxtCharge1.Text = SF * CDec(Me.LcktxtROE.Text)
        Else
            Me.TxtCharge1.Text = SF
        End If
    End Sub
    Private Sub CalServiceFee(ByVal pDK_Status_Validity As String)
        Dim pDK_AL_SBU As String = " and AL+SBU='" & MyAL.ALCode & MySession.Domain & "'"
        Dim dTbl As DataTable
        Dim tmpSFAmt As Decimal, tmpSFName As String, strSQL As String
        Dim tmpFareBase As Decimal, tmpSRV As String = pubVarSRV, ISI As String, RTGType As String

        If InStr("OC", tmpSRV) > 0 Then tmpSRV = "R"
        If InStr("A", tmpSRV) > 0 Or Me.CmbDocType.Text = "EXC" Then tmpSRV = "I"
        tmpFareBase = Me.txtFare.Text
        RTGType = DefineRTGType_SF()
        ISI = CityAPTToCountry_Area_City("Country", Me.txtRTG.Text.Substring(0, 3))
        ISI = IIf(ISI = "VN", "SIT*", "SOT*")

        strSQL = "select top 1 * from serviceFee "
        strSQL = strSQL & " where TRXType='" & tmpSRV & "' and Area in ('ALL','" & RTGType & "') "
        strSQL = strSQL & " and ISI in ('ALL','" & ISI & "') and " & tmpFareBase
        strSQL = strSQL & " between fareFrom and fareThrough and channel='" & Me.CmbCustType.Text & "'"
        strSQL = strSQL & pDK_AL_SBU & pDK_Status_Validity
        If Me.CmbCustType.Text <> "WK" Then
            strSQL = strSQL & " and Channel+CustLevel in (select Channel+CustLevel from cust_channel_level "
            strSQL = strSQL & " where CustID=" & Me.CmbChannel.SelectedValue
            strSQL = strSQL & pDK_AL_SBU & pDK_Status_Validity & ")"
        End If
        strSQL = strSQL & " and G_FIT  ='" & Me.CmbGRPFIT.Text & "' and FID_Fier='" & Me.CmbFSource.Text & "'"
        If Me.CmbPaxType.Text = "INF" Then strSQL = strSQL & " and INF <> 0"
        strSQL = strSQL & " order by Area desc, ISI Desc "
        dTbl = GetDataTable(strSQL)

        If dTbl.Rows.Count > 0 Then
            tmpSFAmt = 0
            tmpSFName = "TVN-OTHER"
            For i As Int16 = 0 To dTbl.Rows.Count - 1
                tmpSFName = "TV" & IIf(dTbl.Rows(i)("Refundable") = 0, "N", "R") & "-AUTOSF"
                tmpSFAmt = dTbl.Rows(i)("VAL")
                If dTbl.Rows(i)("SFType") = "PCT" Then
                    tmpSFAmt = tmpSFAmt * tmpFareBase / 100
                    If tmpSFAmt < dTbl.Rows(i)("MinVal") Then
                        tmpSFAmt = dTbl.Rows(i)("MinVal")
                    ElseIf tmpSFAmt > dTbl.Rows(i)("MaxVal") Then
                        tmpSFAmt = dTbl.Rows(i)("MaxVal")
                    End If
                End If
            Next
            Me.CmbCharge1.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCharge1.Enabled = IIf(tmpSFAmt <> 0, False, True)
            Me.CmbCharge1.Text = tmpSFName
            Me.TxtCharge1.Text = tmpSFAmt
        Else
            ManualSF(DefauSF, "TVN-MANLSF")
        End If

    End Sub
    Private Function DefineRTGType_SF_CWT(ByVal pCountry As String) As String
        Dim KQ As String = "DOM"
        Dim tmpRTG As String, leg As Int16, CT As String, tmpCountry As String
        tmpRTG = Me.txtRTG.Text
        tmpRTG = tmpRTG.Replace(" ", "")
        leg = (tmpRTG.Length - 3) / 5
        For i As Int16 = 0 To leg
            CT = tmpRTG.Substring(5 * i, 3)
            tmpCountry = CityAPTToCountry_Area_City("Country", CT)
            If tmpCountry <> "VN" Then
                KQ = "REG"
                If InStr(pCountry, tmpCountry) = 0 Then
                    KQ = "INTL"
                    Exit For
                End If
            End If
        Next
        Return KQ
    End Function

    Private Function DefineRTGType_SF() As String
        Dim KQ As String = "DOM"
        Dim tmpRTG As String, leg As Int16, CT As String
        tmpRTG = Me.txtRTG.Text
        tmpRTG = tmpRTG.Replace(" ", "")
        leg = (tmpRTG.Length - 3) / 5
        For i As Int16 = 0 To leg
            CT = tmpRTG.Substring(5 * i, 3)
            If CityAPTToCountry_Area_City("Country", CT) <> MyAL.QuocTich Then
                KQ = "INT"
                Exit For
            End If
        Next
        Return KQ
    End Function
    Private Sub ManualSF(ByVal pAmt As Int16, ByVal pSFName As String)
        Me.CmbCharge1.DropDownStyle = ComboBoxStyle.DropDown
        Me.CmbCharge1.Enabled = True
        Me.TxtCharge1.Enabled = True
        Me.CmbCharge1.Text = pSFName
        Me.TxtCharge1.Text = pAmt
    End Sub

    Private Sub OptReIss_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptReIss.CheckedChanged
        Me.CmbDocType.Text = "EXC"
        Me.CmbDocType.Enabled = False
        Me.OptSell.PerformClick()
    End Sub
    Private Sub CalcRefundableCharges()
        If pubVarSRV <> "R" Then Exit Sub
        Dim TVR As Decimal = 0, ALR As Decimal = 0
        Dim tmpChargeD As String, tmpName As String, tmpAmt As Decimal
        For j As Int16 = 0 To Me.GridTKT.RowCount - 1
            tmpChargeD = Me.GridTKT.Item("Charge_D", j).Value
            If Me.GridTKT.Item("qty", j).Value = 1 Then
                For i As Int16 = 0 To 3
                    tmpAmt = tmpChargeD.Split("|")(i).Split(":")(1)
                    If tmpAmt > 0 Then
                        tmpName = tmpChargeD.Split("|")(i).Split(":")(0).Substring(0, 3)
                        If tmpName = "TVR" Then
                            TVR = TVR + tmpAmt
                        ElseIf tmpName = "ALR" Then
                            ALR = ALR + tmpAmt
                        End If
                    End If
                Next
            End If
        Next
        Me.TxtALR_Ch.Text = ALR
        If isFullRefund() Then
            Me.TxtTVR_Ch.Text = TVR
        End If
    End Sub
    Private Function isFullRefund() As Boolean
        Dim KQ As Boolean = True
        If Me.txtRTG.Text <> SalesInforOfThisTKT.Rtg Then
            KQ = False
        End If
        Return KQ
    End Function

    Private Sub CmbGRPFIT_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        CmbGRPFIT.SelectedIndexChanged, CmbFSource.SelectedIndexChanged
        CalServiceFee(DK_Status_Validity)
    End Sub
    Private Function getSQL_GridTKT_to_TableTKT(ByVal varRCPID As Integer, ByVal varSRV As String) As String
        Dim SL As Decimal, LclSRV As String, strSQL As String, Dependent As String = "", RTG As String = ""
        Dim ALStatus As String = "OK", LclTKNO As String
        SL = 1
        LclSRV = varSRV
        strSQL = ""
        Dependent = ""
        For r As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            If InStr(NonAirList, Me.GridTKT.Item("DocType", r).Value) > 0 Then
                If mRCP = "" Then
                    Dependent = Me.GridTKT.Item("Routing", r).Value
                Else
                    Dependent = ScalarToString("TKT", "top 1 [dependent]", " rcpid=" & pubVarRCPID_BeingEdited)
                End If
                If Me.GridTKT.Item("DocType", r).Value = "HTL" Then
                    RTG = Me.GridTKT.Item("PaxType", r).Value & "x" & Me.GridTKT.Item("bkgClass", r).Value.ToString.Split("/")(1) & "N @"
                    RTG = RTG & Me.GridTKT.Item("FareBasis", r).Value & " CKIN" & Format(Me.GridTKT.Item("DOF", r).Value, "ddMMM")
                ElseIf Me.GridTKT.Item("DocType", r).Value = "VSA" Then
                    RTG = Me.GridTKT.Item("PaxType", r).Value & " " & Me.GridTKT.Item("TourCode", r).Value & " to " & Me.GridTKT.Item("FareBasis", r).Value
                ElseIf Me.GridTKT.Item("DocType", r).Value = "TVS" Then
                    RTG = Me.GridTKT.Item("PaxType", r).Value
                End If
            Else
                RTG = Me.GridTKT.Item("Routing", r).Value
            End If
            If varSRV = "V" Then
                SL = 0
            ElseIf Not Me.GridTKT.Item("FTKT", r).Value Is Nothing AndAlso _
                Me.GridTKT.Item("FTKT", r).Value <> "" AndAlso _
                Me.GridTKT.Item("FTKT", r).Value.ToString.Substring(0, 3) <> "___" Then
                SL = 0
            ElseIf InStr("O_R_C", varSRV) > 0 Or Me.GridTKT.Item("DocType", r).Value = "ACM" Then
                SL = -1
                LclSRV = "R"
                If varSRV = "C" Then ALStatus = "XX"
            Else
                SL = 1
            End If
            LclTKNO = Me.GridTKT.Item("TKNO", r).Value.ToString.Trim
            If InStr(LclTKNO, " ") = 0 Then
                LclTKNO = AddSpace2TKNO(LclTKNO)
            End If
            strSQL = strSQL & "; insert into TKT "
            strSQL = strSQL & "(RCPID, RCPNO, TKNO, FTKT, SRV, DOI, DOF, Itinerary, BkgClass, Currency,"
            strSQL = strSQL & "FareBasis, DocType, PaxType, TourCode, PromoCode, FareType, Fare, ShownFare, LLF,"
            strSQL = strSQL & "Tax, Charge, ChargeTV, CommPCT, CommVAL, NettoAL, Qty, AgtDisctPCT, CommOffer,"
            strSQL = strSQL & "Tax_D, Charge_D, AgtDisctVAL, FstUser, PaxName, AL, StatusAL, StockCtrl, RMK, [Dependent]) "
            strSQL = strSQL & " values ('"
            strSQL = strSQL & varRCPID
            strSQL = strSQL & "','" & Me.TxtTRXNO.Text
            strSQL = strSQL & "','" & LclTKNO
            strSQL = strSQL & "','" & Me.GridTKT.Item("FTKT", r).Value
            strSQL = strSQL & "','" & LclSRV
            strSQL = strSQL & "','" & Me.GridTKT.Item("DOI", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("DOF", r).Value
            strSQL = strSQL & "','" & RTG.Replace("--", "")
            strSQL = strSQL & "','" & Me.GridTKT.Item("BkgClass", r).Value.ToString.Replace("--", "")
            strSQL = strSQL & "','" & Me.CmbCurr.Text
            strSQL = strSQL & "','" & Me.GridTKT.Item("FareBasis", r).Value.ToString.Replace("--", "")
            strSQL = strSQL & "','" & Me.GridTKT.Item("DocType", r).Value
            If SL = 0 Then
                strSQL = strSQL & "','" & ""
            Else
                strSQL = strSQL & "','" & Me.GridTKT.Item("PaxType", r).Value.ToString.Substring(0, 3)
            End If
            strSQL = strSQL & "','" & Me.GridTKT.Item("TourCode", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("PromoCode", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("FareType", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("Fare", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("ShownFare", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("LLF", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("Tax", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("Charge", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("ChargeTV", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("CommPCT", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("CommVAL", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("NetToAL", r).Value
            strSQL = strSQL & "','" & SL
            strSQL = strSQL & "','" & Me.GridTKT.Item("AgtDisctPCT", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("CommOffer", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("Tax_D", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("Charge_D", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("AgtDisctVAL", r).Value
            strSQL = strSQL & "','" & myStaff.SICode
            strSQL = strSQL & "','" & Me.GridTKT.Item("PaxName", r).Value.ToString.Replace("--", "")
            strSQL = strSQL & "','" & Me.CmbAL.Text
            strSQL = strSQL & "','" & ALStatus
            strSQL = strSQL & "','" & Me.GridTKT.Item("StockCtrl", r).Value
            strSQL = strSQL & "','" & Me.GridTKT.Item("TKT_RMK", r).Value
            strSQL = strSQL & "','" & Dependent & "')"
        Next
        If varSRV = "C" Then ' C= Voi khach thi Refund nhung bao cao hang la Void
            For r As Int16 = 0 To Me.GridTKT.Rows.Count - 1
                strSQL = strSQL & "; insert into TKT "
                strSQL = strSQL & "(RCPID, RCPNO, TKNO, FTKT, DOI, Currency, DocType, Charge, "
                strSQL = strSQL & "Charge_D, AL, FstUser, StatusAL, Status, Qty, SRV) values ('"
                strSQL = strSQL & varRCPID
                strSQL = strSQL & "','" & Me.TxtTRXNO.Text
                strSQL = strSQL & "','" & Me.GridTKT.Item("TKNO", r).Value.ToString.Trim
                strSQL = strSQL & "','" & Me.GridTKT.Item("FTKT", r).Value
                strSQL = strSQL & "','" & Me.GridTKT.Item("DOI", r).Value
                strSQL = strSQL & "','" & Me.CmbCurr.Text
                strSQL = strSQL & "','" & Me.GridTKT.Item("DocType", r).Value
                strSQL = strSQL & "','" & Me.GridTKT.Item("Charge", r).Value
                strSQL = strSQL & "','" & Me.GridTKT.Item("Charge_D", r).Value
                strSQL = strSQL & "','" & Me.CmbAL.Text
                strSQL = strSQL & "','" & myStaff.SICode & "','OK','XX',0,'V')"
                strSQL = strSQL & "; update TKT set statusAL='XX' where srv='S' and statusal='OK' and tkno='" & _
                    Me.GridTKT.Item("TKNO", r).Value.ToString.Trim & "'"
            Next
        End If
        Return strSQL.Substring(1)
    End Function
    Private Function TKTAlreadyExist(ByVal varRCPID As Integer, ByVal varSRV As String) As Boolean
        Dim strSQL As String, tmpTKNO As String = ""
        If InStr("SVO", varSRV) > 0 Then ' Check ve trung do xung dot
            For r As Int16 = 0 To Me.GridTKT.Rows.Count - 1
                strSQL = " TKNO='" & Me.GridTKT.Item("TKNO", r).Value.ToString.Trim & "' and srv in ('S','R','V') and (status='OK' or statusAL ='OK') and left(RCPNO,2)='" & MySession.TRXCode & "'"
                tmpTKNO = ScalarToString("TKT", "TKNO", strSQL)
                If tmpTKNO <> "" Then
                    MsgBox(Me.GridTKT.Item("TKNO", r).Value.ToString.Trim & " Already Exists. Please Check Your Input")
                    Return True
                End If
            Next
        End If
        Return False
    End Function
    Private Function LogicCheckOK() As Boolean

        If pubVarSRV = "V" Then Return True
        If InStr(NonAirList, Me.CmbDocType.Text) = 0 AndAlso Me.txtRTG.Text.Trim.Length < 10 Then Return False
        If pubVarSRV = "R" Then
            If SalesInforOfThisTKT.Rtg.Trim.Length > 0 And Me.txtRTG.Text.Trim.Length > SalesInforOfThisTKT.Rtg.Trim.Length Then
                FeedBack("Error. Invalid Segments")
                Return False
            ElseIf SalesInforOfThisTKT.Rtg.Trim.Length > 0 And Strings.Right(SalesInforOfThisTKT.Rtg.Trim, Me.txtRTG.Text.Trim.Length) <> Me.txtRTG.Text.Trim Then
                FeedBack("Error. Invalid Segments")
                MsgBox("Invalid Segments. Please Check", MsgBoxStyle.Critical, msgTitle)
                Return True
            End If
            If Not SalesInforOfThisTKT.isEXC And (MySession.Counter = "CWT" Or MySession.Domain = "GSA") Then
                If CDec(Me.txtFare.Text) > CDec(SalesInforOfThisTKT.Fare) And myStaff.SupOf = "" Then
                    FeedBack("Error. Invalid fare Input")
                    Return False
                End If
                If CDec(Me.txtNetToAL.Text) > CDec(SalesInforOfThisTKT.NetToAL) Then
                    If myStaff.SupOf <> "" Then ' Duoc Phep Ban lost sales
                        MsgBox("Be Carefull. Lost Sales", MsgBoxStyle.Critical, msgTitle)
                    Else
                        MsgBox("Check Fare/Net2AL Input!", MsgBoxStyle.Critical, msgTitle)
                        Return False
                    End If
                End If
                If CDec(Me.txtShownFare.Text) > 0 And CDec(Me.txtShownFare.Text) < CDec(Me.txtFare.Text) Then
                    FeedBack("Error. Fare Cant Be Greater Than RefFare")
                    Return False
                End If
                If CDec(Me.txtLLF.Text) > 0 And CDec(Me.txtLLF.Text) > CDec(Me.txtFare.Text) Then
                    FeedBack("Error. LLF Cant Be Greater Than Fare")
                    Return False
                End If

                If CDec(Me.txtShownFare.Text) > 0 And CDec(Me.txtLLF.Text) > 0 And CDec(Me.txtLLF.Text) > CDec(Me.txtShownFare.Text) Then
                    FeedBack("Error. LLF Cant Be Greater Than RefFare")
                    Return False
                End If

                If CDec(Me.TxtTax.Text) > CDec(SalesInforOfThisTKT.Tax) And myStaff.SupOf = "" Then
                    MsgBox("Check Tax Input!", MsgBoxStyle.Critical, msgTitle)
                    Return False
                End If
            End If
        End If
        If pubVarSRV = "S" Then
            Dim Margin As Decimal = CDec(Me.txtFare.Text) - CDec(Me.txtNetToAL.Text)
            Dim UnderTable As Decimal = GetUnderTable()
            If Margin < 0 Then 'Lost Sales
                If Not PresetLostSales(Me.CmbChannel.SelectedValue, Me.CmbAL.Text, Margin * -1, CDec(Me.txtNetToAL.Text)) Then
                    If myStaff.SupOf <> "" Then ' Duoc Phep Ban lost sales
                        MsgBox("Be Carefull. Lost Sales", MsgBoxStyle.Critical, msgTitle)
                    Else
                        MsgBox("Check Fare/Net2AL Input!", MsgBoxStyle.Critical, msgTitle)
                        Return False
                    End If
                End If
            End If
            If UnderTable > 0 And Margin < UnderTable Then
                MsgBox("Too Much or Illogic PayOut!", MsgBoxStyle.Critical, msgTitle)
                Return False
            End If
        End If
        If Me.CmbProduct.Text.Trim = "FOC" And Me.txtNetToAL.Text > 0 Then
            MsgBox("Check Net2AL!", MsgBoxStyle.Critical, msgTitle)
            Return False
        ElseIf Me.CmbProduct.Text.Trim = "NET" And Me.txtALCommVAL.Text > 0 Then
            MsgBox("Check Comm!", MsgBoxStyle.Critical, msgTitle)
            Return False
        ElseIf Me.CmbProduct.Text.Trim = "S+M" And _
            Not CDec(Me.txtFare.Text) > CDec(Me.txtNetToAL.Text) + CDec(Me.txtALCommVAL.Text) Then
            'Selling  co MU nen PHAI du phan MU
            MsgBox("Check Fare/Net2AL/Comm!", MsgBoxStyle.Critical, msgTitle)
            Return False
        End If
        Return True
    End Function
    Private Function GetUnderTable() As Decimal
        Dim ChargeD As String = GomTax_Charge("C")
        Return CalcCharge(ChargeD, "XX", Me.CmbCurr.Text, CDec(Me.LcktxtROE.Text), Me.TxtDOS.Value.Date, MySession.TRXCode)
    End Function
    Private Sub LblSIR_Close_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSIR_Close.LinkClicked
        Me.GrpSIR.Visible = False
    End Sub

    Private Sub LblSIR_Update_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSIR_Update.LinkClicked
        Dim strSQL As String
        For i As Int16 = 0 To Me.GrpSIR.Controls.Count - 1
            If Me.GrpSIR.Controls(i).Name.Substring(0, 3).ToUpper = "CTR" Then
                If Me.GrpSIR.Controls(i).Enabled Then
                    strSQL = "insert SIR (RCPID, CustID, Prod, FstUser, FName, FValue) values ("
                    strSQL = strSQL & pubVarRCPID_BeingCreated & "," & Me.CmbChannel.SelectedValue & ",'"
                    strSQL = strSQL & Me.CmbDocType.Text & "','" & myStaff.SICode & "','"
                    strSQL = strSQL & Me.GrpSIR.Controls(i).Name.Substring(3).ToUpper & "','"
                    strSQL = strSQL & Me.GrpSIR.Controls(i).Text & "')"
                    cmd.CommandText = strSQL
                    cmd.ExecuteNonQuery()
                End If
            End If
        Next
        Me.GrpSIR.Visible = False
    End Sub
    Private Sub InitGrpSIR()

        Exit Sub
        Me.GrpSIR.Top = 2
        Me.GrpSIR.Left = 196
        Me.GrpSIR.Height = 464
        For i As Int16 = 0 To Me.GrpSIR.Controls.Count - 1
            If Me.GrpSIR.Controls(i).Name.Substring(0, 3).ToUpper = "CTR" Then
                Me.GrpSIR.Controls(i).Enabled = False
            End If
        Next
        For i As Int16 = 0 To UBound(MyCust.SIR.Split("|"))
            Me.GrpSIR.Controls("Ctr" & MyCust.SIR.Split("|")(i)).Enabled = True
        Next
    End Sub

    Private Sub BarCopyTKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarCopyTKT.Click
        RemoveOrCopy("C")
    End Sub

    Private Sub BarRemoveTKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarRemoveTKT.Click
        RemoveOrCopy("R")
    End Sub

    Private Sub BarEditTax_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarEditTax.Click
        Dim vStatus As Boolean, txt As String
        vStatus = Not Me.txtTaxLabel1.ReadOnly
        Me.txtTaxLabel1.ReadOnly = vStatus
        Me.txtTaxLabel2.ReadOnly = vStatus
        Me.TxtTaxLabel3.ReadOnly = vStatus
        Me.TxtTaxLabel4.ReadOnly = vStatus
        Me.TxtTaxLabel5.ReadOnly = vStatus
        For i As Int16 = 1 To 5
            txt = "txtTaxLabel" & Trim(Str(i))
            If vStatus = False Then
                Me.GrpTax.Controls(txt).BackColor = Color.White
            Else
                Me.GrpTax.Controls(txt).BackColor = pubVarBackColor
            End If
        Next
    End Sub
    Private Sub BarNewTKT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarNewTKT.Click
        Me.CmdAddTKT.Visible = True
        Me.txtTKNO.Text = MyAL.DocCode
        Me.txtTKNO.Enabled = True
        Me.txtTKNO.Focus()
    End Sub
    Private Sub GenFOP_CURR()
        Dim dTbl As DataTable = GetDataTable("select distinct currency from ForEx where Status='OK'")
        Me.FOP_FOP.Items.Clear()
        Me.RCPCurrency.Items.Clear()
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            Me.RCPCurrency.Items.Add(dTbl.Rows(i)("Currency"))
        Next
        dTbl = GetDataTable("select VAL from MISC where cat='FOP' and status='OK' and (VAL1 like '%YY%' or VAL1 like '%" & MySession.TRXCode & "%')")
        For i As Int16 = 0 To dTbl.Rows.Count - 1
            Me.FOP_FOP.Items.Add(dTbl.Rows(i)("VAL"))
        Next
        Me.GridFOP.Columns(2).ReadOnly = Not CanEditROE
    End Sub

    Private Sub CmbGRPFIT_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbGRPFIT.VisibleChanged
        Me.CmbFSource.Visible = Me.CmbGRPFIT.Visible
        Me.Label54.Visible = Me.CmbGRPFIT.Visible
        Me.Label55.Visible = Me.CmbGRPFIT.Visible
    End Sub

    Private Sub ShowMoreInforToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowMoreInforToolStripMenuItem.Click
        If MyCust.HasSIR Then Me.GrpSIR.Visible = True
    End Sub

    Private Sub GrpSIR_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GrpSIR.VisibleChanged
        If Me.GrpSIR.Visible = False Then Me.txtTKNO.Focus()
    End Sub

    Private Sub OptSell_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSell.EnabledChanged
        If Me.OptSell.Enabled Then GenFOP_CURR()
    End Sub
    Private Sub pStep2_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles pStep2.Leave
        If MyCust.HasSIR Then Me.GrpSIR.Visible = True
    End Sub
    Private Sub CmbHTLName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbHTLName.SelectedIndexChanged
        Me.txtFB.Text = Me.CmbHTLName.Text
    End Sub

    Private Sub CmbPaxType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbPaxType.SelectedIndexChanged
        If Me.CmbPaxType.Width > 49 Then Me.CmbPaxType.Width = 49
    End Sub

    Private Sub CmbBooker_LostFocus(sender As Object, e As EventArgs) Handles CmbBooker.LostFocus
        Me.CmbTayBa.Text = Me.CmbBooker.Text
    End Sub

    Private Sub CmbBooker_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbBooker.VisibleChanged
        Me.Label5.Visible = Me.CmbBooker.Visible
        Me.ChkBizTrip.Visible = Me.CmbBooker.Visible
    End Sub

    Private Sub TxtCountry_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCountry.Enter
        If Me.TxtCountry.Text = "VENDOR" Then
            Me.TxtCountry.Text = ""
            Me.TxtCountry.ForeColor = Color.Black
        End If

    End Sub
    Private Sub TxtCountry_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCountry.LostFocus
        Dim strSelect As String = " select citycode +'-'+HtlName as val from cwt.dbo.go_HotelList"
        Dim strDK As String = " where citycode in (select city from citycode where country='" & Me.TxtCountry.Text & "')"
        strSelect = strSelect & "tv " & strDK & " UNION " & strSelect & "CWT " & strDK
        LoadCmb_MSC(Me.CmbHTLName, strSelect)
    End Sub



    Private Sub BarViewBalance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BarViewBalance.Click
        RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "BlcPrev Refresh", Conn, myStaff.SICode, CnStr)
    End Sub
    Private Function PresetLostSales(ByVal pCustID As Integer, ByVal pAL As String, ByVal pDiff As Double, ByVal pNAL As Double) As Boolean
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "select details from Misc where cat+Status='LOSTSALESOK' and getdate() between fstupdate and lstUpdate " & _
            " and VAL='" & pCustID & "' and val1='" & pAL & "'"
        Dim tmpKQ As String = cmd.ExecuteScalar
        If String.IsNullOrEmpty(tmpKQ) Then
            Return False
        ElseIf CDec(tmpKQ) > 0 Then
            If (CDec(tmpKQ) < 1 And pDiff / pNAL < CDec(tmpKQ)) Or _
                (CDec(tmpKQ) > 1 And pDiff < CDec(tmpKQ)) Then ' 
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function
    Private Sub CmbChannel_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbChannel.SelectedValueChanged
        If MyCust.CustType Is Nothing Then Exit Sub
        If InStr("CS_LC", MyCust.CustType) > 0 Then
            Me.txtShownFare.Enabled = True
            LoadCmbBooker(MyCust.CustID)
        ElseIf MySession.Domain = "GSA" And InStr("GA_PG", MyAL.ALCode) > 0 Then
            Me.Label6.Text = "ShownFare"
            Me.txtShownFare.Enabled = True
        End If
    End Sub

    Private Sub CmbCustType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCustType.SelectedIndexChanged

    End Sub
End Class