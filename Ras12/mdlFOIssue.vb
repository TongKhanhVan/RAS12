Module MdlFOIssue
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Function definePubVarNeedServicing() As String
        Dim KQ As String
        KQ = ScalarToString("MISC", "VAL", "cat='SERVICING'")
        Return KQ
    End Function
    Public Function VeCongNoDaInv(ByVal pTRXID As Integer) As Boolean
        Dim KQ As Boolean = False, tmpRecID As Integer
        tmpRecID = ScalarToInt("FOP", "RecID", "RCPID=" & pTRXID & " and status <>'XX' and FOP='PSP' and profID>0")
        If tmpRecID > 0 Then KQ = True
        Return KQ
    End Function
    Public Function GetApproval(ByVal parRCP As String, ByVal parType As String, ByVal parCustID As Integer, ByVal pTKNo As String) As Decimal
        Dim KQ As Decimal = 0, strDK As String = "where tableName+doWhat='CREXTOK' and f6 in ('" & parType & "','MIX')"
        strDK = strDK & " and (F1='" & parRCP & "' OR F4 like '%" & pTKNo & "%' )"
        strDK = strDK & " and F2='" & parCustID.ToString.Trim & "' order by recID desc"
        KQ = ScalarToDec("ActionLog", "top 1 F5", strDK)
        Return KQ
    End Function
    Public Sub LoadTaxCodeByAl(ByVal ParFOissueTKT As FOissueTKT)
        Dim dTable As DataTable
        Dim TaxCodeAL As String = ParFOissueTKT.CmbAL.Text
        Dim lbl As String, txt As String
        For k As Int16 = 1 To 2
            dTable = GetDataTable("select VAL from MISC where CAT='TaxCode" & TaxCodeAL & "'")
            If dTable.Rows.Count > 0 Then
                For i As Int16 = 0 To dTable.Rows.Count - 1
                    lbl = "txtTaxlabel" & Trim(Str(i + 1))
                    txt = "TxtTax" & Trim(Str(i + 1))
                    ParFOissueTKT.pStep2.Controls("GrpTax").Controls(txt).Visible = True
                    ParFOissueTKT.pStep2.Controls("GrpTax").Controls(lbl).Visible = True
                    ParFOissueTKT.pStep2.Controls("GrpTax").Controls(lbl).Text = dTable.Rows(i)("VAL")
                Next
                Exit For
            End If
            TaxCodeAL = ""
        Next
    End Sub
    Public Sub GenerateComboValue(ByVal ParFOissueTKT As FOissueTKT)
        LoadTaxCodeByAl(ParFOissueTKT)
        LoadCmb_MSC(ParFOissueTKT.CmbPaxType, "PaxType")

        LoadCmb_MSC(ParFOissueTKT.CmbCurr, "select VAL from MISC where cat='CURR' and VAL1='FOI' and Status='OK' ")
        ParFOissueTKT.CmbCurr.Text = "USD"

        LoadCmb_MSC(ParFOissueTKT.CmbProduct, "select distinct Type as VAL from Charge_Comm where cat='COMM' and " & _
                    " AL in ('YY','" & ParFOissueTKT.CmbAL.Text & "') and Status='OK' " & _
                    " and '" & Now.Date & "' between ValidFrom and ValidThru")

        LoadCmb_MSC(ParFOissueTKT.CmbPromoCode, "select PromoCode as VAL from PromoCode where Status='Y' and '" & _
                    Now.Date & "' between validFrom and validthru " & _
                    " and left(PromoCode,2) in ('" & ParFOissueTKT.CmbAL.Text & "','TV','YY','..')")


        LoadCmb_MSC(ParFOissueTKT.CmbTVCharge1, "select VAL from MISC where CAT='TVCharge' or CAT='ALL' ")
        LoadCmb_MSC(ParFOissueTKT.CmbTVCharge2, "select VAL from MISC where CAT='TVCharge' or CAT='ALL' ")

        LoadCmb_MSC(ParFOissueTKT.CmbTVDiscount1, "select VAL from MISC where CAT='TVDiscount' or CAT='ALL' ")
        LoadCmb_MSC(ParFOissueTKT.CmbTVDiscount2, "select VAL from MISC where CAT='TVDiscount' or CAT='ALL' ")
    End Sub
    
    Public Function GenRCPNo(ByVal pTRXCode As String, ByVal parPOS As String) As String
        Dim NewRCPno As String, i As Int32
        Dim ThangNam As String = Format(Now.Date, "MMyy")
        ThangNam = ThangNam & parPOS
        NewRCPno = ScalarToString("RCP", "top 1 RCPno", "left(RCPno,7)='" & pTRXCode & ThangNam & "' order by RCPNO desc")
        If NewRCPno = "" Then
            NewRCPno = pTRXCode & ThangNam & "00001"
        Else
            i = CInt(NewRCPno.Substring(7)) + 1
            NewRCPno = pTRXCode & ThangNam & Format(i, "00000")
        End If
        pubVarRCPID_BeingCreated = Insert_RCP(NewRCPno, pTRXCode)
        Return NewRCPno
    End Function
    Public Sub ClearInput2(ByVal parFOissueTKT As FOissueTKT)
        parFOissueTKT.txtFare.Text = 0
        parFOissueTKT.txtExcDoc.Text = ""
        parFOissueTKT.txtExcDocDetail.Text = ""
        parFOissueTKT.txtExcDocDetail.Visible = True
        parFOissueTKT.txtExcDoc.Enabled = False
        For i As Int16 = 1 To 4
            parFOissueTKT.GrpTax.Controls("TxtTax" & i.ToString.Trim).Text = 0
            parFOissueTKT.GrpCharge.Controls("TxtCharge" & i.ToString.Trim).Text = 0
            parFOissueTKT.GrpCharge.Controls("CmbCharge" & i.ToString.Trim).Text = ".........."
        Next
        parFOissueTKT.TxtTax.Text = 0
        parFOissueTKT.TxtTax5.Text = 0
        parFOissueTKT.TxtAgtDiscVal.Text = 0
        parFOissueTKT.TxtAgtDiscPCT.Text = 0
        parFOissueTKT.txtNetToAL.Text = 0
        parFOissueTKT.txtALCommVAL.Text = 0
        parFOissueTKT.txtShownFare.Text = 0

        If parFOissueTKT.CmbPromoCode.Items.Count > 0 Then parFOissueTKT.CmbPromoCode.Text = parFOissueTKT.CmbPromoCode.Items(0).ToString
        parFOissueTKT.TxtTourCode.Text = "TOUR CODE"
        parFOissueTKT.TxtPaxName.Text = "PAX NAME"
        parFOissueTKT.txtRTG.Text = ""
        parFOissueTKT.txtFB.Text = ""
        parFOissueTKT.TxtBkgClass.Text = ""
        If parFOissueTKT.CmbCustType.Text = "CS" Then
            parFOissueTKT.CmbBooker.Text = ""
            parFOissueTKT.ChkBizTrip.Checked = True
        End If

    End Sub
    Public Sub ReleaseRCP(ByVal varRCPID As Integer)
        cmd.CommandText = "DELETE from RCP where recid=" & pubVarRCPID_BeingCreated & " and status='QQ' and fstuser='" & myStaff.SICode & "'" & _
            "; delete from rcp where status='QQ' and fstupdate < dateadd(d,-2,getdate())"
        '"; update actionlog set DoWhat='XX' where tablename+Dowhat='CREXTQQ' and F11='" & varRCPID & "' and actionBy='" & myStaff.SICode & "'"
        cmd.ExecuteNonQuery()
    End Sub
    Public Sub PrepareForOptSRV(ByVal parFOissueTKT As FOissueTKT)
        parFOissueTKT.CmdNextFr1.Enabled = True
        parFOissueTKT.GrpNewRCP.Visible = False
        parFOissueTKT.GrpSearch.Visible = False
        parFOissueTKT.GrpSearch.Top = parFOissueTKT.GrpSRV.Top
        parFOissueTKT.GrpSearch.Height = parFOissueTKT.GrpAmountInVND.Height
        parFOissueTKT.LstTKTinRCP.Items.Clear()
        parFOissueTKT.CmdRVSelected.Enabled = False
        parFOissueTKT.txtRTG.Enabled = True
        parFOissueTKT.txtFB.Enabled = True
        parFOissueTKT.TxtPaxName.Enabled = True
        parFOissueTKT.TxtTourCode.Enabled = True
        parFOissueTKT.CmbPromoCode.Enabled = True
        parFOissueTKT.TxtBkgClass.Enabled = True
        parFOissueTKT.CmbCurr.Enabled = True
        parFOissueTKT.CmbPaxType.Enabled = True
        parFOissueTKT.txtFltDate.Enabled = True
        parFOissueTKT.CmdNextFr1.Visible = True
    End Sub
    Public Sub SetBackGroudColor(ByVal parFOissueTKT As FOissueTKT)
        parFOissueTKT.TabControl1.TabPages("pStep1").BackColor = pubVarBackColor
        parFOissueTKT.TabControl1.TabPages("pStep2").BackColor = pubVarBackColor
        parFOissueTKT.txtTaxLabel1.BackColor = pubVarBackColor
        parFOissueTKT.txtTaxLabel2.BackColor = pubVarBackColor
        parFOissueTKT.TxtTaxLabel3.BackColor = pubVarBackColor
        parFOissueTKT.TxtTaxLabel4.BackColor = pubVarBackColor
        parFOissueTKT.TxtTaxLabel5.BackColor = pubVarBackColor
    End Sub
    
End Module
