Imports SharedFunctions.Crd_Ctrl
Imports SharedFunctions.MySharedFunctions
Public Class frmMISC
    Private MyCust As New objCustomer
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private WhatAction As String
    Private CharEntered As Boolean = False
    Private strDKCust As String = ""
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal parWhatAction As String)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        WhatAction = parWhatAction
    End Sub

    Private Sub MISC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ccrCust As String = " and RecID in (select val1 from cwt.dbo.misc where cat='KCHOLDER')"
        Me.BackColor = pubVarBackColor
        strDKCust = " in (select custid from cust_Detail where status='OK' and cat='Channel' And Val in " & myStaff.CAccess & ")"
        If InStr(WhatAction, "CRLimit") > 0 Then
            Me.Text = "TransViet Airlines :: RAS :. Payment Options"
            Me.PnlCC_Customer.Visible = True
            Me.PnlCC_Customer.Left = 0
            Me.PnlCC_Customer.Height = 515
            Me.PnlCC_Customer.Width = 808
            If myStaff.City = "SGN" Then
                Me.CmbVAT.Visible = True
                Me.Label2.Visible = True
                Me.CmbVAT.Text = "TVTR"
            Else ' HAN ko ap dung han che in HD khach cong no theo Cty
                Me.CmbVAT.Items.Add("")
                Me.CmbVAT.Text = ""
            End If
            GenCustList()
            GenALList()
            Me.ChkAll.Checked = True
        ElseIf InStr(WhatAction, "CRX") > 0 Then
            Me.Text = "TransViet Airlines :: RAS :. DEB and Over Credit Approval"
            Me.PnlCRExt.Visible = True
            Me.PnlCRExt.Top = 0
            Me.PnlCRExt.Height = 515
            Me.PnlCRExt.Width = 808
            LoadGridQ4Approval(WhatAction)
        End If
        CheckRightForALLForm(Me)
    End Sub
    Private Sub LoadGridQ4Approval(ByVal pWho As String)
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim strSQL As String, UptoDate As Date, BackDate As Int16 = 0
        Me.txtCRExtNote.Text = ""
        If myStaff.CAccess.Contains("CS") Then
            BackDate = -1
            If Now.DayOfWeek = DayOfWeek.Monday Then BackDate = -2
        End If
        UptoDate = Now.Date.AddDays(BackDate)
        Try
            cmd.CommandText = "Drop table #CREXT"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        cmd.CommandText = "select * into #CREXT from Actionlog where tableName='CREXT' and DoWhat='QQ' and Actiondate >'" & UptoDate & "'"
        cmd.ExecuteNonQuery()
        Try
            cmd.CommandText = "insert #CREXT (TableName, ActionBy, ActionDate, doWhat, F1, F2, F3, F4, F5, F6, F7, F8, City, F9, F10, F11, F12)" & _
                "select TableName, ActionBy, ActionDate, doWhat, F1, F2, F3, F4, F5, F6, F7, F8, City, F9, F10, RecID, 'FLX' " & _
                "FROM [42.117.5.86].flx.dbo.ActionLog where tableName='CREXT' and DoWhat='QQ' and Actiondate >'" & UptoDate & "'"
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try

        strSQL = "Select RecID, ActionDate as RQDate, F1 as TRX, f5 as RQSTAmt, f7 as RQUser, f6 as RQSTType, " & _
            "F2 as CustID, F3 as ShortName, F4 as TKTs, F9 as Eml, F11, F12 from #CREXT where dowhat='QQ'"
        If Strings.Right(pWho, 2) = "FO" Then
            strSQL = strSQL & " and F6='DEB' "
        ElseIf Strings.Right(pWho, 2) = "AC" Then
            strSQL = strSQL & " and F6='ITP' "
        Else
            strSQL = strSQL & " and F6 not in ('DEB','ITP') and F2 " & strDKCust
        End If
        strSQL = strSQL & " order by F7, F1"
        Me.GridQ4Approval.DataSource = GetDataTable(strSQL)
        Me.GridQ4Approval.Columns("RecID").Visible = False
        Me.GridQ4Approval.Columns("CustID").Visible = False
        Me.GridQ4Approval.Columns("RQUser").Width = 64
        Me.GridQ4Approval.Columns("RqstType").Width = 64
        Me.GridQ4Approval.Columns("RQSTAmt").DefaultCellStyle.Format = "#,##0.00"
        Me.GridQ4Approval.Columns("RQSTAmt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.CmdCRApprove.Enabled = False
        Me.CmdCRReject.Enabled = False
        Me.txtCRExtNote.Enabled = False
    End Sub
    Private Sub GenALList()
        Dim dTable As DataTable = GetDataTable(myStaff.TVA)
        For i As Int16 = 0 To dTable.Rows.Count - 1
            Me.ChkAL.Items.Add(dTable.Rows(i)("VAL"))
        Next
        Me.ChkAL.Items.Add("FT")
    End Sub
    Private Sub GenCustList()
        Dim strSQL As String
        strSQL = "Select RecID as VAL, custShortName as DIS from CustomerList where status='OK' and " & _
            "custshortname <> custFullName and email <>'' and RecID " & strDKCust & " order by CustShortName"
        RemoveHandler CmbCC_Cust_Customer.SelectedIndexChanged, AddressOf CmbCC_Cust_Customer_SelectedIndexChanged
        LoadCmb_VAL(Me.CmbCC_Cust_Customer, strSQL)

        AddHandler CmbCC_Cust_Customer.SelectedIndexChanged, AddressOf CmbCC_Cust_Customer_SelectedIndexChanged
        Call CmbCC_Cust_Customer_SelectedIndexChanged(CmbCC_Cust_Customer, System.EventArgs.Empty)
        LoadGridCC_Cust()
    End Sub
    Private Sub LoadGridCC_Cust()
        Dim strSQL As String
        strSQL = "Select CustID, CustShortName, CRCoef, PPCoef, AL, MinBLC, FstUser, FstUpdate, RecID from CC_Setting " & _
            " where CustID " & strDKCust & " order by CustShortName "
        Me.GridCC_Cust.DataSource = GetDataTable(strSQL)
        Me.GridCC_Cust.Columns("RecID").Visible = False
        Me.GridCC_Cust.Columns("CRCoef").Width = 56
        Me.GridCC_Cust.Columns("CustID").Width = 56
        Me.GridCC_Cust.Columns("FstUser").Width = 56
        Me.GridCC_Cust.Columns("PPCoef").Width = 56
        Me.GridCC_Cust.Columns("AL").Width = 128
        Me.GridCC_Cust.Columns("MinBLC").Width = 75
        Me.GridCC_Cust.Columns("MinBLC").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridCC_Cust.Columns("MinBLC").DefaultCellStyle.Format = "#,##0"
        MyCust.CustID = 0
        Me.LblDelete.Enabled = False
        Me.CmdCC_Cust_Save.Enabled = False
        Me.LblChangeMinBLC.Visible = False
    End Sub

    Private Sub CmbCC_Cust_Customer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles CmbCC_Cust_Customer.SelectedIndexChanged
        MyCust.CustID = Me.CmbCC_Cust_Customer.SelectedValue
        Me.txtCC_Cust_CustName.Text = MyCust.FullName
        Me.CmdCC_Cust_Save.Enabled = True
    End Sub
    Private Sub GridCC_Cust_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCC_Cust.CellClick
        If e.RowIndex < 0 Then Exit Sub
        Dim ALList As String = Me.GridCC_Cust.Item("AL", e.RowIndex).Value
        Me.CmbCC_Cust_Customer.Text = Me.GridCC_Cust.Item("CustShortName", e.RowIndex).Value
        Me.txtCRCoef.Text = Format(Me.GridCC_Cust.Item("CrCoef", e.RowIndex).Value, "0.0")
        Me.txtPPCoef.Text = Format(Me.GridCC_Cust.Item("PPCoef", e.RowIndex).Value, "0.0")
        If Me.GridCC_Cust.Item("AL", e.RowIndex).Value = "YY" Then
            Me.ChkAll.Checked = True
        Else
            Me.ChkAll.Checked = False
            For i As Int16 = 0 To Me.ChkAL.Items.Count - 1
                If InStr(ALList, Me.ChkAL.Items(i).ToString) > 0 Then
                    Me.ChkAL.SetItemChecked(i, True)
                End If
            Next
        End If
        Me.LblDelete.Enabled = True
        Me.CmdCC_Cust_Save.Enabled = True
    End Sub
    Private Sub CmdCC_Cust_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCC_Cust_Save.Click
        Dim ALList As String, tmpID As Integer
        Dim NewPmtType As String = IIf(CDec(Me.txtCRCoef.Text) > 0, "PSP", "PPD")
        Dim CC_settingStatus As String = IIf(CDec(Me.txtCRCoef.Text) > 0, "QQ", "OK")
        MyCust.CustID = Me.CmbCC_Cust_Customer.SelectedValue
        If MyCust.LstReconcile = Now.Date Then Exit Sub
        If Me.ChkAll.Checked Then
            ALList = "YY"
        Else
            ALList = ""
            For i As Int16 = 0 To Me.ChkAL.Items.Count - 1
                If Me.ChkAL.GetItemChecked(i) = True Then
                    ALList = ALList & "|" & Me.ChkAL.Items(i).ToString
                End If
            Next
            ALList = ALList.Substring(1)
        End If
        tmpID = ScalarToInt("cc_Setting", "RecID ", "status <>'XX' and custid=" & MyCust.CustID)
        If tmpID > 0 Then ' Khach cong no cu, edit. check xem co sua pmttype ko? neu co thi balance phai =0 first
            If NewPmtType <> MyCust.DelayType Then
                If RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, False, "Chk 0BLC 4PmtType CHGE", Conn, myStaff.SICode, CnStr) <> 0 Then
                    MsgBox("Cant Change Delayed Payment Option until The Current Balance Settled to Zero", MsgBoxStyle.Critical, msgTitle)
                    Exit Sub
                Else
                    cmd.CommandText = "Insert into ChotCongNo (CustId,CustShortName,VND_Avail,AsOf,FstUser) values(" _
                                  & MyCust.CustID & ",'" & MyCust.ShortName & "',0,getdate(),'" & myStaff.SICode & "')"
                    cmd.ExecuteNonQuery()
                End If
                If NewPmtType = "PPD" Then
                    cmd.CommandText = ChangeStatus_ByDK("kyBaoCao", "XX", "CustID=" & MyCust.CustID & " and status='OK'")
                    cmd.ExecuteNonQuery()
                End If
            End If
            cmd.CommandText = "Update CC_Setting  set crCoef=" & CDec(Me.txtCRCoef.Text) & _
                ", PPCoef=" & CDec(Me.txtPPCoef.Text) & ", VAT='" & Me.CmbVAT.Text & "' Where recID=" & tmpID & _
                ";" & UpdateLogFile("CC_Setting", "CoefCHG", "CustID=" & Me.CmbCC_Cust_Customer.SelectedValue, _
                "NewPSP:" & Me.txtCRCoef.Text, "NewPPD:" & Me.txtPPCoef.Text, "Old=" & MyCust.DelayType & _
                ":" & MyCust.Coef, "RecID=" & tmpID, "VAT:" & Me.CmbVAT.Text, "", "")
            cmd.ExecuteNonQuery()
        Else
            cmd.CommandText = "insert into CC_Setting (CustID, CustShortName, CRCoef, PPCoef,  AL, status, VAT, FstUser) values ('" & _
                MyCust.CustID & "','" & Me.CmbCC_Cust_Customer.Text & "'," & CDec(Me.txtCRCoef.Text) & "," & _
                CDec(Me.txtPPCoef.Text) & ",'" & ALList & "','" & CC_settingStatus & "','" & Me.CmbVAT.Text & "','" & myStaff.SICode & "')" & _
                ";" & UpdateLogFile("CC_Setting", "NEW CC", "CustID=" & Me.CmbCC_Cust_Customer.SelectedValue, _
                "NewPSP:" & Me.txtCRCoef.Text, "NewPPD:" & Me.txtPPCoef.Text, "VAT:" & Me.CmbVAT.Text, "", "", "", "")
            cmd.ExecuteNonQuery()
            If CC_settingStatus = "QQ" Then MsgBox("Request Has Been Sent To Accounting for Updating Invoice Period", MsgBoxStyle.Information, msgTitle)
        End If
        LoadGridCC_Cust()
    End Sub

    Private Sub txtCRLimit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
         txtCRCoef.KeyDown, txtPPCoef.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub

    Private Sub txtCRLimit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles _
         txtCRCoef.KeyPress, txtPPCoef.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtCRCoef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCRCoef.LostFocus, txtPPCoef.LostFocus
        Dim aa As Decimal, txt As TextBox = CType(sender, TextBox)
        aa = CDec(txt.Text)
        txt.Text = Format(aa, "#,##0.0")

    End Sub
    Private Sub CmdCRApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCRApprove.Click
        Dim TRX As String = Me.GridQ4Approval.CurrentRow.Cells("TRX").Value.ToString
        Dim RQType As String = Me.GridQ4Approval.CurrentRow.Cells("RQSTType").Value.ToString
        Dim KQ As String
        Dim emlBody As String = "Credit extension for " & Me.GridQ4Approval.CurrentRow.Cells("ShortName").Value
        Dim emailList As String = Me.GridQ4Approval.CurrentRow.Cells("Eml").Value.ToString.Replace(";", ".sgn@transviet.com;") & ".sgn@transviet.com"
        KQ = ApproveDEB()
        emlBody = emlBody & KQ
        If emailList = ".sgn@transviet.com" Or RQType = "ITP" Or RQType = "DEB" Then
            MsgBox(KQ, MsgBoxStyle.Information, msgTitle)
        Else
            Process.Start(String.Format("mailto:{0}?subject={1}&body={2}", emailList, TRX, emlBody))
            SendKeys.SendWait("%S")
        End If
    End Sub
    Private Function ApproveDEB() As String
        MyCust.CustID = Me.GridQ4Approval.CurrentRow.Cells("CustID").Value
        If myStaff.City = "SGN" And myStaff.Location = "TVH" Then
            If InStr("TA_TO_CA", MyCust.CustType) > 0 And myStaff.SICode <> "LLN" And InStr("PSP_PPD", Me.GridQ4Approval.CurrentRow.Cells("RqstType").Value) > 0 Then
                Dim MaxOverDue As Int16 = ScalarToInt("MISC", "VAL", "cat='MAXOVRDUE'")
                Dim MaxOverCRD As Decimal = ScalarToDec("MISC", "VAL", "cat='MAXOVRCRD'")
                If MyCust.CurrBLC < MaxOverCRD Then Return " Rejected due Over Safe Limit."
                If MyCust.LstDue > DateSerial(2013, 4, 1) And MyCust.LstDue < Now.Date.AddDays(MaxOverDue) Then Return " Rejected due Aged OverDue."
            End If
        End If
        UpdateStatusCRX("OK")
        Return " Approved."
    End Function
    Private Sub GridQ4Approval_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridQ4Approval.CellContentClick
        Me.txtCRExtNote.Text = ""
        Me.txtCRExtNote.Enabled = True
        Me.txtCRExtNote.Focus()
    End Sub
    Private Sub UpdateStatusCRX(ByVal pStatus As String)
        Dim RecNo As Integer = Me.GridQ4Approval.CurrentRow.Cells("RecID").Value
        Dim App As String = Me.GridQ4Approval.CurrentRow.Cells("F12").Value
        If App = "FLX" Then RecNo = Me.GridQ4Approval.CurrentRow.Cells("F11").Value
        Dim TblName As String = IIf(App = "FLX", "[42.117.5.86].flx.dbo.", "") & "actionlog"
        Me.CmdCRApprove.Enabled = False
        Me.CmdCRReject.Enabled = False
        cmd.CommandText = String.Format("update {0} set DoWhat='{1}', ActionBy='{2}', ActionDate=getdate()" & _
            ", F8 ='{3}' where recid={4}", TblName, pStatus, myStaff.SICode, Me.txtCRExtNote.Text.Replace("--", ""), RecNo)
        cmd.ExecuteNonQuery()
        LoadGridQ4Approval(WhatAction)
    End Sub
    Private Sub CmdCRReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCRReject.Click
        UpdateStatusCRX("XX")
    End Sub
    Private Sub txtCRExtNote_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCRExtNote.LostFocus
        If Me.txtCRExtNote.Text.Length > 0 Then
            Me.CmdCRApprove.Enabled = True
            Me.CmdCRReject.Enabled = True
        End If
    End Sub
    Private Sub LblDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        If RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, False, "Chk 0BLC B4 DEL CC_Cust", Conn, myStaff.SICode, CnStr) = 0 Then
            cmd.CommandText = "delete from cc_setting where recid=" & Me.GridCC_Cust.CurrentRow.Cells("RecID").Value & _
                ";" & UpdateLogFile("CC_Setting", "DEL CC", "CustID=" & Me.CmbCC_Cust_Customer.SelectedValue, _
                "CurrPSP:" & Me.GridCC_Cust.CurrentRow.Cells("CRCoef").Value, "CurrPPD:" & Me.GridCC_Cust.CurrentRow.Cells("PPCoef").Value, _
                "RecID" & Me.GridCC_Cust.CurrentRow.Cells("RecID").Value, "", "", "", "")
            cmd.ExecuteNonQuery()
            cmd.CommandText = "insert ChotCongNo (CustID, CustShortName, AsOf, VND_Avail, FstUser) values (" & _
                MyCust.CustID & ",'" & MyCust.ShortName & "', '" & Format(Now, "dd-MMM-yy") & "',0,'" & myStaff.SICode & "')"
            LoadGridCC_Cust()
        Else
            MsgBox("Cant Delete This Record Until Final Settlement to Zero Balance", MsgBoxStyle.Critical, msgTitle)
        End If
    End Sub

    Private Sub txtCRCoef_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCRCoef.TextChanged
        If CDec(txtCRCoef.Text) > 0 Then Me.txtPPCoef.Text = 0
    End Sub
    Private Sub txtPPCoef_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPPCoef.TextChanged
        If CDec(txtPPCoef.Text) > 0 Then Me.txtCRCoef.Text = 0
    End Sub
    Private Sub ChkAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAll.Click

        For i As Int16 = 0 To Me.ChkAL.Items.Count - 1
            Me.ChkAL.SetItemChecked(i, Me.ChkAll.Checked)
        Next
    End Sub
    Private Sub ChkAL_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAL.SelectedIndexChanged
        Me.ChkAll.Checked = False
    End Sub
    Private Sub GridCC_Cust_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCC_Cust.CellContentClick
        MyCust.CustID = Me.GridCC_Cust.CurrentRow.Cells("CustID").Value
        Me.txtMinBLC.Text = Format(Me.GridCC_Cust.CurrentRow.Cells("MinBLC").Value, "#,##0")
        Me.LblChangeMinBLC.Visible = True
    End Sub

    Private Sub LblChangeMinBLC_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblChangeMinBLC.LinkClicked
        cmd.CommandText = UpdateLogFile("CC_Setting", "ChgeMinBLC", "ID:" & Me.GridCC_Cust.CurrentRow.Cells("RecID").Value.ToString, _
                                        "OLD:" & Me.GridCC_Cust.CurrentRow.Cells("MinBLC").Value, "CustID:" & Me.GridCC_Cust.CurrentRow.Cells("CustID").Value, "", "")
        cmd.CommandText = cmd.CommandText & "; Update CC_setting set MinBLC=" & CDec(Me.txtMinBLC.Text) & " where recID=" & _
            Me.GridCC_Cust.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridCC_Cust()
    End Sub
End Class