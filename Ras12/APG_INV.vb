Imports SharedFunctions.MySharedFunctions
Public Class APG_INV
    Private CharEntered As Boolean = False
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private PseudoCustID As Integer
    Private mblnFirstLoadCompleted As Boolean
    Private Sub APG_INV_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MyAL.ALCode = ""
        Me.Dispose()
    End Sub

    Private Sub APG_INV_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCmb_MSC(Me.CmbAL, "APGAL")
        If Now.Month = 1 And Now.Day < 10 Then
            Me.CmbYear.Items.Add(Now.Year - 2001)
        Else
            Me.CmbYear.Items.Add(Now.Year - 2000)
        End If
        Me.CmbYear.Text = Me.CmbYear.Items(0)
        LoadCmb_VAL(Me.CmbCustomer, "select RecID as VAL, VAL1 as DIS from  MISC where cat='BSPAGT' and val1<>'' " & _
                    " and (details <>'' or val1 in (select CustShortName from customerlist where status <>'XX'))")
        Me.BackColor = pubVarBackColor
        MyAL.Domain = "GSA"
        cboStatus.SelectedIndex = 0
        mblnFirstLoadCompleted = True
    End Sub

    'Private Sub CmbCustomer_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCustomer.LostFocus

    'End Sub
    Private Sub LoadInv()
        Dim strSQL As String, DK_Status As String ' = IIf(Me.ChkXXOnly.Checked, "<>'OK'", "='OK'")
        Dim strShowLast3MonthOnlly As String = String.Empty
        Me.LblPreviewInv.Visible = False
        Me.LblRePrint.Visible = False
        Me.LblDelete.Visible = False

        Select Case cboStatus.Text
            Case "XX"
                DK_Status = " not in ('OK','QQ')"
            Case Else
                DK_Status = "='" & cboStatus.Text & "'"
        End Select
        If chkShowLast3Month.Checked Then
            strShowLast3MonthOnlly = " and DATEDIFF(m, FstUpdate,GETDATE())<=3"
        End If
        strSQL = String.Format("Select RecID, INVNo, Status,  Amount, Currency, ROE, CustShortName, CustFullName, " & _
            " CustAddress, CustTaxCode from INV where al='{0}' and status {1}" & strShowLast3MonthOnlly _
            , Me.CmbAL.Text, DK_Status)
        Me.GridInvHandlerInv.DataSource = GetDataTable(strSQL)
        Me.GridInvHandlerInv.Columns(0).Visible = False
        Me.GridInvHandlerInv.Columns("Status").Width = 32
        Me.GridInvHandlerInv.Columns("Currency").Width = 64
        Me.GridInvHandlerInv.Columns("ROE").Width = 16
        Me.GridInvHandlerInv.Columns("Amount").DefaultCellStyle.Format = "#,##0"
        Me.GridInvHandlerInv.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    Private Sub LblSave_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSave.LinkClicked
        Dim blnEditOnly As Boolean
        If GridInvHandlerInv.CurrentRow IsNot Nothing _
            AndAlso GridInvHandlerInv.CurrentRow.Cells("Status").Value = "QQ" Then
            blnEditOnly = True
        End If
        SaveInv(False, blnEditOnly)
    End Sub
    Private Sub lbkReserveInvNo_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkReserveInvNo.LinkClicked
        SaveInv(True, False)
    End Sub

    Private Function SaveInv(blnReserveInvNbrOnly As Boolean, blnEditOnly As Boolean) As Boolean
        If Not blnReserveInvNbrOnly AndAlso (Me.TxtAmt.Text = "0" Or Me.TxtNoiDung.Text.Length < 8) Then
            MsgBox("Invalid Input", MsgBoxStyle.Critical, msgTitle)
            Exit Function
        End If
        Dim InvNo As String, INVID As Integer
        Dim fstUpdate As Date = Now()
        If Now.Month = 1 And Now.Day < 10 Then fstUpdate = DateSerial(Now.Year - 1, 12, 31)
        InvNo = GenInvNo_QD153(Me.CmbAL.Text & "00" & Me.CmbYear.Text.Trim, MyAL.VAT_KyHieu)
        Try
            Dim strSaveType As String = "Create"
            If blnEditOnly Then
                INVID = GridInvHandlerInv.CurrentRow.Cells("RecId").Value
                strSaveType = "EditQQ"
            Else
                INVID = Insert_INV("E", InvNo, Me.CmbAL.Text, -56, fstUpdate)
                If blnReserveInvNbrOnly Then
                    cmd.CommandText = "Update INV set Status='QQ' where RecId=" & INVID
                    cmd.ExecuteNonQuery()
                End If
            End If

            cmd.CommandText = UpdateLogFile("INV", "APGINV", INVID, Me.TxtNoiDung.Text, strSaveType, "", "")
            cmd.ExecuteNonQuery()
            cmd.CommandText = Update_INV("S", Me.txtCustName.Text, Me.txtCustAddrr.Text, Me.txtTaxCode.Text _
                                         , CDec(Me.TxtAmt.Text), "CSH_VND_" & Me.TxtAmt.Text & "_", 0, INVID, PseudoCustID)
            cmd.ExecuteNonQuery()

            If GridInvHandlerInv.RowCount > 0 AndAlso GridInvHandlerInv.CurrentRow.Cells("Status").Value = "QQ" Then
                cmd.CommandText = "update Inv set Status='OK' where RecId=" & Me.GridInvHandlerInv.CurrentRow.Cells("RecID").Value
                cmd.ExecuteNonQuery()
            End If

            Me.TxtAmt.Text = 0
            Me.TxtNoiDung.Text = ""
            LoadInv()
            Return True
        Catch ex As Exception
            MsgBox("Unable Creating Invoice", MsgBoxStyle.Critical, msgTitle)
            Return False
        End Try
    End Function

    Private Sub GridInvHandlerInv_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridInvHandlerInv.CellClick
        If e.RowIndex < 0 Then Exit Sub
        Select Case Me.GridInvHandlerInv.CurrentRow.Cells("status").Value
            Case "OK"
                Me.LblDelete.Visible = True
            Case "QQ"
                Me.LblDelete.Visible = True
                TxtAmt.Text = GridInvHandlerInv.CurrentRow.Cells("Amount").Value
                CmbCustomer.SelectedIndex = CmbCustomer.FindStringExact(GridInvHandlerInv.CurrentRow.Cells("CustShortName").Value)
            Case Else
                Me.LblDelete.Visible = False
        End Select

        Me.LblPreviewInv.Visible = Me.LblDelete.Visible
    End Sub
    'Private Sub GridInvHandlerInv_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridInvHandlerInv.CellContentClick
    '    If e.RowIndex < 0 Then Exit Sub
    '    Select Case Me.GridInvHandlerInv.CurrentRow.Cells("status").Value
    '        Case "OK"
    '            Me.LblDelete.Visible = True
    '        Case "QQ"
    '            Me.LblDelete.Visible = True
    '            TxtAmt.Text = GridInvHandlerInv.CurrentRow.Cells("Amount").Value
    '            CmbCustomer.SelectedIndex = CmbCustomer.FindStringExact(GridInvHandlerInv.CurrentRow.Cells("CustShortName").Value)
    '        Case Else
    '            Me.LblDelete.Visible = False
    '    End Select

    '    Me.LblPreviewInv.Visible = Me.LblDelete.Visible
    'End Sub

    Private Sub LblDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDelete.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("Inv", "X0", Me.GridInvHandlerInv.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        LoadInv()
    End Sub
    Private Sub ChkXXOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        LoadInv()
    End Sub

    Private Sub LblPreviewInv_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblPreviewInv.LinkClicked
        Dim fName As String, InvNoToPrint As String
        fName = "R12_VAT_APG.xlt"

        InvNoToPrint = Me.GridInvHandlerInv.CurrentRow.Cells("InvNo").Value
        InHoaDon(Application.StartupPath, fName, "V", InvNoToPrint, Now.Date, Now.Date, 0, Me.CmbAL.Text, "TVS")
        Me.LblRePrint.Visible = True
    End Sub
    Private Sub LblRePrint_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblRePrint.LinkClicked
        Dim myAnswer As Int16, fName As String

        Dim InvID As Integer = Me.GridInvHandlerInv.CurrentRow.Cells("RecID").Value
        Dim InvNo As String = Me.GridInvHandlerInv.CurrentRow.Cells("InvNo").Value
        fName = "R12_VAT_APG.xlt"
        myAnswer = MsgBox("Do You Want to Print It? ", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
        If myAnswer = vbNo Then
            Exit Sub
        End If
        InHoaDon(Application.StartupPath, fName, "O", InvNo, Now.Date, Now.Date, 0, Me.CmbAL.Text, MySession.Domain)
        cmd.CommandText = UpdateTblINVHistory(InvID, InvNo, 1)
        cmd.ExecuteNonQuery()
        Me.LblRePrint.Visible = False
    End Sub

    Private Sub TxtAmt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtAmt.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub
    Private Sub TxtAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtAmt.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub TxtAmt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAmt.LostFocus
        Dim aa As Double
        Try
            aa = CDbl(Me.TxtAmt.Text)
            Me.TxtAmt.Text = Format(aa, "#,##0")
        Catch ex As Exception
        End Try
    End Sub
    Private Sub CmbAL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbAL.LostFocus
        MyAL.ALCode = Me.CmbAL.Text

        lbkReserveInvNo.Visible = (CmbAL.Text = "UA")
        LoadInv()
    End Sub

    Private Sub cboStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStatus.SelectedIndexChanged
        If mblnFirstLoadCompleted Then
            LoadInv()
        End If
    End Sub

    Private Sub CmbCustomer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCustomer.SelectedIndexChanged

        If Not mblnFirstLoadCompleted Then
            Exit Sub
        End If
        Dim dTbl As DataTable
        Me.txtCustName.Text = ""
        Me.txtCustAddrr.Text = ""
        Me.txtTaxCode.Text = ""
        If Me.CmbCustomer.Text.Length = 7 _
            AndAlso (Me.CmbCustomer.Text.Substring(0, 3) = "HAN" Or Me.CmbCustomer.Text.Substring(0, 3) = "SGN") Then
            PseudoCustID = 0
            dTbl = GetDataTable("select Details, Description from MISC where recID=" & Me.CmbCustomer.SelectedValue)
            Try
                Me.txtCustName.Text = dTbl.Rows(0)("Details")
                Me.txtCustAddrr.Text = dTbl.Rows(0)("Description").ToString.Split("|")(0)
                Me.txtTaxCode.Text = dTbl.Rows(0)("Description").ToString.Split("|")(1)
            Catch ex As Exception
            End Try
        Else
            dTbl = GetDataTable("select RecID, custFullName, CustAddress, CustTaxCode from CustomerList where CustShortname='" & _
                                Me.CmbCustomer.Text & "' and status='OK'")
            If dTbl.Rows.Count > 0 Then
                Me.txtCustName.Text = dTbl.Rows(0)("CustFullName")
                Me.txtCustAddrr.Text = dTbl.Rows(0)("CustAddress")
                Me.txtTaxCode.Text = dTbl.Rows(0)("CustTaxCode")
                PseudoCustID = dTbl.Rows(0)("RecID")
            End If

        End If
        'LoadInv()
    End Sub

    Private Sub chkShowLast3Month_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowLast3Month.CheckedChanged
        If mblnFirstLoadCompleted Then
            LoadInv()
        End If
    End Sub
End Class
