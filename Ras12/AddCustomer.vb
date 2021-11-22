Public Class AddCustomer
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private WhoIs As String
    Private MyCust As New objCustomer
    Private Sub AddCustomer_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Sub New(ByVal pWho As String)
        InitializeComponent()
        WhoIs = pWho
    End Sub
    Private Sub AddCustomer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.GridCustList.BackgroundColor = pubVarBackColor
        RefreshCustList()
        CheckRightForALLForm(Me)
        Me.txtShortName.Focus()
        Me.txtCustCode.Text = "-56"
        LoadCmb_MSC(Me.CmbChannel, "select VAL from MISC where CAT='Channel' and VAL <>'WK' and val in " & myStaff.CAccess)
        If InStr(WhoIs, "CS") > 0 Then
            Me.CmbChannel.Text = "CS"
            Me.LckChkNotCustomer.Enabled = False
            Me.LckChkNotCustomer.Checked = False
        ElseIf myStaff.City = "SGN" Then
            Me.CmbChannel.Text = "TA"
            Dim i As Int16 = ScalarToInt("MISC", "count(*)", "cat='BSPAGT' and VAL1=''")
            If i > 0 Then
                MsgBox("Con IATA code chua dc gan ten Dai ly.")
            End If
        End If
    End Sub

    Private Sub RefreshCustList()
        Dim StrSQL As String
        If Me.ChkXXOnly.Checked Then
            StrSQL = "select * from Customer where custShortName <> CustFullName and status in ('EX','QQ') and city='" & myStaff.City & "'"
        Else
            StrSQL = "select * from CustomerList where custShortName <> CustFullName and status='OK'" & _
                    " and  RecID in (select CustID from cust_Detail where status+cat='OKChannel' and val in " & myStaff.CAccess & ")"
        End If
        Me.GridCustList.DataSource = GetDataTable(StrSQL)
        Me.GridCustList.Columns("RecID").Width = 45
        Me.GridCustList.Columns("CustTaxCode").Width = 75
        Me.GridCustList.Columns("CustShortName").Width = 75
        Me.LckCmdSave.Visible = False
        Me.LckCmdAdd.Visible = False
    End Sub
    Private Sub txtCustCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.GotFocus
        Me.txtShortName.Text = ""
        Me.txtFullName.Text = ""
        Me.txtTaxCode.Text = ""
        Me.txtAddress.Text = ""
    End Sub
    Private Sub txtCustCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
            txtShortName.LostFocus, txtCustCode.LostFocus
        Dim txt As TextBox = CType(sender, TextBox)
        Dim dtbl As DataTable, DaTonTai As Boolean = False, StrSQL As String

        StrSQL = String.Format("select * from CustomerList where RecID={0}", Me.txtCustCode.Text)
        If Me.txtShortName.Text.Length > 0 Then
            StrSQL = String.Format("{0} or CustShortName ='{1}'", StrSQL, Me.txtShortName.Text)
        End If
        If Me.txtFullName.Text.Length > 0 Then
            StrSQL = String.Format("{0} or CustFullName='{1}'", StrSQL, Me.txtFullName.Text)
        End If
        dtbl = GetDataTable(StrSQL)
        For i As Int16 = 0 To dtbl.Rows.Count - 1
            DaTonTai = True
            Me.txtCustCode.Text = dtbl.Rows(i)("RecID")
            If txt.Name <> "txtShortName" Then
                Me.txtShortName.Text = dtbl.Rows(i)("CustShortName")
            End If
            Me.txtFullName.Text = dtbl.Rows(i)("CustFullName")
            Me.txtTaxCode.Text = dtbl.Rows(i)("CustTaxCode")
            Me.txtAddress.Text = dtbl.Rows(i)("CustAddress")
            Exit For

        Next
        If DaTonTai Then MsgBox(" This Customer Exists. Please Check!", MsgBoxStyle.Critical, msgTitle)
        Me.LckCmdAdd.Visible = Not DaTonTai
        Me.LckCmdSave.Visible = DaTonTai
    End Sub
    Private Function SaiSoDT() As Boolean

        If Me.TxtPhone.Text.Trim <> "" Then
            Me.TxtPhone.Text = Me.TxtPhone.Text.Replace(" ", "")
            If InStr("04_08_09", Me.TxtPhone.Text.Substring(0, 2)) > 0 And Me.TxtPhone.Text.Length <> 10 Then GoTo errMsg
            If Me.TxtPhone.Text.Substring(0, 2) = "01" And Me.TxtPhone.Text.Length <> 11 Then GoTo errMsg
            For i As Int16 = 0 To Me.TxtPhone.Text.Length - 1
                If InStr("0123456789", Me.TxtPhone.Text.Substring(i, 1)) = 0 Then GoTo errMsg
            Next
        End If
        Return False
errMsg:
        MsgBox("Invalid Phone No. Please Check!", MsgBoxStyle.Critical, msgTitle)
        Return True
    End Function
    Private Function SaiEmail(ByVal pSave As Int16) As Boolean
        Dim j As Integer, peMail As String = Me.TxtEmail.Text.Trim
        If peMail <> "" AndAlso peMail.Substring(0, 3) <> "Plz" Then
            peMail = peMail.Replace(" ", "")
            peMail = peMail.Replace(",", ";")
            For i As Int16 = 0 To UBound(peMail.Split(";"))
                j = InStr(peMail.Split(";")(i), "@")
                If j = 0 Then GoTo errMsg
                If InStr(j + 1, peMail.Split(";")(i), "@") > 0 Then GoTo errMsg
                If InStr(j + 1, peMail.Split(";")(i), ".") = 0 Then GoTo errMsg
                If i > 1 Then GoTo errMsg
            Next
        End If
        If pSave = 1 Then
            j = ScalarToInt("cc_Setting", "RecID", "Custid=" & Me.txtCustCode.Text)
            If j > 0 Then
                If peMail = "" Or InStr(peMail, "@") = 0 Then
                    MsgBox("Need At Least One Email Addresses", MsgBoxStyle.Information, msgTitle)
                    Return True
                End If
            End If
        End If
        Return False
errMsg:
        MsgBox("Invalid Email Addresses", MsgBoxStyle.Information, msgTitle)
        Return True
    End Function
    Private Function SaiName() As Boolean
        Dim tmpCustShortName As String
        If Me.txtShortName.Text.Length = Me.txtFullName.Text.Length And Me.txtShortName.Text = Me.txtFullName.Text Then
            MsgBox("Short Name and Full Name Cant Be the Same", MsgBoxStyle.Information, msgTitle)
            Return True
        End If
        For i As Int16 = 0 To Me.txtShortName.Text.Length - 1
            If InStr("?*#-/!", Me.txtShortName.Text.Substring(i, 1)) > 0 Then
                MsgBox("Short Name Cant Contain Special Charator", MsgBoxStyle.Information, msgTitle)
                Return True
            End If
        Next
        cmd.CommandText = String.Format("Select distinct custShortName from rcp where custShortname='{0}'" & _
                " UNION select distinct CustShortName from cc_Setting where custShortname='{0}'", Me.txtShortName.Text)
        tmpCustShortName = cmd.ExecuteScalar
        If tmpCustShortName <> "" Then
            MsgBox("This Name Already Exists", MsgBoxStyle.Information, msgTitle)
            Return True
        End If
        Return False
    End Function
    Private Sub CmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LckCmdAdd.Click
        Dim tmpCustID As Integer, vStatus As String = "OK"
        If MsgBox("Are You Sure That All Input Are Correct and Wanna Create This New Customer?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, msgTitle) = vbNo Then Exit Sub
        If SaiSoDT() Or SaiEmail(0) Or SaiName() Then
            Exit Sub
        End If
        If Me.TxtEmail.Text.Substring(0, 3) = "Plz" Then Me.TxtEmail.Text = ""
        tmpCustID = ScalarToInt("customer", "recID", String.Format(" custshortname='{0}'", Me.txtShortName.Text))
        If tmpCustID > 0 Then
            MsgBox("Customer Already Exists.", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If Me.LckChkNotCustomer.Checked Then vStatus = "QQ"
        tmpCustID = MyCust.AddCustomer(Me.txtShortName.Text.Replace("--", ""), Me.txtFullName.Text.Replace("--", ""), _
            Me.txtTaxCode.Text.Replace("--", ""), Me.txtAddress.Text.Replace("--", ""), Me.TxtEmail.Text.Replace(",", ":").Replace("--", ""), _
            Me.TxtPhone.Text, vStatus)
        MyCust.InsertCustDetail(tmpCustID, "Channel", Me.CmbChannel.Text, False)
        MyCust.InsertCustDetail(tmpCustID, "AL", "YY", False)
        RefreshCustList()
        MsgBox("Customer Has Been Created. Click to [More Detail] to set AL and Channel for it", MsgBoxStyle.Information, msgTitle)
    End Sub

    Private Sub CmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LckCmdSave.Click
        Dim i As Int16 = MsgBox("Are You Sure That All Input Are Correct and Wanna Save Changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, msgTitle)
        If i = vbNo Then Exit Sub
        If SaiSoDT() Or SaiEmail(1) Or Me.txtAddress.Text = "" Or Me.txtFullName.Text = "" Then
            Exit Sub
        End If
        MyCust.SaveChange(Me.txtFullName.Text.Replace("--", ""), Me.txtTaxCode.Text.Replace("--", ""), Replace(Me.TxtEmail.Text, ",", ";"), Me.TxtPhone.Text, Me.txtAddress.Text.Replace("--", ""), Me.txtCustCode.Text, Me.CmbLocation.Text)
        RefreshCustList()
    End Sub

    Private Sub GridCustList_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCustList.CellClick
        Dim custID As Integer, StrSQL As String
        On Error GoTo ErrHandler
        custID = Me.GridCustList.Item("RecID", e.RowIndex).Value
        On Error GoTo 0
        Me.LckCmdAdd.Visible = False
        Me.LckCmdSave.Visible = True
        Me.LckLblDeactivate.Visible = Not Me.ChkXXOnly.Checked
        Me.LckLblReinstate.Visible = Me.ChkXXOnly.Checked
        Me.txtCustCode.Text = Me.GridCustList.CurrentRow.Cells("RecID").Value
        Me.txtShortName.Text = Me.GridCustList.CurrentRow.Cells("CustShortName").Value
        Me.txtFullName.Text = Me.GridCustList.CurrentRow.Cells("CustFullName").Value
        Me.txtTaxCode.Text = Me.GridCustList.CurrentRow.Cells("CustTaxCode").Value
        Me.txtAddress.Text = Me.GridCustList.CurrentRow.Cells("CustAddress").Value
        Me.TxtEmail.Text = Me.GridCustList.CurrentRow.Cells("Email").Value
        Me.TxtPhone.Text = Me.GridCustList.CurrentRow.Cells("Phone").Value
        Me.CmbLocation.Text = Me.GridCustList.CurrentRow.Cells("Location").Value
        StrSQL = String.Format("select SBU, AL, Channel, CustLevel, ValidFrom, ValidThru from Cust_Channel_level where " & _
            " status ='OK' and custID={0}", custID)
        Me.GridCommDetail.DataSource = GetDataTable(StrSQL)
        Me.GridCommDetail.Columns("SBU").Width = 32
        Me.GridCommDetail.Columns("AL").Width = 32
        Me.GridCommDetail.Columns("Channel").Width = 32
        Me.GridCommDetail.Columns("CustLevel").Width = 32
        Me.GridCommDetail.Columns("ValidFrom").Width = 64
        Me.GridCommDetail.Columns("ValidThru").Width = 64
ErrHandler:
        Me.TxtEmail.ForeColor = Color.Black
        Exit Sub
    End Sub

    Private Sub LblDeactivate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblDeactivate.LinkClicked
        Dim i As Integer
        i = MsgBox("This Gonna DeActivate Selected Customer. Are You Sure?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, msgTitle)
        If i = vbNo Then Exit Sub
        i = ScalarToInt("cc_Setting", "RecID", String.Format(" custid={0}", Me.txtCustCode.Text))
        If i > 0 Then
            MsgBox("This Customer Has Credit Agreement with TV. Plz Check and Clear Debit First.", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If Me.GridCustList.CurrentRow.Cells("APP").Value = "RAS" Then
            cmd.CommandText = ChangeStatus_ByID("CustomerList", "EX", Me.txtCustCode.Text)
        Else
            cmd.CommandText = "update CustomerList set App=replace(app,'RAS','') where RecID=" & CInt(Me.txtCustCode.Text)
        End If
        cmd.CommandText = cmd.CommandText & "; " & ChangeStatus_ByDK("OfficeID", "XX", String.Format("CustID={0}", Me.txtCustCode.Text)) & _
            "; " & ChangeStatus_ByDK("Cust_Detail", "XX", String.Format(" cat='Channel' and status='OK' and CustID={0}", Me.txtCustCode.Text)) & _
            "; " & ChangeStatus_ByDK("cust_Channel_level", "XX", String.Format("CustID={0}", Me.txtCustCode.Text))
        cmd.ExecuteNonQuery()
        RefreshCustList()
    End Sub

    Private Sub ChkXXOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkXXOnly.CheckedChanged
        RefreshCustList()
    End Sub

    Private Sub LblReinstate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblReinstate.LinkClicked
        If MsgBox("This Gonna Reinstate Selected Customer. Are You Sure?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, msgTitle) = vbNo Then Exit Sub
        Dim CustID As Integer = ScalarToInt("Cust_detail", "top 1 RecID", "CAT='Channel' and custId=" & Me.txtCustCode.Text & " order by recID desc")
        cmd.CommandText = ChangeStatus_ByID("Customer", "OK", Me.txtCustCode.Text) & _
            ";" & ChangeStatus_ByID("Cust_detail", "OK", CustID)
        cmd.ExecuteNonQuery()
        RefreshCustList()
    End Sub

    Private Sub LblMoreDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblMoreDetail.LinkClicked
        Dim f As Form = New CustDetail("|FT")
        f.ShowDialog()
    End Sub
    Private Sub TxtEmail_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEmail.Enter
        If Me.TxtEmail.Text.Length > 4 AndAlso Me.TxtEmail.Text.Substring(0, 4) = "Plz " Then
            Me.TxtEmail.Text = ""
            Me.TxtEmail.ForeColor = Color.Black
        End If
    End Sub

    Private Sub CmbChannel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbChannel.SelectedIndexChanged
        If Me.CmbChannel.Text <> "TA" And myStaff.SICode <> "SYS" Then
            Me.LckLblMoreDetail.Enabled = False
        Else
            Me.LckLblMoreDetail.Enabled = True
        End If
    End Sub

    Private Sub GridCustList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridCustList.CellContentClick

    End Sub
End Class