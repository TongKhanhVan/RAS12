Imports SharedFunctions.MySharedFunctionsWzConn
Public Class CustDetail
    Private WhatAction As String
    Private OIDUsage As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private MyCust As New objCustomer
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal parWhatAction As String)
        Dim i As Int16 = InStr(parWhatAction, "|")
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        If i = 0 Then
            WhatAction = parWhatAction
            OIDUsage = ""
        Else
            WhatAction = parWhatAction.Substring(0, i - 1)
            OIDUsage = parWhatAction.Substring(i)
        End If
    End Sub

    Private Sub CustDetail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MyCust.CustID = 0
        Me.Dispose()
    End Sub
    Private Sub CustDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        Dim strFilter As String = "Select VAL from MISC where cat='Channel' and val in " & myStaff.CAccess
        If WhatAction <> "" Then
            Me.TabControl1.SelectTab("Tab" & WhatAction)
        End If
        LoadCmb_MSC(Me.CmbChannel, strFilter)
        LoadCmb_MSC(Me.CmbChannelFilter, strFilter)
        If OIDUsage = "SMS" Or WhatAction.ToUpper = "EMAIL" Then
            Me.CmbChannelFilter.Text = "TA"
            Me.CmbChannelFilter.Enabled = False
        End If
        If myStaff.CAccess.Length > 7 Then Me.CmbChannelFilter.Items.Add("") 'phai la sales GSA moi dc doi channel
        LoadGridCust("")
        GenALList_cmbEmailType()
    End Sub

    Private Sub LoadCurrentCA()

        Dim strSQL As String
        strSQL = String.Format("select RecID, Cat, Val from cust_Detail where status='OK' and left(VAL,4)='TAY3' and CustID={0}", MyCust.CustID)
        Me.GridCA.DataSource = GetDataTable(strSQL)
        Me.GridCA.Columns("RecID").Visible = False
        Me.GridCA.Columns("CAT").Width = 36
        Me.GridCA.Columns("VAL").Width = 32
        Me.LblDeleteCA.Visible = False
    End Sub
    Private Sub LoadCurrentAL()
        Dim ALList As String = ""
        ALList = FieldToList("VAL", "Cust_Detail", Conn, "status+CAT='OKAL' and CustID=" & Me.GridCustList.CurrentRow.Cells("RecID").Value)
        For i As Int16 = 0 To Me.ChkAL.Items.Count - 1
            Me.ChkAL.SetItemChecked(i, False)
        Next
        Me.ChkAll.Checked = False
        If InStr(ALList, "YY") > 0 Then
            Me.ChkAll.Checked = True
        Else
            For i As Int16 = 0 To Me.ChkAL.Items.Count - 1
                If InStr(ALList, Me.ChkAL.Items(i).ToString) > 0 Then
                    Me.ChkAL.SetItemChecked(i, True)
                End If
            Next
        End If
    End Sub
    Private Sub LoadGridEmail()
        Dim r As Int16
        Dim dTable As DataTable = GetDataTable("select VAL1, VAL from Cust_Detail where status+CAT='OKEMAIL' and CustID=" & _
                                               Me.GridCustList.CurrentRow.Cells("Recid").Value)
        Me.GridEmail.Rows.Clear()
        For i As Int16 = 0 To dTable.Rows.Count - 1
            Me.GridEmail.Rows.Add()
            r = Me.GridEmail.Rows.Count - 2
            Me.GridEmail.Item(0, r).Value = dTable.Rows(i)("VAL1")
            Me.GridEmail.Item(1, r).Value = dTable.Rows(i)("VAL")
        Next
    End Sub
    Private Sub LoadGridCust(ByVal pDK As String)

        Dim strSQL As String
        strSQL = String.Format("select RecID, CustShortName, CustFullName from CustomerList where city='{0}' and status='OK'" & _
            " and custShortName <> CustFullName and RecID ", MySession.City)
        If Me.CmbChannelFilter.Text = "" Then
            strSQL = String.Format("{0} not in (select Custid from cust_Detail where status+CAT='OKChannel')", strSQL)

        Else
            strSQL = String.Format("{0} in (select Custid from cust_Detail where status+cat='OKChannel' and " & _
                " VAL='{1}')", strSQL, Me.CmbChannelFilter.Text)
        End If
        strSQL = strSQL & pDK
        Me.GridCustList.DataSource = GetDataTable(strSQL)
        Me.GridCustList.Columns("RecID").Visible = False
        Me.GridCustList.Columns("CustShortName").Width = 75
        Me.GridCustList.Columns("CustFullName").Width = 256
        Me.LblUpdateOID.Visible = False
        Me.LblApproveOID.Visible = False
        Me.LblUpdateOID.Visible = False
        Me.LblUpdateAL.Visible = False
        Me.LblUpdateChannel.Visible = False
        Me.LblUpdateEmail.Visible = False
        MyCust.CustID = 0
    End Sub
    Private Sub LoadGridOffice()

        Dim strSQL As String, pStatus As String
        pStatus = IIf(Me.ChkQOnly.Checked, "QQ", "OK")
        strSQL = String.Format("select recID, OfficeID, OID_Usage from OfficeID where custId={0} and status='{1}' order by OfficeID", _
             Me.GridCustList.CurrentRow.Cells("Recid").Value, pStatus)
        Me.GridOffice.DataSource = GetDataTable(strSQL)
        Me.GridOffice.Columns("RecID").Visible = False
        Me.LblRemoveOID.Visible = False
        Me.LblApproveOID.Visible = False
        Me.LblAddUsage.Visible = False
    End Sub
    Private Sub GridCustList_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCustList.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        MyCust.CustID = Me.GridCustList.CurrentRow.Cells("RecID").Value
        Me.TxtOfficeID.Text = ""
        Me.CmbChannel.Text = defOldValue("VAL", "Cust_Detail", "", "Channel")
        Me.LblUpdateOID.Visible = True
        Me.LblUpdateAL.Visible = True
        Me.LblUpdateChannel.Visible = True
        Me.LblUpdateEmail.Visible = True
        LoadGridOffice()
        LoadGridEmail()
        LoadCurrentAL()
        'LoadCurrentCA()
    End Sub
    Private Function defOldValue(ByVal pField As String, ByVal pTable As String, ByVal pDK As String, ByVal pCat As String) As String
        Dim StrDK As String = " where status='OK' and custid=" & Me.GridCustList.CurrentRow.Cells("Recid").Value & " and cat='" & pCat & "'"
        If pDK <> "" Then
            StrDK = StrDK & " and " & pField & "='" & pDK & "'"
        End If
        Return ScalarToString(pTable, pField, StrDK)
    End Function
    Private Function defOldOID() As String
        Dim KQ As String = ""
        KQ = FieldToList("OfficeID", "OfficeID", Conn, "status='OK' and custid=" & Me.GridCustList.CurrentRow.Cells("Recid").Value)
        Return KQ
    End Function
    Private Function OIDExist() As Boolean
        Dim KQ As Boolean = False, tmpCustID As Integer
        tmpCustID = ScalarToInt("OfficeID", "CustID", "status='OK' and officeid='" & Me.TxtOfficeID.Text & "'")
        If tmpCustID <> 0 AndAlso tmpCustID <> Me.GridCustList.CurrentRow.Cells("Recid").Value Then
            KQ = True
        End If
        Return KQ
    End Function
    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateOID.LinkClicked
        Dim OldValue As String
        If OIDExist() Then
            MsgBox("This OID is known for other Customer. Plz check", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If Me.TxtOfficeID.Text.Length = 6 And OIDUsage = "SMS" Then
            MsgBox("Invalid OID Usage Input. Plz check", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        If OIDUsage = "FT" And InStr("PPD_PSP_DDB", MyCust.DelayType) = 0 Then
            MsgBox("No OID Accepted for Customer W/O Credit Agreement Wz TVA. Plz check", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        OldValue = defOldOID()
        If (Me.TxtOfficeID.Text.Length = 9 Or Me.TxtOfficeID.Text.Length = 6) And _
            InStr(OldValue, Me.TxtOfficeID.Text) = 0 Then
            cmd.CommandText = "insert OfficeID (CustID, OfficeID, OID_Usage, FstUser) values (@CustID, @OfficeID, @OID_Usage, @FstUser)"
            cmd.Parameters.Clear()
            cmd.Parameters.Add("@CustID", SqlDbType.Int).Value = Me.GridCustList.CurrentRow.Cells("Recid").Value
            cmd.Parameters.Add("@OfficeID", SqlDbType.VarChar).Value = Me.TxtOfficeID.Text
            cmd.Parameters.Add("@OID_Usage", SqlDbType.VarChar).Value = OIDUsage
            cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
            cmd.ExecuteNonQuery()
            LoadGridOffice()
        End If
        Me.LblUpdateOID.Visible = False
    End Sub

    Private Sub CmbChannelFilter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbChannelFilter.SelectedIndexChanged
        LoadGridCust("")
        If Me.CmbChannelFilter.Text = "TA" Or Me.CmbChannelFilter.Text = "" Or myStaff.SICode = "SYS" Then
            Me.GrpCA.Visible = True
            Me.TabControl1.Enabled = True
        Else
            Me.GrpCA.Visible = False
            Me.TabControl1.Enabled = False
        End If
    End Sub

    Private Sub GridOffice_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridOffice.CellContentClick
        Me.LblRemoveOID.Visible = True
        Me.LblAddUsage.Visible = True
        If Me.ChkQOnly.Checked Then Me.LblApproveOID.Visible = True
    End Sub

    Private Sub LblRemove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblRemoveOID.LinkClicked
        Dim CurrUsage As String, strSQL As String
        CurrUsage = ScalarToString("officeID", "OID_Usage", " Recid=" & Me.GridOffice.CurrentRow.Cells("RecID").Value)
        strSQL = UpdateLogFile("OfficeID", "DEL USGE", "Cust:" & Me.GridCustList.CurrentRow.Cells("CustShortName").Value, _
            "OID:" & Me.GridOffice.CurrentRow.Cells("OfficeID").Value, "DEL:" & OIDUsage, "", "", "", "", "")

        If InStr(Me.GridOffice.CurrentRow.Cells("OID_Usage").Value, OIDUsage) > 0 Then
            strSQL = strSQL & "; update OfficeID set OID_Usage=replace(OID_Usage,'" & OIDUsage & "','')" & _
                    " where recID=" & Me.GridOffice.CurrentRow.Cells("RecID").Value
            Do
                CurrUsage = ScalarToString("officeID", "OID_Usage", " Recid=" & Me.GridOffice.CurrentRow.Cells("RecID").Value)
                If CurrUsage = "" Then Exit Do
                If CurrUsage.Substring(0, 1) = "_" Or Strings.Right(CurrUsage, 1) = "_" Then
                    If CurrUsage.Substring(0, 1) = "_" Then
                        CurrUsage = CurrUsage.Substring(1)
                    End If
                    If Strings.Right(CurrUsage, 1) = "_" Then
                        CurrUsage = Strings.Left(CurrUsage, CurrUsage.Length - 1)
                    End If
                    strSQL = strSQL & "; update OfficeID set OID_Usage='" & CurrUsage & "' where recID=" & _
                        Me.GridOffice.CurrentRow.Cells("RecID").Value
                Else
                    Exit Do
                End If
            Loop
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()
            LoadGridOffice()
        End If
        If CurrUsage = "" Then
            cmd.CommandText = strSQL & ";" & ChangeStatus_ByID("OfficeID", "XX", Me.GridOffice.CurrentRow.Cells("RecID").Value)
        End If
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub ChkQOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkQOnly.CheckedChanged
        LoadGridOffice()
    End Sub
    Private Sub LblApprove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblApproveOID.LinkClicked
        cmd.CommandText = "update OfficeID set Status='OK', ApproveBy=@ApproveBy, ApproveOn=getdate() where recID=@recID"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@ApproveBy", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@recID", SqlDbType.Int).Value = Me.GridOffice.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridOffice()
    End Sub
    Private Sub LblSearch_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSearch.LinkClicked
        LoadGridCust(" and CustShortName like '%" & Me.TxtSearch.Text & "%'")
    End Sub

    Private Sub GenALList_cmbEmailType()
        Dim fList As String

        Me.eMailType.Items.Clear()
        Me.LstOIDUsage.Items.Clear()
        Me.ChkAL.Items.Clear()

        fList = FieldToList("VAL", "MISC", Conn, " cat='EMAILTYPE'")
        For i As Int16 = 0 To UBound(fList.Split(","))
            Me.eMailType.Items.Add(fList.Split(",")(i))
        Next

        fList = FieldToList("VAL", "MISC", Conn, " CAT='OIDUSAGE'")
        For i As Int16 = 0 To UBound(fList.Split(","))
            If fList.Split(",")(i) = OIDUsage Then
                Me.LstOIDUsage.Items.Add(fList.Split(",")(i), True)
            Else
                Me.LstOIDUsage.Items.Add(fList.Split(",")(i), False)
            End If
        Next

        Dim dTable As DataTable = GetDataTable(myStaff.TVA)
        For i As Int16 = 0 To dTable.Rows.Count - 1
            Me.ChkAL.Items.Add(dTable.Rows(i)("VAL"))
        Next
        Me.LstOIDUsage.Visible = False
    End Sub
    Private Sub ChkAll_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkAll.Click
        For i As Int16 = 0 To Me.ChkAL.Items.Count - 1
            Me.ChkAL.SetItemChecked(i, Me.ChkAll.Checked)
        Next
    End Sub

    Private Sub ChkAL_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAL.SelectedIndexChanged
        Me.ChkAll.Checked = False
    End Sub

    Private Sub LblUpdateChannel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateChannel.LinkClicked
        Dim OldValue As String
        OldValue = defOldValue("VAL", "Cust_Detail", "", "Channel")
        If OldValue <> Me.CmbChannel.Text Then
            MyCust.InsertCustDetail(Me.GridCustList.CurrentRow.Cells("Recid").Value, "Channel", Me.CmbChannel.Text, True)
        End If
        Me.LblUpdateChannel.Visible = False
    End Sub

    Private Sub LblUpdateAL_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateAL.LinkClicked
        If Me.ChkAll.Checked Then
            MyCust.InsertCustDetail(Me.GridCustList.CurrentRow.Cells("Recid").Value, "AL", IIf(Me.CmbChannelFilter.Text = "CS", "TS", "YY"), True)
        Else
            cmd.CommandText = ChangeStatus_ByDK("cust_Detail", "XX", "custId=" & Me.GridCustList.CurrentRow.Cells("Recid").Value & " and cat='AL' and status='OK'")
            cmd.ExecuteNonQuery()
            For i As Int16 = 0 To Me.ChkAL.Items.Count - 1
                If Me.ChkAL.GetItemChecked(i) = True Then
                    MyCust.InsertCustDetail(Me.GridCustList.CurrentRow.Cells("Recid").Value, "AL", IIf(Me.CmbChannelFilter.Text = "CS", "TS", Me.ChkAL.Items(i).ToString), False)
                End If
            Next
        End If
        Me.LblUpdateAL.Visible = False
    End Sub
    Private Sub LblUpdateEmail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateEmail.LinkClicked
        Dim tmpEmail As String, eMailTypeList As String = "", strSQL As String
        For i As Int16 = 0 To Me.GridEmail.RowCount - 2
            If InStr(eMailTypeList, Me.GridEmail.Item(0, i).Value) = 0 Then
                eMailTypeList = eMailTypeList & "_" & Me.GridEmail.Item(0, i).Value
            Else
                MsgBox("Invalid Input", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        Next
        strSQL = ChangeStatus_ByDK("cust_Detail", "XX", " status='OK' and custID=" & Me.GridCustList.CurrentRow.Cells("Recid").Value & " and cat='EMAIL'")
        For i As Int16 = 0 To Me.GridEmail.Rows.Count - 2
            tmpEmail = Me.GridEmail.Item(1, i).Value
            If tmpEmail <> "" Then
                If emailValid(tmpEmail) Then
                    tmpEmail = tmpEmail.Replace("--", "")
                    tmpEmail = tmpEmail.Replace(",", ";")
                    strSQL = strSQL & "; Insert cust_Detail (custID, CAT, VAL1, VAL, FstUser) values (" & Me.GridCustList.CurrentRow.Cells("Recid").Value & _
                         ",'EMAIL','" & Me.GridEmail.Item(0, i).Value & "','" & tmpEmail & "','" & myStaff.SICode & "')"
                Else
                    MsgBox("Invalid Email Address. Plz Check", MsgBoxStyle.Critical, msgTitle)
                End If
            End If
        Next
        cmd.CommandText = strSQL
        cmd.ExecuteNonQuery()
        Me.LblUpdateEmail.Visible = False
    End Sub
    Private Function emailValid(ByVal pEmail As String) As Boolean
        Dim soColon As Int16, soAt As Int16, tmpRec As Integer = 0
        For i As Int16 = 0 To pEmail.Length - 1
            If pEmail.Substring(i, 1) = "@" Then
                soAt = soAt + 1
            ElseIf pEmail.Substring(i, 1) = ";" Then
                soColon = soColon + 1
            End If
        Next
        If soAt = 0 Or soAt - 1 <> soColon Or soAt > 3 Then
            Return False
        ElseIf InStr(pEmail, ",") > 0 Then
            Return False
        End If
        Return True
    End Function

    Private Sub TabAL_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        TabAL.Enter, TabChannel.Enter, TabEmail.Enter, TabOID.Enter
        Dim tb As TabPage = CType(sender, TabPage)
        If WhatAction <> "" And InStr(tb.Name, WhatAction) = 0 Then
            Me.TabControl1.SelectTab("Tab" & WhatAction)
        End If
    End Sub
    Private Sub LblAddUsage_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAddUsage.LinkClicked
        If InStr(Me.GridOffice.CurrentRow.Cells("OID_Usage").Value, OIDUsage) = 0 Then
            cmd.CommandText = "update OfficeID set OID_Usage=OID_Usage+ '_" & OIDUsage & "' where recID=" & _
                    Me.GridOffice.CurrentRow.Cells("RecID").Value
            cmd.CommandText = cmd.CommandText & ";" & UpdateLogFile("OfficeID", "Add USGE", "Cust:" & Me.GridCustList.CurrentRow.Cells("CustShortName").Value, _
                "OID:" & Me.GridOffice.CurrentRow.Cells("OfficeID").Value, "Add:" & OIDUsage, "", "", "", "", "")
            cmd.ExecuteNonQuery()
            LoadGridOffice()
        End If
    End Sub
    Private Sub CmbSBU_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbSBU.SelectedIndexChanged
        Me.CmbAL.Items.Clear()
        If Me.CmbSBU.Text = "EDU" Then
            Me.CmbAL.Items.Add("XX")
        ElseIf Me.CmbSBU.Text = "GSA" Then
            LoadCmb_MSC(Me.CmbAL, myStaff.GSA)
        ElseIf Me.CmbSBU.Text = "TVS" Then
            LoadCmb_MSC(Me.CmbAL, myStaff.ALList)
        End If
    End Sub

    Private Sub LblUpdateCA_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdateCA.LinkClicked
        If Me.TxtCA.Text = "" Then Exit Sub
        MyCust.InsertCustDetail(MyCust.CustID, "TAY3:" & Me.CmbSBU.Text & Me.CmbAL.Text, Me.TxtCA.Text, True)
        LoadCurrentCA()
        Me.TxtCA.Text = ""
        Me.LblUpdateCA.Visible = False
    End Sub

    Private Sub GridCA_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCA.CellContentClick
        Me.LblDeleteCA.Visible = True
    End Sub

    Private Sub TxtCA_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCA.TextChanged
        Me.LblUpdateCA.Visible = True
    End Sub

    Private Sub LblDeleteCA_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDeleteCA.LinkClicked
        cmd.CommandText = ChangeStatus_ByID("Cust_detail", "XX", Me.GridCA.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        LoadCurrentCA()
    End Sub
End Class