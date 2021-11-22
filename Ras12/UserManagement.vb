Imports SharedFunctions.MySharedFunctions
Public Class UserManagement
    Private WhoIs As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private SQLRight As String
    Public Sub New(ByVal parWhoIs As String)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        WhoIs = parWhoIs
    End Sub
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub UpdateRights()
        Dim Msg As String

        cmd.CommandText = ChangeStatus_ByDK("tblright ", "XX", "SIcode='" & Me.txtUserManagementSIcode.Text & "'")
        cmd.ExecuteNonQuery()

        SQLRight = ""
        DuyetNode2FindChecked(Me.TreeRight.Nodes)
        cmd.CommandText = SQLRight
        If SQLRight <> "" Then cmd.ExecuteNonQuery()

        UpdateCanSIas()
        UpdateChannelAccess()

        Msg = "Changes Updated For " & Me.txtUserManagementSIcode.Text
        If Me.txtUserManagementSIcode.Text.Contains("**") Then
            Msg = Msg & vbCrLf & "Rights for All Users of this Group Have Been Updated Accordingly"
            RefreshUserRightOfGroup(Me.txtUserManagementSIcode.Text)
        End If
        MsgBox(Msg, MsgBoxStyle.Information, msgTitle)
        Me.txtUserManagementSIcode.Text = ""
        Me.txtUserManagementSIcode.Focus()

    End Sub
    Private Sub UpdateChannelAccess()
        Dim tmpUID As Integer, tmpCAccess As String = ""
        tmpUID = ScalarToInt("tblUser", "RecID", String.Format(" status<>'XX' and siCode='{0}'", Me.txtUserManagementSIcode.Text))
        For i As Int16 = 0 To Me.LstChannelAccess.Items.Count - 1
            If Me.LstChannelAccess.GetItemChecked(i) Then
                tmpCAccess = tmpCAccess & "_" & Me.LstChannelAccess.Items(i).ToString
            End If
        Next
        If tmpCAccess.Length > 0 Then tmpCAccess = tmpCAccess.Substring(1)
        cmd.CommandText = String.Format("update UserAccess set status='XX', LstUser='{0}', LstUpdate=getdate() where UserID={1} " & _
            " and APP='RAS' and cat='ChannelAccess'; insert UserAccess (APP, CAT, fstUser, UserID, VAL) values ('RAS','ChannelAccess'," & _
            "'{0}',{1},'{2}')", myStaff.SICode, tmpUID, tmpCAccess)
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub UpdateCanSIas()
        Dim DomainAccess As String = "EDU", tmpUID As Integer, tmpAAccess As String = ""
        tmpUID = ScalarToInt("tblUser", "RecID", "status<>'XX' and siCode='" & Me.txtUserManagementSIcode.Text & "'")

        If Me.ChkGSA.Checked Then DomainAccess = DomainAccess & "_GSA"
        If Me.ChkTVS.Checked Then DomainAccess = DomainAccess & "_TVS"
        cmd.CommandText = String.Format("update UserAccess set status='XX', LstUser='{0}', LstUpdate=getdate() where UserID={1} " & _
               " and APP='RAS' and cat in ('DomainAccess','ALAccess'); insert UserAccess (APP, CAT, VAL, FstUser, UserID) values " & _
               "('RAS','DomainAccess','{2}','{0}',{1})", myStaff.SICode, tmpUID, DomainAccess)
        cmd.ExecuteNonQuery()

        If Me.ChckAccessYY.Checked Then
            cmd.CommandText = String.Format("insert UserAccess (APP, CAT, VAL, fstUser, UserID) values ('RAS','ALAccess','YY'," & _
                "'{0}',{1})", myStaff.SICode, tmpUID)
            cmd.ExecuteNonQuery()
        Else
            For i As Int16 = 0 To Me.LstAL.Items.Count - 1
                If Me.LstAL.GetItemChecked(i) Then
                    tmpAAccess = tmpAAccess & "_" & Me.LstAL.Items(i).ToString
                End If
            Next
            If tmpAAccess.Length > 0 Then tmpAAccess = tmpAAccess.Substring(1)
            cmd.CommandText = String.Format("insert UserAccess (APP, CAT, FstUser, UserID, VAL) values ('RAS','ALAccess','{0}',{1},'{2}')", _
                myStaff.SICode, tmpUID, tmpAAccess)
            cmd.ExecuteNonQuery()
        End If
    End Sub
    Private Sub RefreshUserRightOfGroup(ByVal parGrp As String)
        Dim dtbl As DataTable = GetDataTable(String.Format("Select SICode from tblUser where status<>'XX' and Template='{0}'", parGrp))
        For i As Int16 = 0 To dtbl.Rows.Count - 1
            CloneRight(dtbl.Rows(i)("SICode"), parGrp)
        Next
    End Sub
    Private Sub CloneRight(ByVal parUI As String, ByVal parGroup As String)
        cmd.CommandText = String.Format("Delete from tblRight where SICode='{0}'" & _
            "; insert into tblRight (Sicode, object, subobject) select '{0}'" & _
            " as SIcode, object, SubObject from tblRight where SICode='{1}'", parUI, parGroup)
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub UpdateCurrentRight(ByVal pSICode As String)
        Dim dtbl As DataTable = GetDataTable(String.Format("select recID, SubObject from tblright where status ='OK' and sicode='{0}' order by recID", pSICode))
        Try
            For i As Int16 = 0 To dtbl.Rows.Count - 1
                'FindNode(dtbl.Rows(i)("SubObject"), Me.TreeRight.Nodes).Checked = True
                CheckTreeNode(dtbl.Rows(i)("SubObject"), True)
            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Function CheckTreeNode(strNodeName As String, blnCheckChildNodes As Boolean) As Boolean
        Dim arrNodes As TreeNode()
        arrNodes = TreeRight.Nodes.Find(strNodeName, True)
        For Each n As TreeNode In arrNodes
            n.Checked = True
        Next
        Return True
    End Function
    Private Function FindNode(ByVal pNodeName As String, ByVal pTree As TreeNodeCollection) As TreeNode
        Dim n As TreeNode
        For Each n In pTree
            If n.Name = pNodeName Or n.Text = pNodeName Then
                Return n
                Exit Function
            Else
                If n.Nodes.Count > 0 Then
                    Dim t As TreeNode = FindNode(pNodeName, n.Nodes)
                    If Not t Is Nothing Then
                        Return t
                        Exit Function
                    End If
                End If
            End If
        Next
        Return Nothing
    End Function
    Private Sub DuyetNode2FindChecked(ByVal pnode As TreeNodeCollection)
        Dim n As TreeNode
        For Each n In pnode
            If n.Checked Then
                SQLRight = SQLRight & "; insert into tblRight (SICode, Object, SubObject) values ('" & _
                    Me.txtUserManagementSIcode.Text & "','"
                If InStr("PAD_BAR", n.Name.ToUpper.Substring(0, 3)) > 0 Then
                    SQLRight = SQLRight & "Menu"
                    SQLRight = SQLRight & "','" & n.Name & "')"
                Else
                    SQLRight = SQLRight & n.Name
                    SQLRight = SQLRight & "','" & n.Text & "')"
                End If

            End If
            If n.Nodes.Count > 0 Then
                DuyetNode2FindChecked(n.Nodes)
            End If
        Next
    End Sub
    Private Sub ResetNodes(ByVal pnode As TreeNodeCollection)
        Dim n As TreeNode
        For Each n In pnode
            If n.ForeColor <> Color.Red Then
                n.Checked = False
            End If
            If n.Nodes.Count > 0 Then
                ResetNodes(n.Nodes)
            End If
        Next
    End Sub

    Private Sub TreeRight_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeRight.AfterCheck
        If e.Node.Checked Then
            e.Node.Collapse()
        Else
            e.Node.Expand()
        End If
    End Sub
    Private Sub CmdUserManagementCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdUserManagementCreate.Click
        Dim tmpUID As Integer, tmpNewPSW As String = GenDefaultPSW(Me.txtUserManagementSIcode.Text, "R12")
        Dim strTemplate As String = IIf(WhoIs = "SYS", Me.CmbTemplate.Text, myStaff.SupOf)
        tmpUID = ScalarToInt("tblUser", "RecID", "SICode='" & Me.txtUserManagementSIcode.Text & "'")
        If tmpUID > 0 Then Exit Sub
        If WhoIs <> "SYS" And myStaff.SupOf = "" Then Exit Sub

        cmd.CommandText = "insert into tbluser (SICode, Template, FstUser, SIname, PSW, counter, location) values (@SICode,@Template,@FstUser, @SIName, @PSW, @counter, @location)"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@SICode", SqlDbType.VarChar).Value = Me.txtUserManagementSIcode.Text.Replace("--", "")
        cmd.Parameters.Add("@Template", SqlDbType.VarChar).Value = strTemplate
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@SIname", SqlDbType.VarChar).Value = Me.CmbUsername.Text
        cmd.Parameters.Add("@PSW", SqlDbType.VarChar).Value = HashToFixedLen(tmpNewPSW)
        If WhoIs <> "SYS" Then
            cmd.Parameters.Add("@counter", SqlDbType.VarChar).Value = myStaff.Counter
            cmd.Parameters.Add("@location", SqlDbType.VarChar).Value = myStaff.Location
        Else
            cmd.Parameters.Add("@counter", SqlDbType.VarChar).Value = Me.CmbCounter.Text
            cmd.Parameters.Add("@location", SqlDbType.VarChar).Value = IIf(Me.ChkHAN.Checked, "HAN", Me.CmbLocation.Text)
        End If
        cmd.ExecuteNonQuery()
        tmpUID = ScalarToInt("tblUser", "RecID", String.Format("SICode='{0}'", Me.txtUserManagementSIcode.Text))

        If WhoIs <> "SYS" Then
            cmd.CommandText = String.Format("insert UserAccess (UserID, App, CAT, VAL, fstUser) select {0},'RAS','ChannelAccess',VAL,'{1}'" & _
                    " from UserAccess where Cat='ChannelAccess' and UserID={2} and status<>'XX'", tmpUID, myStaff.SICode, myStaff.UID)
            cmd.ExecuteNonQuery()

            UpdateCanSIas()

            cmd.CommandText = String.Format("insert tblright (SIcode, Object, SubObject) select '{0}', Object, SubObject from tblright where sicode='{1}'", _
                Me.txtUserManagementSIcode.Text, myStaff.SupOf)
            cmd.ExecuteNonQuery()
        Else
            UpdateRights()
        End If
        MsgBox("User " & Me.txtUserManagementSIcode.Text & " Has Been Created. Initial PSW is " & tmpNewPSW & ". Ask Him/Her to Change PSW", MsgBoxStyle.Information, msgTitle)
    End Sub
    Private Sub CmdUserManagementDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdUserManagementDelete.Click
        Dim myAnswer As Int16, strDK As String = String.Format(" where SIcode='{0}'", Me.txtUserManagementSIcode.Text)
        Dim strSQL As String, BeingDeletedID As Integer

        myAnswer = MsgBox("Wanna Delete This User Permanently?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, msgTitle)
        If myAnswer = vbNo Then Exit Sub

        If WhoIs <> "SYS" Then
            If myStaff.SupOf <> Me.CmbTemplate.Text Then
                MsgBox("You Dont Have Enough Right To Perform Action On " & Me.txtUserManagementSIcode.Text, MsgBoxStyle.Information, msgTitle)
                Exit Sub
            End If
        End If

        BeingDeletedID = ScalarToInt("tblUser", "RecID", strDK)
        strSQL = UpdateLogFile("tblUser", "Del User", Me.txtUserManagementSIcode.Text, Me.CmbUsername.Text, "", "", "")
        strSQL = strSQL & "; Delete from tblUser " & strDK
        strSQL = strSQL & "; delete from tblright " & strDK & "; delete from UserAccess where userid=" & BeingDeletedID

        MsgBox(Me.txtUserManagementSIcode.Text & " Has Been Deleted!", MsgBoxStyle.Information, msgTitle)
        If WhoIs = "S**" Then strSQL = strSQL & "; " & ChangeStatus_ByDK("SI_Agent", "XX", strDK)
        cmd.CommandText = strSQL
        cmd.ExecuteNonQuery()

        Me.txtUserManagementSIcode.Text = ""
        Me.txtUserManagementSIcode.Focus()

    End Sub

    Private Sub CmdChange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdChange.Click
        If Me.txtUserManagementSIcode.Text.Length < 3 Or Me.txtUserManagementSIcode.Text.Length > 4 Then Exit Sub
        Me.txtUserManagementSIcode.Enabled = False
        If WhoIs = "SYS" Then
            UpdateRights()
        Else
            If myStaff.SupOf <> Me.CmbTemplate.Text Then
                MsgBox("You Dont Have Enough Right To Perform Action On " & Me.txtUserManagementSIcode.Text, MsgBoxStyle.Information, msgTitle)
                Exit Sub
            End If
            UpdateCanSIas()
            UpdateChannelAccess()
        End If
        Me.txtUserManagementSIcode.Enabled = True
    End Sub
    Private Sub txtUserManagementSIcode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserManagementSIcode.LostFocus
        Dim tmpSIcode As String = "", tmpUID As Integer
        Dim dtbl As DataTable = GetDataTable(String.Format("Select * from tblUser where SIcode='{0}'", Me.txtUserManagementSIcode.Text))
        Dim Ustatus As String = ""
        If Me.txtUserManagementSIcode.Text = "" Then Exit Sub
        If InStr(Me.txtUserManagementSIcode.Text, "**") > 0 Then
            Me.CmbTemplate.Text = ""
            Me.CmbTemplate.Enabled = False
        Else
            If WhoIs = "SYS" Then Me.CmbTemplate.Enabled = True
        End If

        Me.CmdUserManagementCreate.Enabled = False
        If dtbl.Rows.Count > 0 Then
            Me.CmbUsername.DropDownStyle = ComboBoxStyle.DropDown
            tmpSIcode = dtbl.Rows(0)("SICode")
            tmpUID = dtbl.Rows(0)("RecID")
            Ustatus = dtbl.Rows(0)("Status")
            Me.CmbUsername.Text = dtbl.Rows(0)("SIName")
            Me.CmbTemplate.Text = dtbl.Rows(0)("Template")
            Me.CmdUserManagementCreate.Enabled = True
            CmbLocation.SelectedIndex = CmbLocation.FindStringExact(dtbl.Rows(0)("Location"))
            CmbCounter.SelectedIndex = CmbCounter.FindStringExact(dtbl.Rows(0)("Counter"))
        Else
            Me.CmbUsername.DropDownStyle = ComboBoxStyle.DropDownList
            If Me.ChkHAN.Checked Then
                LoadCmb_MSC(Me.CmbUsername, "Select ShortName as VAL from [42.117.5.86].lib.dbo.dsnv_HAN " & _
                            " where ShortName not in (select SIName from tblUser where status<>'XX') ")
            Else
                LoadCmb_MSC(Me.CmbUsername, "Select ShortName as VAL from lib.dbo.StaffList where ShortName not in (select SIName from tblUser where status<>'XX') ")
            End If
            Me.CmbUsername.Items.Add("Template 4 Counter")
            Me.CmbUsername.Items.Add("Template for NonAir")
            Me.CmbUsername.Items.Add("Template 4 Operation")
            Me.CmdUserManagementCreate.Enabled = False
        End If

        LoadTreeRight()

        GenALList(tmpUID)
        GenChannelList(tmpUID)

        If tmpSIcode <> "" Then
            Me.CmdUserManagementCreate.Enabled = IIf(Ustatus = "XX", True, False)
            Me.CmbUsername.Enabled = False
            UpdateCurrentRight(tmpSIcode)
        Else
            Me.CmbUsername.Enabled = True
            Me.CmdUserManagementCreate.Enabled = True
        End If
        Me.CmdUserManagementDelete.Enabled = Not Me.CmbUsername.Enabled
        Me.CmdChange.Enabled = Not Me.CmbUsername.Enabled
        Me.CmdResetPSW.Enabled = Me.CmdChange.Enabled
        Me.LblChangeCounter.Enabled = Me.CmdChange.Enabled
    End Sub
    Private Sub GenALList(ByVal pUID As Integer)
        Dim tmpDomainAccess As String, tmpI As Integer
        Dim tmpDK As String = String.Format("and UserID={0}", pUID)
        Dim dtbl As DataTable = GetDataTable(myStaff.TVA & " order by AL")
        Me.LstAL.Items.Clear()
        For i As Int16 = 0 To dtbl.Rows.Count - 1
            Me.LstAL.Items.Add(dtbl.Rows(i)("VAL"))
        Next
        If pUID = 0 Then Exit Sub
        tmpI = ScalarToInt("UserAccess", "RecID", "Cat='ALAccess' and status<>'XX' and vaL='YY' " & tmpDK)
        If tmpI > 0 Then
            Me.ChckAccessYY.Checked = True
        Else
            Me.ChckAccessYY.Checked = False
            dtbl = GetDataTable("select VAL from UserAccess where Cat='ALAccess' and status<>'XX' " & tmpDK)
            For j As Int16 = 0 To dtbl.Rows.Count - 1
                For i As Int16 = 0 To Me.LstAL.Items.Count - 1
                    If Me.LstAL.Items(i).ToString = dtbl.Rows(j)("VAL") Then
                        Me.LstAL.SetItemChecked(i, True)
                    End If
                Next

            Next
        End If
        tmpDomainAccess = ScalarToString("UserAccess", "VAL", "Cat='DomainAccess' and status<>'XX' " & tmpDK)
        Me.ChkGSA.Checked = IIf(InStr(tmpDomainAccess, "GSA") > 0, True, False)
        Me.ChkTVS.Checked = IIf(InStr(tmpDomainAccess, "TVS") > 0, True, False)
    End Sub

    Private Sub ChckAccessYY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChckAccessYY.CheckedChanged
        For i As Int16 = 0 To Me.LstAL.Items.Count - 1
            Me.LstAL.SetItemChecked(i, False)
        Next
        If Me.ChckAccessYY.Checked Then
            Me.LstAL.Enabled = False
        Else
            Me.LstAL.Enabled = True
        End If
    End Sub

    Private Sub UserManagement_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.BackColor = pubVarBackColor
        If WhoIs <> "SYS" Then
            'Me.TreeRight.ExpandAll()
            TreeRight.Enabled = False

            Me.CmbTemplate.Text = WhoIs
            Me.CmbTemplate.Enabled = False
            Me.ChkHAN.Enabled = False
        End If
        LoadCmb_MSC(Me.CmbLocation, "Location")
        LoadCmb_MSC(Me.CmbCounter, "Counter")

        lbkCcAccess.Visible = (myStaff.SupOf = "CS**")

    End Sub

    Private Sub GenChannelList(ByVal pUID As Integer)
        Dim dtbl As DataTable = GetDataTable("select VAL from MISC where cat='Channel' order by val")
        Me.LstChannelAccess.Items.Clear()
        For j As Int16 = 0 To dtbl.Rows.Count - 1
            If InStr(myStaff.CAccess, dtbl.Rows(j)("VAL")) > 0 Then
                Me.LstChannelAccess.Items.Add(dtbl.Rows(j)("VAL"))
            End If
        Next
        If pUID = 0 Then Exit Sub
        dtbl = GetDataTable(String.Format("select VAL from UserAccess where cat='ChannelAccess' and status<>'XX' and Userid={0}", pUID))
        For j As Int16 = 0 To dtbl.Rows.Count - 1
            For i As Int16 = 0 To Me.LstChannelAccess.Items.Count - 1
                If InStr(dtbl.Rows(j)("VAL"), Me.LstChannelAccess.Items(i).ToString) > 0 Then
                    Me.LstChannelAccess.SetItemChecked(i, True)
                End If
            Next
        Next
    End Sub
    Private Sub CmdResetPSW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdResetPSW.Click
        Dim tmpNewPSW As String = GenDefaultPSW(Me.txtUserManagementSIcode.Text, "R12")
        Dim strSQL As String
        strSQL = String.Format("Update tblUser set PSW='{0}' where sicode='{1}'", HashToFixedLen(tmpNewPSW), Me.txtUserManagementSIcode.Text)
        If WhoIs <> "SYS" Then
            If myStaff.SupOf <> Me.CmbTemplate.Text Then
                MsgBox("You Dont Have Enough Right To Perform Action On " & Me.txtUserManagementSIcode.Text, MsgBoxStyle.Information, msgTitle)
                Exit Sub
            End If
            strSQL = String.Format("{0} and template <>'' and Template='{1}'", strSQL, myStaff.SupOf)
        End If
        cmd.CommandText = strSQL
        cmd.ExecuteNonQuery()
        cmd.CommandText = UpdateLogFile("tblUser", "ResetPSW", Me.txtUserManagementSIcode.Text, Me.CmbUsername.Text, "", "", "", "", "", "")
        cmd.ExecuteNonQuery()
        MsgBox(tmpNewPSW & " Has Been Set for " & Me.txtUserManagementSIcode.Text, MsgBoxStyle.Information, msgTitle)
    End Sub

    Private Sub LbLoadForm_LinkClicked() ' KO XOA 3 CODE Sau
        'For Each t As Type In My.Application.GetType().Assembly.GetTypes()
        ' If UCase(t.BaseType.ToString) = "SYSTEM.WINDOWS.FORMS.FORM" Then
        ' Me.LstForm.Items.Add(t.Name)
        ' End If
        ' Next
    End Sub
    Private Sub LstForm_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim frm As Form = System.Activator.CreateInstance(Type.GetType("RAS12." & Me.LstForm.Text))
        'Me.LstControlsOfForm.Items.Clear()
        'AddControlToList(frm)
    End Sub
    Private Sub AddControlToList(ByVal pCtrl As Control)
        'If pCtrl.Name <> "" AndAlso InStr("CMD_LBL", pCtrl.Name.Substring(0, 3).ToUpper) > 0 Then
        ' Me.LstControlsOfForm.Items.Add(pCtrl.Name)
        ' ElseIf pCtrl.Controls.Count > 0 Then
        ' For Each i As Control In pCtrl.Controls
        ' AddControlToList(i)
        ' Next
        'End If
    End Sub
    Private Sub LoadTreeRight()
        Dim mnu As ToolStripDropDownItem, MnuI As ToolStripItem
        Me.TreeRight.Nodes.Clear()
        Me.TreeRight.Nodes.Add("TreeRoot", "CHECK TO PROHIBIT")
        Me.TreeRight.Nodes(0).ForeColor = Color.Red
        For Each mnu In frmMain.MenuStrip1.Items
            Me.TreeRight.Nodes("TreeRoot").Nodes.Add(mnu.Name, mnu.ToString)


            For Each MnuI In mnu.DropDownItems
                If InStr(MnuI.Name, "Separator") = 0 Then
                    If MnuI.Tag <> "0" Then
                        Me.TreeRight.Nodes("TreeRoot").Nodes(mnu.Name).Nodes.Add(MnuI.Name, MnuI.ToString)
                        If Not MnuI.ToolTipText Is Nothing AndAlso MnuI.ToolTipText.Substring(0, 1) = "1" Then
                            FindNode(MnuI.Name, Me.TreeRight.Nodes).Checked = True
                            FindNode(MnuI.Name, Me.TreeRight.Nodes).ForeColor = Color.Magenta
                        ElseIf Not MnuI.ToolTipText Is Nothing AndAlso MnuI.ToolTipText.Substring(0, 1) = "2" Then
                            FindNode(MnuI.Name, Me.TreeRight.Nodes).ForeColor = Color.Blue
                        ElseIf Not MnuI.ToolTipText Is Nothing AndAlso MnuI.ToolTipText.Substring(0, 1) = "0" Then
                            Dim frm As Form = System.Activator.CreateInstance(Type.GetType("RAS12." & MnuI.ToolTipText.Substring(2)))
                            AddControlToTree(frm, MnuI.Name, frm.Name)
                        End If
                    End If
                End If
            Next
        Next
        If myStaff.SICode = "SYS" Or myStaff.SICode = "NMK" Then Me.TreeRight.Nodes(0).Nodes("PadSystem").Nodes("BarUserManager").Checked = False
    End Sub
    Private Sub AddControlToTree(ByVal pCtrl As Control, ByVal NodeName As String, FormName As String)
        Dim i As Control
        If pCtrl.Name <> "" AndAlso InStr(pCtrl.Name.Substring(0, 3).ToUpper, "LCK") > 0 Then
            FindNode(NodeName, Me.TreeRight.Nodes).Nodes.Add(FormName, pCtrl.Name)
            'Me.LstControlsOfForm.Items.Add(pCtrl.Name)
        ElseIf pCtrl.Controls.Count > 0 Then
            For Each i In pCtrl.Controls
                AddControlToTree(i, NodeName, FormName)
            Next
        End If
    End Sub

    Private Sub Label7_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label7.DoubleClick
        cmd.CommandText = "update tblUser set status='XX' where SICode <>'SYS' and location <>'HAN' and SIName not in (select ShortName from Lib.dbo.StaffList)"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "delete from tblright where SICode in (select SICode from tblUser where status='XX')"
        cmd.ExecuteNonQuery()
    End Sub
    Private Sub LblChangeCounter_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblChangeCounter.LinkClicked
        If InStr("SYS_NNK_NVS_DHU_NMK", myStaff.SICode) = 0 Then Exit Sub
        Me.LblChangeCounter.Enabled = False
        cmd.CommandText = String.Format("update tblUser set Counter='{0}', Location='{1}' where Sicode='{2}'", Me.CmbCounter.Text, Me.CmbLocation.Text, Me.txtUserManagementSIcode.Text)
        cmd.ExecuteNonQuery()
    End Sub

   
   
    Private Sub lbkCcAccess_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkCcAccess.LinkClicked
        frmCcAccess.ShowDialog()
    End Sub
End Class