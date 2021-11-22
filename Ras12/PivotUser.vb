Public Class PivotUser
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private strChannel As String = "TA"

    Private Sub PivotUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        Dim strSQL As String
        strSQL = "select SICode as VAL, SIName as DIS from tblUser where status in ('OK','ON') and "
        strSQL = strSQL & " right(sicode,2) <>'**' and sicode <>'SYS'"
        If myStaff.SICode = "SYS" Then
            strSQL = strSQL & " and (template='S**' or template='') "
        Else
            Me.LblCheckAll.Visible = False
            Me.OptCS.Visible = False
            strSQL = strSQL & " and (template='S***')"
        End If
        LoadCmb_VAL(Me.CmbUser, strSQL)
    End Sub

    Private Sub CmbUser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbUser.SelectedIndexChanged
        LoadLstAgent()
        Me.LblUpdate.Enabled = True
    End Sub
    Private Sub LoadLstAgent()
        Dim dTble As DataTable
        Dim tmpSQL As String, strSQL As String
        Try
            strSQL = "Select CustShortName from CustomerList where status='OK' and RecID in "
            strSQL = strSQL & "(Select CustID from Cust_Detail where status='OK' and VAL='" & Me.CmbUser.SelectedValue & "'"
            strSQL = strSQL & " and cat='PIVOT')"
            dTble = GetDataTable(strSQL & " order by CustShortName")
            Me.LstAgent.Visible = False
            Me.LstAgent.Items.Clear()
            For i As Int16 = 0 To dTble.Rows.Count - 1
                Me.LstAgent.Items.Add(dTble.Rows(i)("CustShortName"), True)
            Next

            tmpSQL = "Select CustShortName from Customerlist where status='OK' and CustShortName not in ("
            tmpSQL = tmpSQL & strSQL & ") and Recid In"
            tmpSQL = tmpSQL & " (select custID from Cust_Detail where Cat='Channel' And VAL='" & strChannel & "'"
            If Me.ChkUnAssign.Checked Then
                tmpSQL = tmpSQL & " and CustID not in (select RecID from CustomerList where CustID in "
                tmpSQL = tmpSQL & "(select CustID from Cust_Detail where status='OK' and cat='PIVOT'))"
            End If
            tmpSQL = tmpSQL & ") order by CustShortName"
            dTble = GetDataTable(tmpSQL)
            For i As Int16 = 0 To dTble.Rows.Count - 1
                Me.LstAgent.Items.Add(dTble.Rows(i)("CustShortName"), False)
            Next
            Me.LstAgent.Visible = True
        Catch ex As Exception

        End Try
    End Sub
    Private Sub setStatus(ByVal pCheck As Boolean)

        For i As Int16 = 0 To Me.LstAgent.Items.Count - 1
            Me.LstAgent.SetItemChecked(i, pCheck)
        Next
    End Sub

    Private Sub LblCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCheckAll.LinkClicked
        setStatus(True)
    End Sub

    Private Sub LblUnCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUnCheckAll.LinkClicked
        setStatus(False)
    End Sub

    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        Dim CustID As Integer, strSQL As String = ""
        Me.LblUpdate.Enabled = False

        Try
            cmd.CommandText = "delete from cust_Detail where status='OK' and cat='PIVOT' and VAL='" & Me.CmbUser.SelectedValue & "'"
            cmd.ExecuteNonQuery()
            For i As Int16 = 0 To Me.LstAgent.Items.Count - 1
                If Me.LstAgent.GetItemChecked(i) Then
                    CustID = ScalarToInt("customerList", "RecID", "status='OK' and CustShortName='" & Me.LstAgent.Items(i).ToString & "'")
                    strSQL = strSQL & "; insert Cust_Detail (CAT, VAL, CustID, fstUser) values ('PIVOT','" & _
                        Me.CmbUser.SelectedValue & "'," & CustID & ",'" & myStaff.SICode & "')"
                End If
            Next
            cmd.CommandText = strSQL.Substring(1)
            cmd.ExecuteNonQuery()
            MsgBox("Updated", MsgBoxStyle.Information, msgTitle)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub OptTA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        OptTA.CheckedChanged, OptCA.CheckedChanged, OptTO.CheckedChanged, OptCS.CheckedChanged
        Dim vOpt As RadioButton = CType(sender, RadioButton)
        On Error Resume Next
        strChannel = vOpt.Name.Substring(3, 2)
        On Error GoTo 0
        LoadLstAgent()
    End Sub

    Private Sub ChkUnAssign_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkUnAssign.Click
        LoadLstAgent()
    End Sub
End Class
