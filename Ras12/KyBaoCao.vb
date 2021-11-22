Public Class KyBaoCao
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private Sub KyBaoCao_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
    Private Sub KyBaoCao_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        Dim strSQL As String = "Select CustID as VAL, CustShortName as DIS from cc_setting where status<>'XX' and crCoef>0 "
        If MySession.Domain = "EDU" Then
            strSQL = strSQL & " and custid <0"
        End If
        LoadCmb_VAL(Me.CmbCustomer, strSQL & " order by CustShortName")
        LoadGridKyBC()
        Me.CmbDOW.Text = Me.CmbDOW.Items(0).ToString
    End Sub
    Private Sub LoadGridKyBC()

        Me.GridKyBC.DataSource = GetDataTable("Select * from KyBaoCao where status='OK'")
        Me.GridKyBC.Columns("RecID").Visible = False
        For c As Int16 = 1 To Me.GridKyBC.ColumnCount - 5
            Me.GridKyBC.Columns(c).Width = 50
        Next
    End Sub

    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        Dim Ky As String, strSQL As String, CC_SettingStatus As String
        If Me.OptDOM.Checked Then
            Ky = Me.TxtTu1.Text.Trim
            For i As Int16 = 2 To 4
                Ky = Ky & "/" & Me.GrpDOM.Controls("TxtTu" & i.ToString.Trim).Text.Trim
            Next
        Else
            Ky = Me.CmbDOW.Text
        End If
        CC_SettingStatus = ScalarToString("CC_Setting", "Status", "CustID=" & Me.CmbCustomer.SelectedValue)
        Try
            strSQL = ChangeStatus_ByDK("kyBaoCao", "XX", " status='OK' and custid=" & Me.CmbCustomer.SelectedValue)
            strSQL = strSQL & "; Insert KyBaoCao (CustId, CustshortName, Periods, BLCMonitor, daystoCfm, Daystopmt, FstUser, DueIn) values ('" & _
                    Me.CmbCustomer.SelectedValue & "','" & Me.CmbCustomer.Text & "','" & Ky & "','" & IIf(Me.ChkMonitorBLC.Checked, "Y", "N") & _
                    "'," & CInt(Me.TxtCFM.Text) & "," & CInt(Me.TxtVAT.Text) & ",'" & myStaff.SICode & "'," & Me.TxtHan.Text & ")"
            If CC_SettingStatus = "QQ" Then
                strSQL = strSQL & "; Update CC_Setting set Status='OK' where crCoef>0 and status='QQ' and custid=" & Me.CmbCustomer.SelectedValue
            End If
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()
            Me.TxtTu1.Text = 1
            LoadGridKyBC()
        Catch ex As Exception
            MsgBox("Error Updating DataBase", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub

    Private Sub TxtDen1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        TxtDen1.LostFocus, TxtDen2.LostFocus, TxtDen3.LostFocus
        Dim txt As TextBox = CType(sender, TextBox)
        Dim i As Int16
        i = CInt(txt.Name.Substring(6))
        If txt.Text > 0 Then
            If Not CInt(txt.Text) > CInt(Me.GrpDOM.Controls("txtTu" & Trim(Str(i))).Text) Or CInt(txt.Text) > 27 Then
                MsgBox("Invalid Input", MsgBoxStyle.Critical, msgTitle)
                txt.Focus()
                Exit Sub
            End If
            Disable_Enable_Rest(i + 1, i + 1, True)
            Me.GrpDOM.Controls("txtTu" & Trim(Str(i + 1))).Text = Me.GrpDOM.Controls("txtDen" & Trim(Str(i))).Text + 1
            Me.GrpDOM.Controls("txtDen" & Trim(Str(i + 1))).Text = 0
        Else
            Disable_Enable_Rest(i + 1, 4, False)
        End If
    End Sub
    Private Sub Disable_Enable_Rest(ByVal frm As Int16, ByVal tto As Int16, ByVal pStatus As Boolean)
        For i As Int16 = frm To tto
            Me.GrpDOM.Controls("txtDen" & Trim(Str(i))).Enabled = pStatus
            If Not pStatus Then
                Me.GrpDOM.Controls("txtTu" & Trim(Str(i))).Text = 0
                Me.GrpDOM.Controls("txtDen" & Trim(Str(i))).Text = 0
            End If
        Next
        Me.TxtDen4.Enabled = False
    End Sub

    Private Sub Label1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label1.DoubleClick
        Me.TxtTu1.Text = 0
        Me.Label1.Enabled = False
    End Sub
    Private Sub OptDOM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptDOM.Click, OptDOW.Click
        Me.GrpDOM.Enabled = Me.OptDOM.Checked
        Me.CmbDOW.Enabled = Me.OptDOW.Checked
        Me.Label17.Visible = Me.OptDOW.Checked
        Me.Label4.Visible = Me.OptDOW.Checked
        Me.TxtHan.Visible = Me.OptDOW.Checked

    End Sub
End Class