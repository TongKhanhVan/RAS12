Imports SharedFunctions.MySharedFunctions
Public Class F1SUserManager
    Dim f1sConn As New SqlClient.SqlConnection
    Private Sub F1SUserManager_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        f1sConn.Close()
        f1sConn.Dispose()
        Conn_Web.Close()
        Me.Dispose()
    End Sub
    Private Sub LoadGridAgent()
        Dim strSQL As String
        strSQL = "select RecID, Office, MobileNbr, Status, RasShortName, email from f1s_office "
        Me.GridAgent.DataSource = GetDataTable(strSQL, f1sConn)
        Me.GridAgent.Columns("RecID").Visible = False
        Me.GridAgent.Columns(1).Width = 64
        Me.GridAgent.Columns(2).Width = 64
        Me.GridAgent.Columns(3).Width = 32
        Me.GridAgent.Columns(4).Width = 75
        Me.LblDeactive.Visible = False
        Me.LblResetPSW.Visible = False
        Me.LblUpdate.Visible = False
        Me.TxtEmail.Text = ""
        Me.TxtMobi.Text = ""
    End Sub
    Private Sub GridAgent_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridAgent.CellContentClick
        If Me.GridAgent.CurrentRow.Cells("Status").Value = "OK" Then
            Me.LblResetPSW.Visible = True
            Me.LblDeactive.Visible = True
            Me.LblUpdate.Visible = True
            Me.TxtEmail.Text = Me.GridAgent.CurrentRow.Cells("email").Value
            Me.TxtMobi.Text = Me.GridAgent.CurrentRow.Cells("MobileNbr").Value
        End If
    End Sub
    Private Sub LblDeactive_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblDeactive.LinkClicked
        Dim cmd As SqlClient.SqlCommand = f1sConn.CreateCommand
        cmd.CommandText = ChangeStatus_ByID("f1s_Office", "XX", Me.GridAgent.CurrentRow.Cells("RecID").Value)
        cmd.ExecuteNonQuery()
        LoadGridAgent()

    End Sub

    Private Sub LblResetPSW_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblResetPSW.LinkClicked
        Dim cmd1s As SqlClient.SqlCommand = f1sConn.CreateCommand
        Dim cmdWeb As SqlClient.SqlCommand = Conn_Web.CreateCommand
        Dim pSW As String = GenPSW(), OwnerSI As Integer
        cmd1s.CommandText = String.Format("select top 1 RecID from f1s_TVSI where Office='{0}' order by RecID", _
            Me.GridAgent.CurrentRow.Cells("Office").Value)
        OwnerSI = cmd1s.ExecuteScalar

        cmd1s.CommandText = String.Format("update f1s_Office set PSW='{0}' where recid={1}; update f1s_TVSI set PSW='{0}' where recid={2} ", _
            HashToFixedLen(pSW), Me.GridAgent.CurrentRow.Cells("RecID").Value, OwnerSI)
        cmd1s.ExecuteNonQuery()

        cmdWeb.CommandText = "Insert SMSLog (CustID, SMSText, Location, MobileNbr) values (-1,'Your New PSW is " & pSW & "','SGN','" & _
            Me.GridAgent.CurrentRow.Cells("mobileNbr").Value & "')"
        cmdWeb.ExecuteNonQuery()
        LoadGridAgent()
    End Sub
    Private Function GenPSW() As String
        Dim tmpStr As String = "123456789abcdefghjikmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ_@"
        Dim KQ As String, j As Int16
        KQ = ""
        Randomize()
        For i As Int16 = 1 To 8
            j = Rnd() * 58
            KQ = KQ & tmpStr.Substring(j, 1)
        Next
        Return KQ
    End Function
    Private Function ValidContact() As Boolean
        If Me.TxtMobi.Text.Length < 10 Or Me.TxtMobi.Text.Length > 11 Then Return False
        If Me.TxtMobi.Text.Substring(0, 2) <> "09" And Me.TxtMobi.Text.Length = 10 Then Return False
        If Me.TxtMobi.Text.Substring(0, 2) <> "01" And Me.TxtMobi.Text.Length = 11 Then Return False
        If InStr(Me.TxtEmail.Text, "@") = 0 Or InStr(Me.TxtEmail.Text, ",") + InStr(Me.TxtEmail.Text, ";") > 0 Then Return False
        Return True
    End Function
    Private Sub LblCreate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCreate.LinkClicked
        If Not ValidContact() Then Exit Sub
        On Error GoTo ErrHandler
        Dim MinBlc As Decimal = CDec(Me.TxtMinBLC.Text)
        On Error GoTo 0
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim cmd1s As SqlClient.SqlCommand = f1sConn.CreateCommand
        Dim cmdWeb As SqlClient.SqlCommand = Conn_Web.CreateCommand
        Dim pswRaw As String = GenPSW(), pswEncript As String
        Dim tmpRecID As Integer, tmpOffice As String = MySession.City & Me.CmbCustomer.SelectedValue.ToString.Trim
        pswEncript = HashToFixedLen(pswRaw)
        cmd1s.CommandText = "select recID from f1s_office where office='" & tmpOffice & "'"
        tmpRecID = cmd1s.ExecuteScalar
        If tmpRecID > 0 Then Exit Sub
        cmd1s.CommandText = "insert f1s_Office (Office, MobileNbr, email, chucnang, isTVoffc, psw, fstUser, RasShortName) " & _
            "values (@Office, @MobileNbr, @email, 'RES', 0, @psw, @fstUser, @RasShortName)" & _
            "; insert f1s_TVSI (Office, ShortName, Psw, Status) values (@Office, @ShortName, @Psw,'OK')" & _
            "; SELECT SCOPE_IDENTITY() AS [RecID]"
        cmd1s.Parameters.Clear()
        cmd1s.Parameters.Add("@Office", SqlDbType.VarChar).Value = tmpOffice
        cmd1s.Parameters.Add("@MobileNbr", SqlDbType.VarChar).Value = Me.TxtMobi.Text
        cmd1s.Parameters.Add("@email", SqlDbType.VarChar).Value = Me.TxtEmail.Text
        cmd1s.Parameters.Add("@psw", SqlDbType.VarChar).Value = pswEncript
        cmd1s.Parameters.Add("@fstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd1s.Parameters.Add("@RasShortName", SqlDbType.VarChar).Value = Me.CmbCustomer.Text
        cmd1s.Parameters.Add("@ShortName", SqlDbType.VarChar).Value = Me.TxtOwnerName.Text
        tmpRecID = cmd1s.ExecuteScalar
        cmdWeb.CommandText = "Insert SMSLog (CustID, SMSText, Location, MobileNbr) values (-1,'" & "Account Has Been Created. PSW is " & pswRaw & "','SGN','" & _
            Me.TxtMobi.Text & "')"
        cmdWeb.ExecuteNonQuery()
        cmd.CommandText = "update CC_Setting set F1sCoef=1, MinBLC=" & MinBlc & " where CustID=" & Me.CmbCustomer.SelectedValue & " and status='OK'"
        cmd.ExecuteNonQuery()
        LoadGridAgent()
        MsgBox("Master User Has Been Created. PSW is Being Sent To Owner's Mobile." & vbCrLf & " - PCC: " & tmpOffice & vbCrLf & " - Sicode: " & tmpRecID, MsgBoxStyle.Information, msgTitle)
ErrHandler:

    End Sub

    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        If Not ValidContact() Then Exit Sub
        Dim cmd As SqlClient.SqlCommand = f1sConn.CreateCommand
        cmd.CommandText = "update f1s_office set mobileNbr='" & Me.TxtMobi.Text & "', Email='" & Me.TxtEmail.Text & "' where recid=" & Me.GridAgent.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridAgent()
    End Sub

    Private Sub CmbCustomer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbCustomer.SelectedIndexChanged
        Me.TxtMobi.Text = ""
        Me.TxtEmail.Text = ""
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Timer1.Enabled = False
        Try
            f1sConn.ConnectionString = CnStr_F1S
            f1sConn.Open()
            Conn_Web.Open()
            Me.TxtWait.Visible = False
            Me.GroupBox1.Visible = True
            LoadGridAgent()
        Catch ex As Exception
            MsgBox("Unable to Connect To Front 1S Server. Plz Try Later", MsgBoxStyle.Critical, msgTitle)
            Me.Close()
        End Try
    End Sub

    Private Sub F1SUserManager_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        LoadCmb_VAL(Me.CmbCustomer, "select custID as val, CustShortName as DIS from cc_setting where status='OK' and custid in (select custid from cust_detail where cat='Channel' and val='TA' and status='OK')")
    End Sub

End Class