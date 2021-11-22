Imports SharedFunctions.Crd_Ctrl
Public Class BalancePreview
    Private Whois As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private MyCust As New objCustomer
    Public Sub New(ByVal parWhatAction As String)
        InitializeComponent()
        Whois = parWhatAction
    End Sub
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub OptPSP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        OptPSP.Click, OptPPD.Click, OptCWT.Click, OptFT.Click, Opt1S.Click, OptALL.Click, OptLC.Click
        LoadGridBalance()
    End Sub

    Private Sub BalancePreview_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MyCust.CustID = 0
    End Sub
    Private Sub BalancePreview_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        LoadGridBalance()
        Me.PnlDetail.Top = 2
        Me.PnlDetail.Left = 2
        Me.PnlDetail.Width = 828
        Me.PnlDetail.Height = 497
    End Sub
    Private Sub LoadGridBalance()
        Me.PnlDetail.Visible = False
        Dim strSQL As String

        If Me.OptPPD.Checked Then
            strSQL = "Select c.CustID, c.CustShortName, VND_PPD_Avail as VND_Avail"
            strSQL = strSQL & " from CC_BLC c where PPCoef>0 and RecID in (Select Max(RecID) From CC_BLC  group by CustID) "
        Else
            strSQL = "Select c.CustID, c.CustShortName, VND_PSP_Avail as VND_Avail"
            strSQL = strSQL & ", OverDue, getdate() as CtrExpDate from CC_BLC c inner join Kybaocao k on k.custid=c.custid and k.status<>'XX' "
            strSQL = strSQL & " inner join CC_Setting s on s.custid=c.custid and s.status<>'XX' and s.crCoef>0  "
            strSQL = strSQL & " where c.RecID in (Select Max(RecID) From CC_BLC group by CustID) "
        End If
        If Me.OptFT.Checked Then
            strSQL = strSQL & " and c.CustID in (select CustID from officeID where status<>'XX')"
        ElseIf Me.OptCWT.Checked Then
            strSQL = strSQL & " and c.CustID in (select CustID from cust_detail where cat+VAl='ChannelCS' and status<>'XX')"
        ElseIf Me.OptLC.Checked Then
            strSQL = strSQL & " and c.CustID in (select CustID from cust_detail where cat+VAl='ChannelLC' and status<>'XX')"
        ElseIf Me.Opt1S.Checked Then
            strSQL = strSQL & " and c.CustID in (select CustID from Fox_iAmChick where status<>'XX')"
        End If
        strSQL = strSQL & " order by c.CustShortName "
        Me.GridBalance.DataSource = GetDataTable(strSQL)
        Me.GridBalance.Columns("CustID").Width = 48

        Me.GridBalance.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.GridBalance.Columns(1).DefaultCellStyle.Format = "#,##0"
        Me.GridBalance.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridBalance.Columns(2).DefaultCellStyle.Format = "#,##0"
        Me.GridBalance.Columns(1).Width = 80
        Me.GridBalance.Columns(2).Width = 90
        If Me.OptPSP.Checked Then Me.GridBalance.Columns(3).Width = 56
        Me.TxtEmail.Text = ""
        Me.LstOffice.Items.Clear()
        Me.LstTKT.Items.Clear()
        Me.LstTKT.Visible = Me.GridBalance.Columns(2).Visible
        Me.Label2.Visible = Me.GridBalance.Columns(2).Visible
        Me.LblRefresh.Enabled = False
        Me.LblChickBalance.Visible = False
        MyCust.CustID = 0

        If Me.OptPSP.Checked Then
            Me.GridBalance.Columns("CtrExpDate").Width = 75
            Dim tmpDate As String
            For i As Int16 = 0 To Me.GridBalance.RowCount - 1
                cmd.CommandText = String.Format("select top 1 ContractValidity from BG where status='OK' and bank not in ('CSH') and custid={0} order by ContractValidity desc ", _
                    Me.GridBalance.Item("Custid", i).Value)
                tmpDate = cmd.ExecuteScalar
                If tmpDate = "" OrElse tmpDate = "01-01-2000" Then
                    Me.GridBalance.Item("CtrExpDate", i).Value = "01-Jan-2000"
                    Me.GridBalance.Item("CtrExpDate", i).Style.ForeColor = Color.White
                Else
                    Me.GridBalance.Item("CtrExpDate", i).Value = tmpDate
                End If
            Next
        End If

    End Sub

    Private Sub GridBalance_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridBalance.CellClick
        If e.RowIndex < 0 Then Exit Sub
        Dim ChickBLC As Decimal
        MyCust.CustID = Me.GridBalance.CurrentRow.Cells("CustID").Value
        Me.TxtEmail.Text = ""
        Me.LstOffice.Items.Clear()
        Me.LstTKT.Items.Clear()
        For i As Int16 = 0 To UBound(MyCust.Email.Split(";"))
            Me.TxtEmail.Text = Me.TxtEmail.Text & MyCust.Email.Split(";")(i).Trim & vbCrLf
        Next
        For i As Int16 = 0 To UBound(MyCust.OID.Split("_"))
            Me.LstOffice.Items.Add(MyCust.OID.Split("_")(i))
        Next

        If Me.Opt1S.Checked Then
            cmd.CommandText = "select isnull(sum(qty*creditAmt) ,0) from funcTKT_1AnotUpdated2RAS_wzparam('" & _
                Format(Now.Date.AddDays(-16), "dd-MMM-yy") & "')" & _
                " where custid=" & Me.GridBalance.CurrentRow.Cells("CustID").Value & " and left(tkno,3)='738' and PRG='M1S'"
            ChickBLC = cmd.ExecuteScalar
            ChickBLC = Me.GridBalance.CurrentRow.Cells("VND_Avail").Value - ChickBLC
            Me.LblChickBalance.Text = "Balance For StopSales On 1S: VND" & Format(ChickBLC, "#,##0")
        End If
        Me.LblChickBalance.Visible = Me.Opt1S.Checked
        Me.LblRefresh.Enabled = True
    End Sub

    Private Sub LblSearch_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSearch.LinkClicked
        MyCust.CustID = ScalarToInt("OfficeID", "top 1 CustID", String.Format("status='OK' and OfficeID='{0}'", Me.TxtOfficeID.Text))
        If MyCust.CustID <> 0 Then
            MsgBox("CustName: " & MyCust.ShortName & " (" & MyCust.CustID & ")" & vbCrLf & "Email: " & MyCust.Email, MsgBoxStyle.Information, msgTitle)
        Else
            MsgBox("Not Found", MsgBoxStyle.Information, msgTitle)
        End If
        MyCust.CustID = 0
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblRefresh.LinkClicked
        MyCust.CustID = Me.GridBalance.CurrentRow.Cells("CustID").Value
        Me.GridBalance.CurrentRow.Cells("VND_Avail").Value = RefreshBalance(MyCust.DelayType, MyCust.LstReconcile, MyCust.CustID, True, "BlcPrev Refresh", Conn, myStaff.SICode, CnStr)
    End Sub
    Private Sub LoadGridDetail()
        Dim strSQL As String = "select VND_" & IIf(Me.OptPPD.Checked, "PPD", "PSP") & "_avail, FT_Sale, SabreSale, FstUser, " & _
            "FstUpdate as ActionTime "
        If Me.OptRAS.Checked Then strSQL = strSQL & ", RMK"
        strSQL = strSQL & " from cc_BLC where custID=" & Me.GridBalance.CurrentRow.Cells("CustID").Value & _
            " and FstUpdate between '" & Format(Me.txtFrm.Value, "dd-MMM-yy") & "' and '" & _
            Format(Me.TxtTo.Value, "dd-MMM-yy") & "' order by Fstupdate Desc, RecID desc"
        If Me.OptRAS.Checked Then
            Me.GridDetail.DataSource = GetDataTable(strSQL)
        Else
            Me.GridDetail.DataSource = GetDataTable(strSQL, Conn_Web)
        End If
        Me.PnlDetail.Visible = True
        For c As Int16 = 0 To 2
            Me.GridDetail.Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
            Me.GridDetail.Columns(c).DefaultCellStyle.Format = "#,##0.00"
        Next
        Me.GridDetail.Columns(3).Width = 56
        For c As Int16 = 4 To Me.GridDetail.Columns.Count - 1
            Me.GridDetail.Columns(c).Width = 128
        Next
    End Sub
    Private Sub GridBalance_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridBalance.CellContentDoubleClick
        loadGridDetail()
    End Sub

    Private Sub LblCloseDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblCloseDetail.LinkClicked
        Me.PnlDetail.Visible = False
    End Sub

    Private Sub OptRAS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptRAS.Click, OptM1S.Click, txtFrm.ValueChanged, TxtTo.ValueChanged
        LoadGridDetail()
    End Sub

End Class