Imports SharedFunctions.MySharedFunctions
Public Class View
    Private Whois As String
    Private tmpBiz As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal parWhois As String)
        InitializeComponent()
        Whois = parWhois
    End Sub
    Private Sub CmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSearch.Click
        Me.GridFOP.Visible = False
        If Me.CmbAL.Text = "" Then
            MsgBox("Plz Enter Airline Code", MsgBoxStyle.Information, msgTitle)
            Exit Sub
        End If
        If Me.OptByDate.Checked Then
            SearchByDate()
        ElseIf Me.OptBySearch.Checked = True Then
            SearchSpecific()
        End If
        ReSizeColR(0)
    End Sub
    Private Sub SearchByDate()
        Dim strSQL As String
        strSQL = "select RecID, RCPNO as TRX, SRV, CustType as Channel, CustID, CustShortName as CustName,"
        strSQL = strSQL & " Status, TTLDue, Currency as Curr, ROE, Charge, Discount, DOS, Counter, Location,"
        strSQL = strSQL & " RPTNO, PrintedCustName, PrintedCustAddrr, PrintedTaxCode, FstUser, FstUpdate from RCP "
        strSQL = strSQL & " where status ='QQ' and FstUpdate between '" & Me.txtFrmDate.Text
        strSQL = strSQL & "' and ' " & Me.txtToDate.Text & " 23:59'"
        If tmpBiz = "GSA" Then
            strSQL = strSQL & " and al='" & Me.CmbAL.Text & "' "
        End If
        Try
            If Me.cmbValueToSearch.Text <> "0" And Me.cmbValueToSearch.Text <> "CustID" Then strSQL = strSQL & " and Custid=" & CInt(Me.cmbValueToSearch.Text)
        Catch ex As Exception
        End Try
        Me.GridRCP.DataSource = GetDataTable(strSQL)
        Me.GridRCP.Columns(0).Visible = False
    End Sub
    Private Sub SearchSpecific()
        Dim tmpRCPID As Integer = 0, DK As String, strSQL As String
        If Me.CmbAL.Text = "" Then Exit Sub
        If Me.CmbFindWhat.Text.Substring(0, 3) = "TKT" Then
            DK = " RecID in (select RCPID from TKT "
            DK = DK & " where TKNO='" & Me.cmbValueToSearch.Text & "' and tkt.Statusal<>'xx')"
        ElseIf Me.CmbFindWhat.Text.Substring(0, 3) = "TRX" Then
            DK = " RCPNO='" & Me.cmbValueToSearch.Text & "'"
        ElseIf Me.CmbFindWhat.Text.Substring(0, 3) = "RPT" Then
            DK = " RPTNO='" & Me.cmbValueToSearch.Text & "'"
        ElseIf Me.CmbFindWhat.Text = "TCODE" Then
            DK = " DeliveryStatus='" & Me.cmbValueToSearch.Text & "'"
        Else
            DK = " RecID in (select RCPID from TKT "
            DK = DK & " where TKNO like '%TV%' and dependent='" & Me.cmbValueToSearch.Text & "')"
        End If
        strSQL = "select RecID, RCPNO as TRX, SRV, CustType as Channel, CustID, CustShortName as CustName,"
        strSQL = strSQL & " Status, TTLDue, Currency as Curr, ROE, Charge, Discount, DOS, Counter,  Location,"
        strSQL = strSQL & " RPTNO, PrintedCustName, PrintedCustAddrr, PrintedTaxCode, FstUser, FstUpdate from RCP "
        strSQL = strSQL & " where " & DK & " and custType in " & myStaff.CAccess
        Me.GridRCP.DataSource = GetDataTable(strSQL)
        If Me.GridRCP.Rows.Count = 0 Then
            MsgBox("Transaction not found!", MsgBoxStyle.Critical, msgTitle)
        End If
    End Sub
    Private Sub ReSizeColR(ByVal pFOP As Int16)
        Dim strColName As String, ColSize As Integer
        For i As Int16 = 0 To Me.GridRCP.ColumnCount - 1
            strColName = Me.GridRCP.Columns(i).Name
            If InStr(strColName, "ID") > 0 Or InStr(strColName, "User") > 0 Or _
                strColName = "Type" Or strColName = "Channel" Or strColName = "Curr" Or _
                strColName = "SRV" Or InStr(strColName, "Status") > 0 Or strColName = "Copy" Then
                ColSize = 36
            ElseIf InStr(strColName, "Discount") > 0 Or InStr(strColName, "Charge") > 0 _
                Or strColName = "ROE" Then
                ColSize = 70
            ElseIf strColName = "TRX" Then
                ColSize = 85
            Else
                ColSize = 100
            End If
            Me.GridRCP.Columns(i).Width = ColSize
            If InStr("Charge-Discount-TTLDue_ROE", strColName) > 0 Then
                Me.GridRCP.Columns(i).DefaultCellStyle.Format = "#,##0.00"
                Me.GridRCP.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Else
                Me.GridRCP.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End If
        Next
        If pFOP = 0 Then Exit Sub
        Me.GridFOP.Columns("RCPID").Visible = False
        Me.GridFOP.Columns("RCPNO").Visible = False
        For i As Int16 = 0 To Me.GridFOP.ColumnCount - 1
            strColName = Me.GridFOP.Columns(i).Name
            If InStr(strColName, "ID") > 0 Or InStr(strColName, "User") > 0 Or InStr(strColName, "Status") > 0 Or _
                strColName = "FOP" Or strColName = "Currency" Then
                ColSize = 36
            ElseIf strColName = "ROE" Then
                ColSize = 56
            Else
                ColSize = 100
            End If
            Me.GridFOP.Columns(i).Width = ColSize
            If strColName = "ROE" Or strColName = "Amount" Then
                If strColName = "Amount" Then
                    Me.GridFOP.Columns(i).DefaultCellStyle.Format = "#,##0.00"
                Else
                    Me.GridFOP.Columns(i).DefaultCellStyle.Format = "#,##0"
                End If
                Me.GridFOP.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Else
                Me.GridFOP.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End If
        Next

    End Sub
    Private Sub View_Edit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        LoadCmb_MSC(Me.CmbAL, myStaff.TVA)
    End Sub

    Private Sub CmbAL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbAL.LostFocus
        tmpBiz = GetDomainNameFromTRXCode(Me.CmbAL.Text)
    End Sub

    Private Sub OptByDate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptByDate.Click
        Me.CmbFindWhat.Visible = False
    End Sub

    Private Sub OptBySearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptBySearch.Click
        Me.CmbFindWhat.Visible = True
    End Sub

    Private Sub GridRCP_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridRCP.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If Me.GridRCP.CurrentRow.Cells("Status").Value = "QQ" Then Exit Sub
        Dim SQLstr As String = ""
        SQLstr = " select * From FOP where RCPID=" & Me.GridRCP.CurrentRow.Cells("RecID").Value
        Me.GridFOP.DataSource = GetDataTable(SQLstr)
        Me.TxtRPTNo.Text = Me.GridRCP.CurrentRow.Cells("RPTNO").Value
        ReSizeColR(1)
        Me.GridFOP.Visible = True
    End Sub
    Private Sub GridRCP_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridRCP.CellContentDoubleClick
        InHoaDon(Application.StartupPath, "R12_TKTdetailByRCP.xlt", "V", "", Now.Date, Now.Date, Me.GridRCP.CurrentRow.Cells("RecID").Value, "", tmpBiz)
    End Sub
    Private Sub TxtRPTNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRPTNo.TextChanged
        If Me.TxtRPTNo.Text <> "" Then
            Me.LblReprint.Enabled = True
        Else
            Me.LblReprint.Enabled = False
        End If
    End Sub
    Private Sub LblReprint_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblReprint.LinkClicked
        Dim fName As String, varFrm As Date, varTo As Date
        fName = "AR_DailySalesReport.xlt"
        varFrm = DateSerial(2000 + CInt(Me.TxtRPTNo.Text.Substring(7, 2)), CInt(Me.TxtRPTNo.Text.Substring(5, 2)), CInt(Me.TxtRPTNo.Text.Substring(10, 2)))
        varTo = varFrm
        InHoaDon(Application.StartupPath, fName, "P", MySession.TRXCode, varFrm, varTo, 0, MySession.TRXCode, MySession.Domain, 0)
    End Sub

    Private Sub cmbValueToSearch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbValueToSearch.Enter
        If Me.cmbValueToSearch.Text = "CustID" Then
            Me.cmbValueToSearch.Text = 0
            Me.cmbValueToSearch.ForeColor = Color.Black
        End If
    End Sub
    Private Sub cmbValueToSearch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbValueToSearch.LostFocus
        If Me.cmbValueToSearch.Text.Length < 13 Then Exit Sub
        If Me.cmbValueToSearch.Text.Substring(0, 1) = "Z" Then Exit Sub
        If Me.CmbFindWhat.Text.Substring(0, 3) = "TKT" And Me.cmbValueToSearch.Text.Length = 13 Then
            Me.cmbValueToSearch.Text = AddSpace2TKNO(Me.cmbValueToSearch.Text)
        End If
    End Sub
    Private Sub CmbFindWhat_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbFindWhat.LostFocus
        Me.cmbValueToSearch.Items.Clear()
        If Me.CmbFindWhat.Text = "RPT No." Then
            Dim ngayDauThangNay As Date = DateSerial(Now.Year, Now.Month, 1)
            Dim ngayCuoiThangTruoc As Date = ngayDauThangNay.AddDays(-1)
            Dim strDK As String = " select distinct RPTNo from RCP where al='" & Me.CmbAL.Text & "' and status='OK' and rptno <>''"
            Dim strQry As String = strDK & " and dos >'" & Format(ngayCuoiThangTruoc, "dd-MMM-yy") & "' UNION " & strDK & _
                " and dos between '" & Format(ngayDauThangNay.AddMonths(-1), "dd-MMM-yy") & "' and '" & _
                Format(ngayCuoiThangTruoc, "dd-MMM-yy") & "' order by RPTNO desc "
            LoadCmb_MSC(Me.cmbValueToSearch, strQry)
        End If
    End Sub

    Private Sub CmbFindWhat_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbFindWhat.VisibleChanged
        Me.txtFrmDate.Visible = Not Me.CmbFindWhat.Visible
        Me.txtToDate.Visible = Not Me.CmbFindWhat.Visible
        If CmbFindWhat.Visible Then
            Me.cmbValueToSearch.Left = 285
            Me.cmbValueToSearch.Width = 132
            Me.cmbValueToSearch.Text = ""
            Me.cmbValueToSearch.ForeColor = Color.Black
        Else
            Me.cmbValueToSearch.Left = 365
            Me.cmbValueToSearch.Width = 56
            Me.cmbValueToSearch.Text = "CustID"
            Me.cmbValueToSearch.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub LblSearchTC_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSearchTC.LinkClicked
        Dim strDK As String = " RCPID IN (select RCPID from fop where Document='" & Me.TxtTC.Text & "')"
        Try
            Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
            cmd.CommandText = "drop table #MCO"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        cmd.CommandText = "select TKNO into #MCO from TKT where DocType in ('MCO','GRP') and " & strDK
        cmd.ExecuteNonQuery()
        Me.GridTKTinTC.DataSource = GetDataTable("Select RecID, RCPID, RCPNO, TKNO, DOI, PaxName, Itinerary from TKT where status<>'XX' " & _
                                                  " and (" & strDK & " or " & _
                                                  " RCPID in ( Select RCPID from FOP where Document in (select TKNO from #MCO)))")
        Me.GridTKTinTC.Columns(0).Visible = False
        Me.GridTKTinTC.Columns(1).Visible = False
        Me.GridTKTinTC.Columns("DOI").Width = 64
        Me.GridTKTinTC.Columns("Paxname").Width = 128
        Me.GridTKTinTC.Columns("Itinerary").Width = 256

    End Sub
End Class