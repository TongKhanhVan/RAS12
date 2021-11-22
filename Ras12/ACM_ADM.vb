Imports SharedFunctions.MySharedFunctionsWzConn
Public Class ACM_ADM
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private MyCust As New objCustomer
    Private Sub TxtNote_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNote.GotFocus
        If Me.TxtNote.Text = "Note" Then
            Me.TxtNote.Text = ""
            Me.TxtNote.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TxtNote_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNote.LostFocus
        Me.LblUpdate.Visible = False
        If Me.TxtNote.Text = "Note" Or Me.TxtNote.Text = "" Then
            Me.TxtNote.Text = "Note"
            Me.TxtNote.ForeColor = Color.DarkGray
        Else
            If (Me.OptACM.Checked Or Me.OptADM.Checked) And Me.TxtAmt.Text <> 0 Then
                Me.LblUpdate.Visible = True
            End If
        End If
    End Sub

    Private Sub ACM_ADM_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MyCust.CustID = 0
        Me.Dispose()
    End Sub
    Private Sub ACM_ADM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MyCust.GenCustList()
        Dim StrSQL As String = MyCust.List_CC
        If MySession.Domain = "EDU" Then
            strSQL = strSQL & " and custid <0"
        End If
        LoadCmb_VAL(Me.CmbCustomer, StrSQL & " order by CustShortName")
        Me.BackColor = pubVarBackColor
    End Sub
    Private Sub LoadGridACM()
        Me.GridACMADM.DataSource = Nothing
        Dim StrSQL As String = String.Format("Select RecID, CustID, CustShortName, OrgDate, OrgAmt as Amount, OrgCurr as Curr, FOP, " & _
            "PmtType as type, Note, FstUser from khachtra where status='OK' and CustID={0} and fop in ('ACM','ADM')", Me.CmbCustomer.SelectedValue)
        Me.GridACMADM.DataSource = GetDataTable(StrSQL)
        Me.GridACMADM.Columns("RecID").Visible = False
        Me.GridACMADM.Columns("CustID").Width = 56
        Me.GridACMADM.Columns("CustShortName").Width = 128
        Me.GridACMADM.Columns("OrgDate").Width = 72
        Me.GridACMADM.Columns("Amount").Width = 75
        Me.GridACMADM.Columns("Curr").Width = 32
        Me.GridACMADM.Columns("FOP").Width = 32
        Me.GridACMADM.Columns("Type").Width = 32
        Me.GridACMADM.Columns("FstUser").Width = 32
        Me.GridACMADM.Columns("Note").Width = 256
        Me.GridACMADM.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridACMADM.Columns("Amount").DefaultCellStyle.Format = "#,###"

    End Sub
    Private Sub OptACM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptACM.Click
        LoadGridACM()
    End Sub

    Private Sub OptADM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptADM.Click
        LoadGridADM()
    End Sub

    Private Sub LoadGridADM()
        Me.GridACMADM.DataSource = Nothing
        Dim StrSQL As String = String.Format("Select RecID, CustID, CustShortName, InvDate, InvAmt as Amount, InvCurr as Curr, " & _
            "DebType as Type, Note, FstUser from GhiNoKhach where status<>'XX' and CustID={0} and left(note,4)='ADM:'", Me.CmbCustomer.SelectedValue)
        Me.GridACMADM.DataSource = GetDataTable(StrSQL)
        Me.GridACMADM.Columns("RecID").Visible = False
        Me.GridACMADM.Columns("CustID").Width = 56
        Me.GridACMADM.Columns("CustShortName").Width = 128
        Me.GridACMADM.Columns("invDate").Width = 72
        Me.GridACMADM.Columns("Amount").Width = 75
        Me.GridACMADM.Columns("Curr").Width = 32
        Me.GridACMADM.Columns("Type").Width = 32
        Me.GridACMADM.Columns("FstUser").Width = 32
        Me.GridACMADM.Columns("Note").Width = 256
        Me.GridACMADM.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridACMADM.Columns("Amount").DefaultCellStyle.Format = "#,###"
    End Sub
    Private Sub TxtAmt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAmt.LostFocus
        Dim aa As Decimal
        Try
            aa = CDec(Me.TxtAmt.Text)
            Me.TxtAmt.Text = Format(aa, "#,##0")
            If aa > 0 Then Me.LblTRF2SMS.Visible = True
        Catch ex As Exception
            Me.TxtAmt.Focus()
        End Try
    End Sub
    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        Dim OrgDoc As String, Amt As Decimal = CDec(Me.TxtAmt.Text)
        If Amt = 0 Or Me.TxtNote.Text = "Note" Or Me.TxtNote.Text = "" Then
            MsgBox("Invalid Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        OrgDoc = GenOrgDoc()
        Me.LblUpdate.Visible = False
        If Me.OptACM.Checked Or (Me.OptADM.Checked And MyCust.DelayType = "PPD") Then
            If Me.OptADM.Checked And MyCust.DelayType = "PPD" Then
                Amt = Amt * -1
            End If
            Insert_KhachTra("E", Me.TxtDate.Value.Date, "VND", IIf(Amt > 0, "ACM", "ADM"), OrgDoc, Amt, Me.TxtNote.Text, "AUTO", 1, OrgDoc, MyCust.DelayType, MyCust.CustID)
        ElseIf Me.OptADM.Checked And MyCust.DelayType = "PSP" Then
            Insert_GhiNoKhach("E", "VND", Now.Date, CDec(Me.TxtAmt.Text), "IV", "ADM:" & Me.TxtNote.Text, Me.TxtDate.Value.Date, 1, MyCust.CustID)
        End If
        If Me.OptACM.Checked Then
            LoadGridACM()
        Else
            LoadGridADM()
        End If
    End Sub
    Private Function GenOrgDoc() As String
        Dim HHMMs As String, KQ As String
        HHMMs = GetHHMMfrmSQL(Conn)
        KQ = "B" & Format(Now, "ddMMyy") & HHMMs
        Return KQ
    End Function

    Private Sub CmbCustomer_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCustomer.Validated
        MyCust.CustID = Me.CmbCustomer.SelectedValue
    End Sub
End Class