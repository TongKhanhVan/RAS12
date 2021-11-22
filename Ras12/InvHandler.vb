Imports SharedFunctions.MySharedFunctions
Public Class InvHandler
    Private varWho As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Sub New()
        InitializeComponent()
    End Sub
    Public Sub New(ByVal pWho As String)
        InitializeComponent()
        varWho = pWho
    End Sub
    Private Sub InvHandler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        If HasNewerVersion_R12(Application.ProductVersion) Then
            Me.Close()
            Me.Dispose()
            End
        End If
        LoadCmb_MSC(Me.CmbAL, myStaff.TVA)
        CheckRightForALLForm(Me)
        If varWho = "FO" Then
            Me.LckLblDelete.Visible = False
            Me.LckLblDelete.Enabled = False
        End If
    End Sub
    Private Sub LblSearchInv_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSearchInv.LinkClicked
        Dim strSQL As String, strDK As String, pSeudoINV As String = Me.TxtInvHandlerInv2Search.Text
        Me.LblPreviewInv.Visible = False
        Me.LblRePrint.Visible = False
        Me.LckLblDelete.Visible = False
        strDK = "left(InvNo,2)='" & Me.CmbAL.Text & "' and "
        Try
            If InStr(pSeudoINV, "/") > 0 Then
                strDK = strDK & "substring(invno,5,2)='" & pSeudoINV.Substring(0, 2)
                strDK = strDK & "' and right(invno,5)='" & Format(CInt(pSeudoINV.Substring(3)), "00000") & "'"
            Else
                strDK = strDK & "substring(invno,5,2)='" & Format(Now, "yy")
                strDK = strDK & "' and right(invno,5)='" & Format(CInt(pSeudoINV), "00000") & "'"
            End If
            strDK = strDK & " and substring(invno,3,1) not in ('0','1') "
            strSQL = "Select RecID, SRV,  CustShortName, Amount, Charge, Discount, PrintCopy, CustFullName, "
            strSQL = strSQL & " CustAddress, CustTaxCode, CustID, INVNO, FstUpdate from INV "
            strSQL = strSQL & " where Status='OK' and city='" & MySession.City & "' and " & strDK
            strSQL = strSQL & " and custid in (select CustID from RCP where CustType in " & myStaff.CAccess & ")"
            'INV HAN TRUNG'



            Me.GridInvHandlerInv.DataSource = GetDataTable(strSQL)
            Me.GridInvHandlerInv.Columns(0).Visible = False
            Me.GridInvHandlerInv.Columns("FstUpdate").Visible = False
            Me.GridInvHandlerInv.Columns("SRV").Width = 30

            Me.GridInvHandlerInv.Columns("CustShortName").Width = 75
            Me.GridInvHandlerInv.Columns("Amount").Width = 100
            Me.GridInvHandlerInv.Columns("Charge").Width = 65
            Me.GridInvHandlerInv.Columns("Discount").Width = 65
            Me.GridInvHandlerInv.Columns("PrintCopy").Width = 50
            Me.GridInvHandlerInv.Columns("CustID").Width = 45
            Me.GridInvHandlerInv.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.GridInvHandlerInv.Columns("Charge").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.GridInvHandlerInv.Columns("Discount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.GridInvHandlerInv.Columns("PrintCopy").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Me.GridInvHandlerInv.Columns("Amount").DefaultCellStyle.Format = "#,##0.00"
            Me.GridInvHandlerInv.Columns("Charge").DefaultCellStyle.Format = "#,##0.00"
            Me.GridInvHandlerInv.Columns("Discount").DefaultCellStyle.Format = "#,##0.00"
        Catch ex As Exception
            MsgBox("Invalid Invoice Number Entered", MsgBoxStyle.Critical, msgTitle)
        End Try
    End Sub

    Private Sub GridInvHandlerInv_CellContentClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridInvHandlerInv.CellContentClick
        Me.LblPreviewInv.Visible = True
        Me.LckLblDelete.Visible = True
    End Sub
    Private Sub LblPreviewInv_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblPreviewInv.LinkClicked
        Dim fName As String, InvNoToPrint As String
        InvNoToPrint = Me.GridInvHandlerInv.CurrentRow.Cells("InvNo").Value
        fName = "R12_VATInvoice.xlt"
        InHoaDon(Application.StartupPath, fName, "V", InvNoToPrint, Now.Date, Now.Date, 0, Me.CmbAL.Text, MySession.Domain)
        Me.LblRePrint.Visible = True
    End Sub
    Private Sub LblDelete_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LckLblDelete.LinkClicked
        Dim StrMsgBox As String = "Wanna Delete This Invoice."
        Dim YearINV As String = Me.GridInvHandlerInv.CurrentRow.Cells("InvNo").Value.ToString.Substring(4, 2)
        Dim INVID As Integer = Me.GridInvHandlerInv.CurrentRow.Cells("recID").Value, RCPID As Integer
        Dim InvNo As String = Me.GridInvHandlerInv.CurrentRow.Cells("InvNo").Value

        If Me.GridInvHandlerInv.CurrentRow.Cells("PrintCopy").Value > 0 Then
            StrMsgBox = StrMsgBox & " Dont Forget To Collect All Copies of It and others related with it."
        End If
        StrMsgBox = StrMsgBox & " Abort This Action?"
        If MsgBox(StrMsgBox, MsgBoxStyle.Critical Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton1, msgTitle) = vbYes Then Exit Sub
        RCPID = ScalarToInt("INV", "RCPID", "RecID=" & INVID)
        If varWho = "FO" Then
            Dim rcpNo As String = ScalarToString("RCP", "RCPNO", "RecID=" & RCPID)
            If (rcpNo.Substring(2, 2) <> Format(Now.Month, "00") And Now.Day > 9) Or _
                YearINV <> Format(Now(), "yy") Then
                MsgBox("You Cant Delete Invoices of Previous Year or After day 10 of Next Month. Contact Accounting Dept for Help", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        End If
        Try
            Dim XX_Status As String = getNextXX_Status(RCPID)
            cmd.CommandText = ChangeStatus_ByDK("INV", XX_Status, "status='OK' and RCPID =" & RCPID) & _
                "; Update TKTNO_INVNO set status='XX', lstuser='" & myStaff.SICode & "', lstupdate=getdate() where status='OK' and rcpid=" & RCPID
            cmd.ExecuteNonQuery()
            MsgBox("Invoice Deleted", MsgBoxStyle.Information, msgTitle)
            Me.LckLblDelete.Visible = False
            Me.LblRePrint.Visible = False
            Me.LblPreviewInv.Visible = False
        Catch ex As Exception
            MsgBox("Error Deleting Invoice", MsgBoxStyle.Information, msgTitle)
        End Try
    End Sub
    Private Sub LblRePrint_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblRePrint.LinkClicked
        Dim myAnswer As Int16, fName As String
        Dim InvID As Integer = Me.GridInvHandlerInv.CurrentRow.Cells("RecID").Value
        Dim InvNo As String = Me.GridInvHandlerInv.CurrentRow.Cells("InvNo").Value
        Dim RCPID As Integer = ScalarToInt("TKTNO_INVNO", "RCPID", "INVID=" & InvID)
        Dim Counter As String = ScalarToString("RCP", "Counter", "REcID=" & RCPID)
        If RCPID = 0 Then ' ko co ban ghi TKTNO_INVNO tuong ung
            TaoBanGhiTKTNO_INVNO_Standard(RCPID, InvNo, InvID, IIf(Counter = "CWT", "WO", "WZ"))
        End If
        myAnswer = MsgBox("Do You Want to Print It? ", MsgBoxStyle.Question Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2, msgTitle)
        If myAnswer = vbNo Then
            Exit Sub
        End If
        If varWho = "FO" Then
            If Me.GridInvHandlerInv.CurrentRow.Cells("PrintCopy").Value > 0 Then
                If myStaff.SupOf = "" Or Format(Me.GridInvHandlerInv.CurrentRow.Cells("FstUpdate").Value, "dd-MMM-yy") <> Format(Now, "dd-MMM-yy") Then
                    MsgBox("Plz Ask Supervisor or Accounting People to Reprint it For You.", MsgBoxStyle.Information, msgTitle)
                    Exit Sub
                End If
            End If
        End If
        fName = "R12_VATInvoice.xlt"
        InHoaDon(Application.StartupPath, fName, "O", InvNo, Now.Date, Now.Date, 0, Me.CmbAL.Text, MySession.Domain)
        cmd.CommandText = UpdateTblINVHistory(InvID, InvNo, 1)
        cmd.ExecuteNonQuery()
    End Sub
End Class