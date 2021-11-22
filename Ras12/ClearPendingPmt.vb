Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.MySharedFunctionsWzConn

Public Class ClearPendingPmt
    Private BasedCurr As String
    Private TTLVND As Decimal = -1
    Private DaTra As Decimal
    Private varAction As String
    Private VND_Avail As Decimal
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal parAction As String)
        InitializeComponent()
        varAction = parAction
    End Sub
    Private Sub ClearPendingPmt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If HasNewerVersion_R12(Application.ProductVersion) Then
            Me.Close()
            Me.Dispose()
            End
        End If
        If SysDateIsWrong(Conn) Then
            Me.Close()
            Me.Dispose()
            End
        End If

        Me.BackColor = pubVarBackColor
        If varAction.Substring(0, 2) = "FO" Then
            Me.CmbAL.Visible = False
            Me.LblAL.Visible = False
            Me.LblLoad.Visible = False
        End If
        GenComboValue()
        LoadGridPendingRCP()
        If myStaff.Counter = "N-A" Then
            CmdApply.Visible = False
        End If
    End Sub
    Private Sub LoadGridPendingRCP()
        Dim AprBy As String, strSQL As String
        strSQL = "select r.Counter,f.RecID, f.RCPNO, RCPID, FOP, Amount, f.Currency, f.Fstuser as Owner, SRV, Document" _
            & ",r.CustShortName,f.CcId,m.Details as CcNbr,f.rmk, CustomerID" _
            & " from FOP f " _
            & " inner join RCP r on r.recid=f.rcpid" _
            & " left join Misc m on m.recid=f.Ccid" _
            & " where f.RecId>0"

        Select Case myStaff.Counter
            Case "CWT", "TVS", "N-A"
                strSQL = strSQL & " and r.Counter='" & myStaff.Counter & "'"
        End Select

        If Me.OptALL.Checked Then
            strSQL = strSQL & " and ((FOP in ('DEB','DLV') and f.status='OK' ) or (FOP ='CRD' and f.status='QQ')) "
        ElseIf Me.OptDEB.Checked Then
            strSQL = strSQL & " and FOP in ('DEB','DLV') and f.status='OK'  "
        ElseIf Me.OptCRD.Checked Then
            strSQL = strSQL & " and FOP ='CRD' and f.status='QQ' "
        End If


        Me.GridPendingRCP.DataSource = GetDataTable(strSQL)
        Me.CmdApply.Enabled = False
        Me.GridPendingRCP.Enabled = True
        Me.LblViewTRX.Visible = False
        For i As Int16 = 0 To Me.GridPendingRCP.RowCount - 1
            If Me.GridPendingRCP.Item("FOP", i).Value = "DEB" Then
                AprBy = ScalarToString("Actionlog", "top 1 ActionBy", "tableName+doWhat='CREXTOK' and F1='" & Me.GridPendingRCP.Item("RCPNO", i).Value & "'")
                If Not String.IsNullOrEmpty(AprBy) Then Me.GridPendingRCP.Item("Owner", i).Value = AprBy
            End If
        Next
        ResizeGridPendingRCP()
        'GridPendingRCP.Columns("RecId").Visible = False
        GridPendingRCP.Columns("CcId").Visible = False
    End Sub
    Private Sub ResizeGridPendingRCP()
        Me.GridPendingRCP.Columns("Counter").Width = 45
        Me.GridPendingRCP.Columns("Currency").Width = 45
        Me.GridPendingRCP.Columns("Owner").Width = 40
        Me.GridPendingRCP.Columns("SRV").Width = 32
        Me.GridPendingRCP.Columns("RCPNO").Width = 85
        Me.GridPendingRCP.Columns("FOP").Width = 40
        Me.GridPendingRCP.Columns("Currency").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.GridPendingRCP.Columns("FOP").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.GridPendingRCP.Columns("RecID").Visible = False
        Me.GridPendingRCP.Columns("RCPID").Visible = False
        Me.GridPendingRCP.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridPendingRCP.Columns("Amount").DefaultCellStyle.Format = "#,##0.00"
        Me.GridPendingRCP.Columns("Amount").Width = 80
        Me.GridPendingRCP.Columns("RCPNO").ReadOnly = True
        Me.GridPendingRCP.Columns("Amount").ReadOnly = True
        Me.GridPendingRCP.Columns("FOP").ReadOnly = True
        Me.GridPendingRCP.Columns("Currency").ReadOnly = True
        Me.GridPendingRCP.Columns("CustomerID").Width = 65
        Me.GridPendingRCP.Columns("CcNbr").Width = 40
    End Sub
    Private Sub GenComboValue()
        Dim strSQL As String = "select VAL from MISC where CAT='FOP' and details like '%CC%' " & _
            " and val not in ('PPD','PSP','ITP') order by VAL"
        Me.FOP.Items.Clear()
        Me.FOP.DataSource = GetDataTable(strSQL)
        Me.FOP.DisplayMember = "VAL"
        Me.FOP.ValueMember = "VAL"

        strSQL = "select distinct currency from ForEx where Status='OK'"
        Me.RCPCurrency.DataSource = GetDataTable(strSQL)
        Me.RCPCurrency.DisplayMember = "Currency"
        Me.RCPCurrency.ValueMember = "Currency"
        Me.GridFOP.Columns(2).ReadOnly = True

        Me.CmbAL.Items.Clear()
        If MySession.Domain <> "GSA" Then
            Me.CmbAL.Items.Add(MySession.TRXCode)
        Else
            LoadCmbAL(Me.CmbAL)
        End If
    End Sub

    Private Sub GridFOP_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridFOP.CellEndEdit
        If e.ColumnIndex = 3 Then
            Me.GridFOP.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDec(Me.GridFOP.Item(e.ColumnIndex, e.RowIndex).Value), "#,##0.00")
        End If
    End Sub

    Private Sub GridFOP_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles GridFOP.RowsAdded
        Me.GridFOP.Item(0, e.RowIndex).Value = "CSH"
        Me.GridFOP.Item(1, e.RowIndex).Value = "VND"
        Me.GridFOP.Item(2, e.RowIndex).Value = 1
    End Sub
    Private Sub GridFOP_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles GridFOP.CellValidating
        If e.ColumnIndex = 3 Or e.ColumnIndex = 2 Then
            Me.GridFOP.Rows(e.RowIndex).ErrorText = ""
            Dim newInteger As Decimal, ErrText As String = "Invalid Input"
            If GridFOP.Rows(e.RowIndex).IsNewRow Then Return
            If Not Decimal.TryParse(e.FormattedValue.ToString(), newInteger) Then
                e.Cancel = True
                Me.GridFOP.Rows(e.RowIndex).ErrorText = ErrText
                MsgBox(ErrText, MsgBoxStyle.Critical, msgTitle)
            End If
        End If
    End Sub
    Private Sub GridFOP_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridFOP.CellEnter
        Dim R As Integer, varROE As Decimal, varCurr As String, tmpAL As String
        If MySession.Domain = "TVS" Then
            tmpAL = "TS"
        Else
            tmpAL = Me.CmbAL.Text
        End If
        R = e.RowIndex
        For i As Int16 = Me.GridFOP.RowCount - 1 To R + 1 Step -1
            Me.GridFOP.Item(3, i).Value = 0
            Me.GridFOP.Item(4, i).Value = ""
        Next

        If e.ColumnIndex = 2 Then
            varCurr = Me.GridFOP.Item(1, R).Value
            varROE = IIf(varCurr = "VND", 1, ForEX_12(Me.TxtDate.Value, varCurr, IIf(BasedCurr = "VND", "BBR", "BSR"), tmpAL))
            Me.GridFOP.Item(2, R).Value = varROE
        ElseIf e.ColumnIndex = 3 Then
            If R = 0 Then
                If CDec(Me.GridFOP.Item(3, R).Value) = 0 Then
                    Me.GridFOP.Item(3, R).Value = TTLVND / Me.GridFOP.Item(2, R).Value
                End If
                Exit Sub
            End If
            If R > 0 Then
                DaTra = 0
                For i As Int16 = 0 To R - 1
                    DaTra = DaTra + CDec(Me.GridFOP.Item(3, i).Value) * Me.GridFOP.Item(2, i).Value
                Next
            End If
            Me.GridFOP.Item(3, R).Value = (TTLVND - DaTra) / Me.GridFOP.Item(2, R).Value
        ElseIf e.ColumnIndex = 4 Then
            DaTra = 0
            For i As Int16 = 0 To R
                DaTra = DaTra + CDec(Me.GridFOP.Item(3, i).Value) * Me.GridFOP.Item(2, i).Value
            Next
        End If
        Me.CmdApply.Enabled = IIf(DaTra = TTLVND, True, False)
        Me.GridFOP.Columns(2).DefaultCellStyle.Format = "#,##0.00"
        Me.GridFOP.Columns(3).DefaultCellStyle.Format = "#,##0.00"
    End Sub

    Private Sub GridPendingRCP_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridPendingRCP.CellClick
        Dim tmpAL As String = Me.GridPendingRCP.CurrentRow.Cells("RCPNo").Value.ToString.Substring(0, 2)
        BasedCurr = Me.GridPendingRCP.CurrentRow.Cells("Currency").Value
        Me.GridPendingRCP.Enabled = False
        TTLVND = Me.GridPendingRCP.CurrentRow.Cells("Amount").Value
        If BasedCurr <> "VND" Then
            TTLVND = TTLVND * ForEX_12(Me.TxtDate.Value, BasedCurr, "BSR", tmpAL).Amount
        End If
        Me.GridFOP.Rows.Clear()
        Me.CmdApply.Enabled = False
        Me.LblViewTRX.Visible = True
        If Me.GridPendingRCP.CurrentRow.Cells("FOP").Value = "CRD" _
            AndAlso (myStaff.Counter = "CWT" Or myStaff.Counter = "N-A") Then
            lblAddCc.Visible = True
        Else
            lblAddCc.Visible = False
        End If
    End Sub
    Private Sub CmdApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdApply.Click
        If SysDateIsWrong(Conn) Then Exit Sub
        Dim strAction As String = varAction
        Dim strSQL As String, tmpCustID As Integer = Me.GridPendingRCP.CurrentRow.Cells("CustomerID").Value
        Dim tmpDoc As String = Me.GridPendingRCP.CurrentRow.Cells("Document").Value
        Dim tmpRMK As String = Me.GridPendingRCP.CurrentRow.Cells("RMK").Value
        Dim lstQuerries As New List(Of String)

        If Me.GridPendingRCP.CurrentRow.Cells("FOP").Value = "CRD" Then strAction = "BOSwipeCRD|" & tmpDoc & "|" & tmpRMK & "|"
        strSQL = ChangeStatus_ByID("FOP", "XX", Me.GridPendingRCP.CurrentRow.Cells("RecID").Value)
        lstQuerries.Add(strSQL)
        For i As Int16 = 0 To Me.GridFOP.RowCount - 1
            If Me.GridFOP.Item("Amount", i).Value <> 0 Then
                Dim strNewFOP As String = Insert_FOP(Me.GridPendingRCP.CurrentRow.Cells("RCPID").Value, Me.GridPendingRCP.CurrentRow.Cells("RCPNO").Value, _
                    Me.GridFOP.Item("FOP", i).Value, Me.GridFOP.Item("RCPCurrency", i).Value, Me.GridFOP.Item("RCPROE", i).Value, _
                    CDec(Me.GridFOP.Item("Amount", i).Value), Me.GridFOP.Item("Document", i).Value, strAction _
                    & Me.GridPendingRCP.CurrentRow.Cells("RecID").Value.ToString, tmpCustID _
                    , Me.GridPendingRCP.CurrentRow.Cells("CcID").Value)

                'strSQL = strSQL & "; " & strNewFOP
                lstQuerries.Add(strNewFOP)
            End If
        Next
        If Not UpdateListOfQuerries(lstQuerries, Conn) Then
            MsgBox("Unable to Clear Pending Payment")
        End If
        'cmd.CommandText = strSQL
        'cmd.ExecuteNonQuery()
        LoadGridPendingRCP()
        Me.CmdClearAnother.PerformClick()
        For i As Int16 = 0 To Me.GridFOP.RowCount - 1
            Me.GridFOP.Item("Amount", i).Value = 0
        Next
    End Sub
    Private Sub CmdClearAnother_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClearAnother.Click
        Me.GridPendingRCP.Enabled = True
    End Sub
    Private Sub CmbRCVDby_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbRCVDby.SelectedIndexChanged
        If Me.CmbRCVDby.Text.Substring(0, 2) = "FO" Then
            Me.CmbAL.Enabled = False
        Else
            Me.CmbAL.Enabled = True
        End If
    End Sub

    Private Sub LblLoad_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblLoad.LinkClicked
        LoadGridPendingRCP()
    End Sub

    Private Sub OptALL_CheckedChanged(sender As Object, e As EventArgs) Handles OptALL.Click, OptCRD.Click, OptDEB.Click
        LoadGridPendingRCP()
    End Sub

    Private Sub LblViewTRX_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblViewTRX.LinkClicked
        Dim vAL As String = Me.GridPendingRCP.CurrentRow.Cells("RCPNO").Value.ToString.Substring(0, 2)
        Dim vDomain As String = IIf(vAL = "TS", "TVS", "GSA")
        Try
            InHoaDon(Application.StartupPath, "R12_TRX_CFM.xlt", "V", Me.GridPendingRCP.CurrentRow.Cells("RCPNO").Value, Now.Date, Now.Date, 0, vAL, vDomain, "")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub lblAddCc_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblAddCc.LinkClicked

        With GridPendingRCP.CurrentRow
            Dim intRecId As Integer = ScalarToInt("Misc", "RecId", " Cat='CcAccess' and Status='OK' and Val='" _
                                                  & .Cells("CustShortName").Value & "' and Val1='" & myStaff.SICode & "'")
            If intRecId = 0 Then
                MsgBox("You are not In charge of this Customer")
                Exit Sub
            End If
            Dim frmAddCc As New frmSelectCcNbr(GridPendingRCP.CurrentRow)
            If frmAddCc.ShowDialog = Windows.Forms.DialogResult.OK Then
                If ExecuteNonQuerry("Update FOP set CcId =" & frmAddCc.CcId & " where RecId=" _
                                 & GridPendingRCP.CurrentRow.Cells("RecId").Value, Conn) Then
                    LoadGridPendingRCP()
                End If
            End If
        End With

    End Sub
End Class
