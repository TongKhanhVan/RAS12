Imports SharedFunctions.MySharedFunctions
Imports SharedFunctions.MySharedFunctionsWzConn
Public Class UpdateForEx
    Private CharEntered As Boolean = False
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Private CGOConn As New SqlClient.SqlConnection
    Private Sub UpdateForEx_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    Private Sub UpdateForEx_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CGOConn.ConnectionString = CnStr_CGO
        CGOConn.Open()
        Me.BackColor = pubVarBackColor
        GenComboValue()
        LoadROE("Status='OK' and EffectDate >'" & Now.Date.AddDays(-10) & "'")
    End Sub
    Private Sub GenComboValue()
        LoadCmb_MSC(Me.CmbCurr, "select VAL from MISC where CAT='CURR' and VAL <>'VND'")
        Me.CmbCurr.Text = "USD"
        Dim tbl As DataTable = GetDataTable("Select AL from airline where status='OK' and istva<>0 ")
        Me.txtApplyROEto.Text = ""
        Me.TxtApplyROE_COS.Text = ""
        For i As Int16 = 0 To tbl.Rows.Count - 1
            Me.txtApplyROEto.Text = Me.txtApplyROEto.Text & " " & tbl.Rows(i)("AL")
        Next
        tbl = GetDataTable("Select AL from airline where status='OK' AND CITY IN ('ALL','" & myStaff.City & "')", CGOConn)
        For i As Int16 = 0 To tbl.Rows.Count - 1
            Me.TxtApplyROE_COS.Text = Me.TxtApplyROE_COS.Text & " " & tbl.Rows(i)("AL")
        Next
        Me.txtApplyROEto.Text = Me.txtApplyROEto.Text.Trim
        Me.TxtApplyROE_COS.Text = Me.TxtApplyROE_COS.Text.Trim
    End Sub
    Private Sub LoadROE(ByVal varDK As String)
        Dim Col As Int16, strSQL As String
        strSQL = "select * from ForEX where not (currency ='VND' and bsr=1) "
        If varDK <> "" Then
            strSQL = strSQL & " and " & varDK
        End If
        strSQL = strSQL & " order by recID desc"
        Me.GridROE.DataSource = GetDataTable(strSQL)
        For Col = 0 To Me.GridROE.ColumnCount - 1
            If Me.GridROE.Columns(Col).Name = "RecID" Or _
                Me.GridROE.Columns(Col).Name = "Currency" Or _
                Me.GridROE.Columns(Col).Name = "FstUser" Or _
                Me.GridROE.Columns(Col).Name = "LstUser" Or _
                Me.GridROE.Columns(Col).Name = "IsActive" Then
                Me.GridROE.Columns(Col).Width = 55
            ElseIf Me.GridROE.Columns(Col).Name = "EffectDate" Or _
                Me.GridROE.Columns(Col).Name = "FstUpdate" Or _
                Me.GridROE.Columns(Col).Name = "LstUpdate" Then
                Me.GridROE.Columns(Col).Width = 110
            ElseIf Me.GridROE.Columns(Col).Name = "ApplyROEto" Then
                Me.GridROE.Columns(Col).Width = 250
                Me.GridROE.Columns(Col).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            ElseIf Me.GridROE.Columns(Col).Name = "BRS" Or _
                Me.GridROE.Columns(Col).Name = "BBR" Or _
                Me.GridROE.Columns(Col).Width = 75 Then
                Me.GridROE.Columns(Col).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
        Next
    End Sub

    Private Sub txtBSR_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles _
        txtBSR.KeyDown, txtBBR.KeyDown, txtSurcharge.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub

    Private Sub txtBSR_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles _
        txtBSR.KeyPress, txtBBR.KeyPress, txtSurcharge.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtBSR_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
        txtBSR.LostFocus, txtBBR.LostFocus, txtSurcharge.LostFocus
        Dim txt As TextBox = CType(sender, TextBox)
        Dim aa As Decimal = CDec(txt.Text)
        txt.Text = Format(aa, "#,##0.00")
    End Sub

    Private Sub CmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdFilter.Click
        Dim StrDK As String, i As Integer, CellValStr As String, j As Integer, colName As String
        Try
            j = Me.GridROE.CurrentCell.ColumnIndex()
            colName = Me.GridROE.Columns(j).Name

            If Me.txtFilterValue.Text.Trim = "" Then
                i = Me.GridROE.CurrentCell.RowIndex()
                CellValStr = Me.GridROE.Rows(i).Cells.Item(j).Value
                StrDK = colName & "='" & CellValStr & "'"
            Else
                StrDK = colName & " like '%" & Me.txtFilterValue.Text.Trim & "%'"
            End If
            LoadROE(StrDK)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CmdMark_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdMark.Click
        Dim strSQL As String = ChangeStatus_ByID("ForEx", "XX", GridROE.CurrentRow.Cells("RecID").Value)
        cmd = New SqlClient.SqlCommand(strSQL, Conn)
        Cmd.ExecuteNonQuery()
        LoadROE("")
    End Sub

    Private Sub CmdShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdShow.Click
        LoadROE("")
    End Sub

    Private Sub CmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdAdd.Click
        If SysDateIsWrong(Conn) Then Exit Sub
        If MySession.City = "SGN" AndAlso Me.TxtApplyROE_COS.Text.Trim = "" Then
            MsgBox("Invalid Apply2 (COS)!")
            Exit Sub
        End If
        cmd.CommandText = "Insert into ForEx (Currency, BSR, BBR, EffectDate, SurCharge, ApplyROEto, ApplySCto, FstUser)" & _
            " values ('" & Me.CmbCurr.Text & "','" & CDec(Me.txtBSR.Text) & "','" & CDec(Me.txtBBR.Text) & "','" & _
            Me.txtEffectDate.Value & "','" & CDec(Me.txtSurcharge.Text) & "','" & Me.txtApplyROEto.Text & "','" & _
            Me.txtApplySCto.Text & "','" & myStaff.SICode & "')"
        cmd.ExecuteNonQuery()
        LoadROE("")

        If MySession.City = "SGN" Then
            Dim cmd_CGO As SqlClient.SqlCommand = CGOConn.CreateCommand
            Try
                If CGOConn.State <> ConnectionState.Open Then CGOConn.Open()
                cmd_CGO.CommandText = "Insert into ForEx (Currency, BSR, BBR, EffectDate, SurCharge, ApplyROEto, ApplySCto, City, FstUser)" & _
                    " values ('" & Me.CmbCurr.Text & "','" & CDec(Me.txtBSR.Text) & "','" & CDec(Me.txtBBR.Text) & "','" & _
                    Me.txtEffectDate.Value & "','" & CDec(Me.txtSurcharge.Text) & "','" & Me.TxtApplyROE_COS.Text & "','" & _
                    Me.txtApplySCto.Text & "','" & myStaff.City & "','" & myStaff.SICode & "')"
                cmd_CGO.ExecuteNonQuery()
                CGOConn.Close()
            Catch ex As Exception
                MsgBox("Cant Update ForEx For Cargo. Retry Later", MsgBoxStyle.Critical, msgTitle)
            End Try
        End If
    End Sub
End Class