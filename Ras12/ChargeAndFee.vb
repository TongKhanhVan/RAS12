Imports SharedFunctions.MySharedFunctions
Public Class ChargeAndFee
    Private CharEntered As Boolean = False
    Dim WhatAction As String
    Private cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal parWhatAction As String)
        InitializeComponent()
        WhatAction = parWhatAction
    End Sub
    Private Sub ChargeAndFee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
        GenComboValue()
        LoadChargeAndFee("")
        Me.CmbSBU.Text = "GSA"
        Me.CmbChannel.Text = "??"
        If WhatAction = "COMM" Then
            Me.OptTV.Visible = False
            Me.OptAL.Visible = False
            Me.Label5.Visible = False
            Me.CmbChargeOn.Visible = False
            Me.ChkNonRefundable.Visible = False
            Me.Label2.Text = "Product"
            Me.CmbChannel.Enabled = False
            Me.CmbSBU.Enabled = False
        Else
            Me.CmbChargeOn.Text = "TKT"
            If Me.OptAL.Checked Then
                Me.CmbChannel.Enabled = False
                Me.CmbChannel.Text = "??"
            Else
                Me.CmbChannel.Enabled = True
            End If
        End If
        Me.CmbTRX.Text = Me.CmbTRX.Items(0).ToString
    End Sub
    Private Sub GenComboValue()
        LoadCmb_MSC(Me.CmbCurr, "select VAL from MISC where CAT='CURR'")
        Me.CmbCurr.Text = "USD"

        LoadCmb_MSC(Me.CmbAL, myStaff.TVA)
        Me.CmbAL.Items.Add("YY")
        Me.CmbAL.Text = "XX"
    End Sub
    Private Sub LoadChargeAndFee(ByVal varDK As String)
        Dim Col As Int16, strSQL As String = String.Format("select * from Charge_comm where CAT+AL='{0}'", WhatAction & Me.CmbAL.Text)
        If varDK <> "" Then
            strSQL = String.Format("{0}  and {1}", strSQL, varDK)
        End If
        If Me.ChkActive.Checked Then
            strSQL = String.Format("{0} and status='OK'", strSQL)
        End If
        strSQL = String.Format("{0} order by recID desc", strSQL)
        Me.GridROE.Columns.Clear()
        Me.GridROE.DataSource = GetDataTable(strSQL)
        Me.GridROE.Columns(0).Visible = False
        Me.GridROE.Columns(1).Visible = False
        For Col = 1 To Me.GridROE.ColumnCount - 1
            If Me.GridROE.Columns(Col).Name = "CAT" Or _
                Me.GridROE.Columns(Col).Name = "Currency" Or _
                Me.GridROE.Columns(Col).Name = "FstUser" Or _
                Me.GridROE.Columns(Col).Name = "LstUser" Or _
                Me.GridROE.Columns(Col).Name = "Amount" Then
                Me.GridROE.Columns(Col).Width = 56
            ElseIf InStr(Me.GridROE.Columns(Col).Name, "Valid") + _
                InStr(Me.GridROE.Columns(Col).Name, "Date") Then
                Me.GridROE.Columns(Col).Width = 75
            ElseIf InStr("AL_Status", Me.GridROE.Columns(Col).Name) > 0 Then
                Me.GridROE.Columns(Col).Width = 32
            ElseIf InStr("_TRX", Me.GridROE.Columns(Col).Name) > 0 Then
                Me.GridROE.Columns(Col).Visible = False
            End If
        Next
        Me.GridROE.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    Private Sub CmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdFilter.Click
        Dim StrDK As String, i As Integer, CellValStr As String, j As Integer, colName As String
        On Error GoTo errHandler
        j = Me.GridROE.CurrentCell.ColumnIndex()
        colName = Me.GridROE.Columns(j).Name

        If Me.txtFilterValue.Text.Trim = "" Then
            i = Me.GridROE.CurrentCell.RowIndex()
            CellValStr = Me.GridROE.Rows(i).Cells.Item(j).Value
            StrDK = colName & "='" & CellValStr & "'"
        Else
            StrDK = colName & " like '%" & Me.txtFilterValue.Text.Trim & "%'"
        End If
        LoadChargeAndFee(StrDK)
errHandler:
        Exit Sub
    End Sub

    Private Sub CmdMark_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdMark.Click
        Dim r As Int16 = Me.GridROE.CurrentCell.RowIndex
        cmd.CommandText = ChangeStatus_ByID("Charge_Comm", "XX", Me.GridROE.Item(0, r).Value)
        cmd.ExecuteNonQuery()
        LoadChargeAndFee("")
    End Sub

    Private Sub CmdShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdShow.Click
        LoadChargeAndFee("")
    End Sub

    Private Sub CmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdAdd.Click
        Dim strChargeName As String, strAmt As String = Me.txtAmount.Text
        If Me.CmbAmtType.Text = "PCT" And CInt(strAmt) > 9 Then
            MsgBox("Illogic Input", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If
        If WhatAction = "CHGE" Then
            If Me.CmbChargeOn.Text = "SEG" Then
                strAmt = strAmt & "/S"
            End If
            If Me.txtChargeName.Text.Length < 1 Or Me.txtChargeName.Text.Length > 6 Or _
                InStr(Me.txtChargeName.Text, "-") > 0 Or _
                InStr("AL_TV", Me.txtChargeName.Text.Substring(0, 2)) > 0 Then
                MsgBox("Invalid Charge Name.", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
            strChargeName = IIf(Me.OptTV.Checked, "TV", "SP")
            If Me.ChkNonRefundable.Checked Then
                strChargeName = strChargeName & "N-"
            Else
                strChargeName = strChargeName & "R-"
            End If
            strChargeName = strChargeName & Me.txtChargeName.Text
        Else
            strChargeName = Me.txtChargeName.Text
        End If
        cmd.CommandText = "Insert into Charge_Comm (AL, CAT, Currency, Type, amount, ValidFrom, ValidThru, FstUser," & _
                " SBU, Channel, TRX, AmtType) values (@AL, @CAT, @Currency, @Type, @amount, @ValidFrom, @ValidThru, @FstUser," & _
                " @SBU, @Channel, @TRX, @AmtType)"
        cmd.Parameters.Clear()
        cmd.Parameters.Add("@AL", SqlDbType.VarChar).Value = Me.CmbAL.Text
        cmd.Parameters.Add("@CAT", SqlDbType.VarChar).Value = WhatAction
        cmd.Parameters.Add("@Currency", SqlDbType.VarChar).Value = Me.CmbCurr.Text
        cmd.Parameters.Add("@Type", SqlDbType.VarChar).Value = strChargeName
        cmd.Parameters.Add("@amount", SqlDbType.Decimal).Value = strAmt
        cmd.Parameters.Add("@ValidFrom", SqlDbType.DateTime).Value = Me.txtEffectFrom.Value.Date
        cmd.Parameters.Add("@ValidThru", SqlDbType.DateTime).Value = Me.txtEffectThru.Value.Date.AddDays(1)
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@SBU", SqlDbType.VarChar).Value = Me.CmbSBU.Text
        cmd.Parameters.Add("@Channel", SqlDbType.VarChar).Value = Me.CmbChannel.Text
        cmd.Parameters.Add("@TRX", SqlDbType.VarChar).Value = Me.CmbTRX.Text.Substring(0, 1)
        cmd.Parameters.Add("@AmtType", SqlDbType.VarChar).Value = Me.CmbAmtType.Text
        cmd.ExecuteNonQuery()
        LoadChargeAndFee("")
    End Sub

    Private Sub txtAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyDown
        CharEntered = checkCharEntered(e.KeyValue)
    End Sub

    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        If CharEntered Then
            e.Handled = True
        End If
    End Sub
    Private Function GetStrDK_AL() As String
        Dim KQ As String = ""
        For i As Int16 = 0 To Me.LstAL_ext.Items.Count - 1
            If Me.LstAL_ext.GetItemChecked(i) Then
                KQ = KQ & "," & Me.LstAL_ext.Items(i).ToString
            End If
        Next
        If KQ <> "" Then
            KQ = KQ.Substring(1).Replace(",", "','")
            KQ = "('" & KQ & "')"
        End If
        Return KQ
    End Function
    Private Sub LblExtend_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblExtend.LinkClicked
        Dim StrDK_AL As String = GetStrDK_AL()
        If StrDK_AL = "" Then Exit Sub
        If WhatAction = "COMM" Then
            cmd.CommandText = "update Charge_Comm Set ValidThru='" & Format(Me.txtNewValidTo.Value, "dd-MMM-yy") & " 23:59' " & _
                ", LstUser='" & myStaff.SICode & "', LstUpdate=Getdate() where status='OK' " & _
                " and cat='" & WhatAction & "' and ValidThru='" & Format(Me.txtCurrentValidTo.Value, "dd-MMM-yy") & " 23:59' " & _
                " and AL in " & StrDK_AL
        Else

        End If
        cmd.ExecuteNonQuery()
    End Sub

    Private Sub CmbSBU_ext_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbSBU_ext.SelectedIndexChanged
        Dim tblAL As DataTable = GetDataTable(myStaff.ALList, Conn)

        For Each objRow As DataRow In tblAL.Rows
            LstAL_ext.Items.Add(objRow.Item(0))
        Next

    End Sub
End Class