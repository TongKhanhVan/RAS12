Public Class ExportToTSP
    Private RCPNO As String, Curr As String
    Dim MaxDOF As String = Format(DateAdd(DateInterval.Day, 32, Now), "dd-MMM-yy")
    Private StrDK_TKT As String = " qty<>0 and Status ='OK' and (left(TKNO,1)='Z' or dof >'" & MaxDOF & "') "
    Private Sub ExportToTSP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim LstDate As Date = DateAdd(DateInterval.Day, -4, Now)
        Dim StrDK As String = " and FstUpdate >'" & Format(LstDate, "dd-MMM-yy") & "'"
        Me.GridTC.DataSource = GetDataTable("select distinct document as TourCode from FOP " & _
                                                 " where CustomerID=5102 and Document <>'' " & StrDK & _
                                                 " and RCPID in (select RcpID from TKT where " & StrDK_TKT & ")")
    End Sub
    Private Sub GridTC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridTC.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Dim StrDK As String = StrDK_TKT
        If Me.txtUser.Text.Replace(" ", "") <> "" Then
            StrDK = StrDK & " and fstuser in ('" & Me.txtUser.Text.Replace(" ", "").Replace(",", "','") & "')"
        End If
        Me.GridTKT.DataSource = GetDataTable("select TKNO, PaxName, PaxType, Fare, Tax, Charge, Itinerary, qty, Currency, RCPNO, DOI, DOF, FstUser from TKT " & _
                                              " where " & StrDK & " and RCPID in (select RCPID from FOP where customerid=5102 and Document='" & _
                                              Me.GridTC.CurrentRow.Cells(0).Value & "' and status<>'XX')")
        Me.GridTKT.Columns("Qty").Visible = False
        Me.GridTKT.Columns("Currency").Visible = False
        Me.GridTKT.Columns("PaxName").Width = 128
        Me.GridTKT.Columns("PaxType").Width = 32
        Me.GridTKT.Columns("Fare").Width = 64
        Me.GridTKT.Columns("Tax").Width = 64
        Me.GridTKT.Columns("Charge").Width = 64
        Me.GridTKT.Columns("Fare").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("Tax").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.GridTKT.Columns("Charge").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        ResetStatus()
    End Sub
    Private Sub ResetStatus()
        Me.StatusADL.Text = ""
        Me.StatusCHD.Text = ""
        Me.StatusINF.Text = ""
        Me.StatusTTL.Text = ""
        Me.LblExport.Visible = False
    End Sub
    Private Sub LblCheck_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCheck.LinkClicked
        Dim ADL As Int16, CHD As Int16, INF As Int16
        Dim TTL As Decimal
        RCPNO = ""
        Curr = ""
        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            If Me.GridTKT.Item(0, i).Value Then
                If RCPNO = "" Then
                    RCPNO = Me.GridTKT.Item("RCPNO", i).Value
                Else
                    If RCPNO <> Me.GridTKT.Item("RCPNO", i).Value Then
                        MsgBox("You Can Export Only 1 RCPNo Per Go", MsgBoxStyle.Critical, msgTitle)
                        Exit Sub
                    End If
                End If
                If Curr = "" Then Curr = Me.GridTKT.Item("Currency", i).Value
                If Me.GridTKT.Item("PaxType", i).Value = "ADL" Then ADL = ADL + Me.GridTKT.Item("Qty", i).Value
                If Me.GridTKT.Item("PaxType", i).Value = "CHD" Then CHD = CHD + Me.GridTKT.Item("Qty", i).Value
                If Me.GridTKT.Item("PaxType", i).Value = "INF" Then INF = INF + Me.GridTKT.Item("Qty", i).Value
                TTL = TTL + (Me.GridTKT.Item("Fare", i).Value + Me.GridTKT.Item("Tax", i).Value) * Me.GridTKT.Item("Qty", i).Value + Me.GridTKT.Item("Charge", i).Value
            End If
        Next
        If ADL > 0 Then Me.StatusADL.Text = "ADL: " & ADL.ToString
        If CHD > 0 Then Me.StatusCHD.Text = "CHD: " & CHD.ToString
        If INF > 0 Then Me.StatusINF.Text = "INF: " & INF.ToString
        Me.StatusTTL.Text = "TTL Value: " & Format(TTL, "#,##0.00")
        If TTL > 0 Then Me.LblExport.Visible = True
    End Sub

    Private Sub GridTKT_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridTKT.CellContentClick
        If e.ColumnIndex = 0 Then ResetStatus()
    End Sub

    Private Sub LblCheckAll_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCheckAll.LinkClicked
        If Me.LblCheckAll.Text = "SelectAll" Then
            Me.LblCheckAll.Text = "DeSelectAll"
            CheckUnCheck(True)
        Else
            Me.LblCheckAll.Text = "SelectAll"
            CheckUnCheck(False)
        End If
    End Sub
    Private Sub CheckUnCheck(pStatus As Boolean)
        For i As Int16 = 0 To Me.GridTKT.Rows.Count - 1
            Me.GridTKT.Item(0, i).Value = pStatus
        Next
        ResetStatus()
    End Sub
    Private Sub LblExport_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblExport.LinkClicked
        Dim connFLX As New SqlClient.SqlConnection
        connFLX.ConnectionString = CnStr_FLX
        connFLX.Open()
        Dim cmdFLX As SqlClient.SqlCommand = connFLX.CreateCommand
        cmdFLX.CommandText = "select RecID from TOS_Cost Where SupplierID = 694 and SVC = 'AirAuto' and Status <> 'XX' and " & _
                            " TourCode = '" & Me.GridTC.CurrentRow.Cells("TourCode").Value & "' and Descr='" & RCPNO & "'"
        Dim ExistCostID As Integer = cmdFLX.ExecuteScalar
        If ExistCostID > 0 Then
            cmdFLX.CommandText = "select RecID from tos_pmt where costid=" & ExistCostID & " and status<>'XX'"
            If cmdFLX.ExecuteScalar > 0 Then ' Da in DNTT
                MsgBox("Payment Order Has Been Issued For This Batch. Export Aborted", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
            cmdFLX.CommandText = "Update TOS_Cost set status='XX', LstUser='" & myStaff.SICode & "', LstUpdate=getdate() where recID=" & ExistCostID & _
                                "; Update TOS_Bkg set status='XX', LstUser='" & myStaff.SICode & "', LstUpdate=getdate() where CostID=" & ExistCostID
            cmdFLX.ExecuteNonQuery()
        End If
        cmdFLX.CommandText = "Insert into TOS_Cost(isVATIncl, VAT, SupplierID, SVC, TourCode, Supplier, TTLCost, Curr, PmtMethod, IsEstimation, " & _
                "Fstuser, Descr) Values (1,0,694,'AirAuto','" & Me.GridTC.CurrentRow.Cells("TourCode").Value & "','TRAVEL SHOP'," & _
                CDec(Me.StatusTTL.Text.Substring(10)) & ",'" & Curr & "','ITP',1,'" & myStaff.SICode & "','" & RCPNO & "')" & _
                "; SELECT SCOPE_IDENTITY() AS [RecID]"
        Dim costID As Integer = cmdFLX.ExecuteScalar
        cmdFLX.CommandText = "insert into TOS_Bkg (TourCode, CostID, Qty, Cost, DIC, FstUser, Type1) values ('" & _
                            Me.GridTC.CurrentRow.Cells("TourCode").Value & "'," & costID & ",1," & CDec(Me.StatusTTL.Text.Substring(10)) & _
                            ",'CNTR','" & myStaff.SICode & "','" & RCPNO & "')"
        cmdFLX.ExecuteNonQuery()
        connFLX.Close()
        connFLX.Dispose()
        MsgBox("Sucessfully Exported to TSP. Please Check", MsgBoxStyle.Information, msgTitle)
        ResetStatus()
    End Sub

End Class