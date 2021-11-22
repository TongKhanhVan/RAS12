Public Class BSP_Agent
    Private MyCust As New objCustomer
    Private WhoIs As String
    Sub New(pWhois As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        WhoIs = pWhois
    End Sub
    Private Sub BSP_Agent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyCust.GenCustList()
        LoadCmb_VAL(Me.CmbAgt, MyCust.List_TA)
        LoadGridCode()
        If WhoIs = "SALES" Then
            Me.TabControl1.TabPages.Remove(TabControl1.TabPages("TabPage2"))
        Else
            Me.TabControl1.TabPages.Remove(TabControl1.TabPages("TabPage1"))
        End If
    End Sub
    Private Sub ChkNewOnly_Click(sender As Object, e As EventArgs) Handles ChkNewOnly.Click
        LoadGridCode()
    End Sub
    Private Sub LoadGridCode()
        Dim strQry As String = "Select VAL as IATANO,VAL1 as CustShortName, VAL2 as City, RecID, details as FullName, description as Address" & _
            " from MISC where cat='BSPAGT' and "
        If WhoIs = "ACC" Then strQry = strQry & " Left(val1,3) in ('HAN','SGN') and len(val1)=7 and "
        If Me.ChkNewOnly.Checked Then
            Me.GridCode.DataSource = GetDataTable(strQry & " VAL1=''")
        Else
            Me.GridCode.DataSource = GetDataTable(strQry & " VAL1 <>''")
        End If
        Me.LblHAN.Visible = False
        Me.LblSGN.Visible = False
        Me.LblUpdate.Visible = False
        Me.LblUpdateTax.Visible = False
    End Sub
    Private Sub GridCode_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridCode.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If Me.ChkNewOnly.Checked Then
            Me.LblHAN.Visible = True
            Me.LblSGN.Visible = True
        Else
            Me.TxtFullName.Text = ""
            Me.TxtAddress.Text = ""
            Me.TxtTaxCode.Text = ""
            If Me.GridCode.CurrentRow.Cells("CustShortName").Value.ToString.Substring(0, 3) = "HAN" Or _
                Me.GridCode.CurrentRow.Cells("CustShortName").Value.ToString.Substring(0, 3) = "SGN" Then
                Me.LblUpdateTax.Visible = True
                Me.TxtFullName.Text = Me.GridCode.CurrentRow.Cells("FullName").Value
                If Me.GridCode.CurrentRow.Cells("Address").Value.ToString.Contains("|") Then
                    Me.TxtAddress.Text = Me.GridCode.CurrentRow.Cells("Address").Value.ToString.Split("|")(0)
                    Me.TxtTaxCode.Text = Me.GridCode.CurrentRow.Cells("Address").Value.ToString.Split("|")(1)
                End If
            Else
                Me.LblUpdateTax.Visible = False
            End If
        End If
    End Sub
    Private Sub LblUpdate_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        UpdateIATACode(Me.CmbAgt.Text, "SGN")
    End Sub
    Private Sub UpdateIATACode(pShortName As String, pCity As String)
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update MISC set VAL1='" & pShortName & "', VAL2='" & pCity & "' where RecID=" & _
            Me.GridCode.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridCode()
    End Sub
    Private Sub CmbAgt_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbAgt.SelectedIndexChanged
        If Me.ChkNewOnly.Checked Then Me.LblUpdate.Visible = True
    End Sub
    Private Sub LblHAN_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblHAN.LinkClicked
        UpdateIATACode("HAN" + Me.GridCode.CurrentRow.Cells(0).Value.ToString.Substring(4, 4), "HAN")
    End Sub

    Private Sub LblSGN_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSGN.LinkClicked
        UpdateIATACode("SGN" + Me.GridCode.CurrentRow.Cells(0).Value.ToString.Substring(4, 4), "SGN")
    End Sub
    Private Sub LblUpdateTax_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdateTax.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        cmd.CommandText = "update MISC set Details =@FullName, Description=@Address_Tax where RecID=" & Me.GridCode.CurrentRow.Cells("RecID").Value
        cmd.Parameters.Add("@FullName", SqlDbType.NVarChar).Value = Me.TxtFullName.Text
        cmd.Parameters.Add("@Address_Tax", SqlDbType.NVarChar).Value = Me.TxtAddress.Text & "|" & Me.TxtTaxCode.Text
        cmd.ExecuteNonQuery()
        Me.TxtFullName.Text = ""
        Me.TxtAddress.Text = ""
        Me.TxtTaxCode.Text = ""
        LoadGridCode()
    End Sub
End Class