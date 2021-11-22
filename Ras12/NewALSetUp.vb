Public Class NewALSetUp
    Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand

    Private Sub TxtAL_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAL.LostFocus
        Dim tmpRecID As Int16, tmpStr As String
        tmpRecID = ScalarToInt("Airline", "RecID", " AL='" & Me.TxtAL.Text & "'")
        If tmpRecID > 0 Then
            Me.LblAdd.Enabled = False
            tmpStr = ScalarToString("Airline", "VAT", " AL='" & Me.TxtAL.Text & "'")
            Me.ChkBSP.Checked = IIf(tmpStr.Substring(2, 2) = "TY", True, False)
        Else
            Me.LblAdd.Enabled = True
        End If
        Me.LblSave.Enabled = Not Me.LblAdd.Enabled
        Me.OptGSA.Enabled = Me.LblAdd.Enabled
        Me.TxtMS.Enabled = Me.LblAdd.Enabled
        Me.TxtKH.Enabled = Me.LblAdd.Enabled
        Me.CmbCurr.Enabled = Me.LblAdd.Enabled
        Me.CmbTVCompany.Enabled = Me.LblAdd.Enabled
        Me.TxtALName.Text = ScalarToString("AirlineList", "ALName", " ALCode='" & Me.TxtAL.Text & "'")
        Me.TxtDocCode.Text = ScalarToString("AirlineList", "DocCode", " ALCode='" & Me.TxtAL.Text & "'")

    End Sub

    Private Sub OptGSA_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptGSA.Click
        Me.TxtMS.Enabled = True
        Me.TxtKH.Enabled = True
    End Sub

    Private Sub LblAdd_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblAdd.LinkClicked
        Dim tmpRecID As Integer
        If Me.OptGSA.Checked And (Me.TxtMS.Text.Length <> 3 Or Me.TxtKH.Text.Length <> 2) Then
            MsgBox("invalid MS/KH. Plz Check.", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        Else
            tmpRecID = ScalarToInt("Airline", "RecID", " Maso='" & Me.TxtMS.Text & "' or KyHieu='" & Me.TxtKH.Text & "'")
            If tmpRecID > 0 Then
                MsgBox("invalid MS/KH. Plz Check.", MsgBoxStyle.Critical, msgTitle)
                Exit Sub
            End If
        End If
        cmd.CommandText = "insert Airline (al, docCode, ALName, Country, isTVA, MultiCurr, VAT, DefauCurr, MauSo, KyHieu)" & _
            " values ('" & Me.TxtAL.Text & "','" & Me.TxtDocCode.Text & "','" & Me.TxtALName.Text & "','??'," & _
            IIf(Me.OptGSA.Checked, True, False) & ",false,'" & IIf(Me.OptGSA.Checked, "GY", "GN") & _
            IIf(Me.ChkBSP.Checked, "TY", "TN") & "','" & Me.CmbCurr.Text & "','" & Me.TxtMS.Text & "','" & Me.TxtKH.Text & "')"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "update MISC set details=details + ' " & Me.TxtAL.Text & "' where cat+Desdcription='TVCompany+" & _
            Me.CmbTVCompany.Text & "'"
        cmd.ExecuteNonQuery()
        Me.LblAdd.Enabled = False
    End Sub

    Private Sub LblSave_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSave.LinkClicked
        cmd.CommandText = "update Airline set VAT=left(VAT,2) + '" & IIf(Me.ChkBSP.Checked, "TY", "TN") & "' where AL='" & _
            Me.TxtAL.Text & "'"
        cmd.ExecuteNonQuery()
        Me.LblSave.Enabled = False
    End Sub

    Private Sub NewALSetUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.BackColor = pubVarBackColor
    End Sub
End Class