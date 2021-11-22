Imports SharedFunctions.MySharedFunctions
Public Class NewsUpdater
    Dim cmd_web As SqlClient.SqlCommand = Conn_TVVN.CreateCommand
    Private Sub GridNew_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridNew.CellContentClick
        Me.LblRemove.Visible = True
        Me.TxtSubj.Text = Me.GridNew.CurrentRow.Cells("Subj").Value
        Me.TxtFull.Text = Me.GridNew.CurrentRow.Cells("FullText").Value
        Me.CmbPos.Text = Me.GridNew.CurrentRow.Cells("Position").Value
    End Sub

    Private Sub LblRemove_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblRemove.LinkClicked
        cmd_web.CommandText = ChangeStatus_ByID("News", "XX", Me.GridNew.CurrentRow.Cells("recID").Value)
        cmd_web.ExecuteNonQuery()
        LoadGridNews()
    End Sub

    Private Sub NewsUpdater_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Conn_TVVN.Close()
    End Sub

    Private Sub NewsUpdater_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Conn_TVVN.ConnectionString = CnStr_TVVN
        Conn_TVVN.Open()
        Me.BackColor = pubVarBackColor
        LoadGridNews()
    End Sub
    Private Sub LoadGridNews()
        Me.GridNew.DataSource = GetDataTable("Select * from News where status='OK'", Conn_TVVN)
        Me.GridNew.Columns(0).Visible = False
    End Sub

    Private Sub LblUpdate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUpdate.LinkClicked
        Dim pictID0 As Integer
        cmd_web.Parameters.Clear()
        cmd_web.Parameters.Add("@HinhAnh", SqlDbType.Binary)
        If Me.txtPic0.Text = "" Then
            MsgBox("You Have to Specify Path Name for Glance Picture", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        Try
            cmd_web.CommandText = "insert HinhAnh (HinhAnh) values (@HinhAnh); SELECT SCOPE_IDENTITY() AS [RecID]"
            cmd_web.Parameters("@HinhAnh").Value = ImageToBytes(Me.txtPic0.Text)
            pictID0 = cmd_web.ExecuteScalar
        Catch ex As Exception
        End Try
        cmd_web.CommandText = "insert News (Subj, FullText, Position, Pict0, FstUser, html,City) values " & _
            "(@Subj, @FullText, @Position, @Pict0, @FstUser, @html, @City)"
        cmd_web.Parameters.Clear()
        cmd_web.Parameters.Add("@Subj", SqlDbType.NVarChar).Value = Me.TxtSubj.Text
        cmd_web.Parameters.Add("@FullText", SqlDbType.NVarChar).Value = Me.TxtFull.Text
        cmd_web.Parameters.Add("@Position", SqlDbType.VarChar).Value = Me.CmbPos.Text
        cmd_web.Parameters.Add("@Pict0", SqlDbType.Int).Value = pictID0
        cmd_web.Parameters.Add("@html", SqlDbType.NText).Value = txtHTML.DocumentHtml
        cmd_web.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd_web.Parameters.Add("@City", SqlDbType.VarChar).Value = IIf(Me.ChkALLACO.Checked, "ALL", myStaff.City)
        cmd_web.ExecuteNonQuery()
        LoadGridNews()
        MsgBox("News Has Been Updated", MsgBoxStyle.Information, msgTitle)
    End Sub
    Private Sub cmdPict0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        cmdPict0.Click, cmdPict1.Click, cmdPict2.Click, cmdPict3.Click
        Dim cmd As Button = CType(sender, Button)
        Dim i As String, KQ As String
        i = cmd.Name.Substring(7, 1)
        Me.OpenFileDialog1.ShowDialog()
        KQ = Me.OpenFileDialog1.FileName
        Me.TabPage1.Controls("txtPic" & i).Text = KQ
    End Sub
    Private Sub LblUploadPict_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUploadPict.LinkClicked
        Dim Pict1ID As Integer, Pict2ID As Integer, Pict3ID As Integer
        Try
            cmd_web.CommandText = "insert HinhAnh (HinhAnh) values (@HinhAnh); SELECT SCOPE_IDENTITY() AS [RecID]"
            cmd_web.Parameters.Clear()
            cmd_web.Parameters.Add("@HinhAnh", SqlDbType.Binary)
            cmd_web.Parameters("@HinhAnh").Value = ImageToBytes(Me.TxtPic1.Text)
            Pict1ID = cmd_web.ExecuteScalar
            cmd_web.Parameters("@HinhAnh").Value = ImageToBytes(Me.TxtPic2.Text)
            Pict2ID = cmd_web.ExecuteScalar
            cmd_web.Parameters("@HinhAnh").Value = ImageToBytes(Me.TxtPic3.Text)
            Pict3ID = cmd_web.ExecuteScalar
        Catch ex As Exception

        End Try
        Me.TxtPic1.Text = "http://transviet.vn/GetImage.aspx?ID=" & Pict1ID.ToString
        Me.TxtPic2.Text = "http://transviet.vn/GetImage.aspx?ID=" & Pict2ID.ToString
        Me.TxtPic3.Text = "http://transviet.vn/GetImage.aspx?ID=" & Pict3ID.ToString
    End Sub

    Private Sub LblUploadFile_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUploadFile.LinkClicked
        Dim f As New OpenFileDialog, ReturnPath As String
        f.Multiselect = False
        f.ShowDialog()
        If f.FileName <> "" Then
            ReturnPath = UploadFileToFtp(f.FileName, "ftp://transviet.vn/Upload/", "transviet.vn", "Abcd1234", "WEB", "http://transviet.vn/Upload/")
            My.Computer.Clipboard.SetText(ReturnPath)
        End If
    End Sub
End Class