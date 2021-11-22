Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.IO
Imports System.Drawing.Printing.PrintDocument

Public Class MyDoc
    Private fName As String = ""
    Private FileToBeEmailed As String
    Private Sub MyDoc_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            Kill("d:\XXX*.*")
        Catch ex As Exception

        End Try
        Me.Dispose()
    End Sub
    Private Sub MyDoc_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCmb_MSC(Me.CmbCat, "select distinct CAT as VAL from DocList where status='OK'")
        If myStaff.Counter = "N-A" Then
            Me.CmbCat.Text = "VISA"
        ElseIf myStaff.Counter = "CWT" Then
            Me.CmbCat.Text = "OPS KIT"
        Else
            Me.CmbCat.Text = "QUITRINH"
        End If
        LoadTreeView("")
        Me.txtSearch.Width = Me.TreeRight.Width
        Me.LblSearch.Left = Me.txtSearch.Width + 8
    End Sub
    Private Sub LoadTreeView(pDK As String)
        Dim tblOng As DataTable, tblCha As DataTable, tblCon As DataTable, tblChau As DataTable, tblChat As DataTable
        Dim Ong As String, Cha As String, Con As String, Chau As String, Chat As String
        Dim strDKOng As String = " as VAL from DocList where Status='OK' and Cat='" & Me.CmbCat.Text & "'" & pDK
        Dim strDKCha As String, strDKCon As String, strDKChau As String, strDKChat As String

        Me.TreeRight.Nodes.Clear()
        tblOng = GetDataTable("select distinct DocCat " & strDKOng)
        For o As Int16 = 0 To tblOng.Rows.Count - 1
            Ong = tblOng.Rows(o)("VAL")
            Me.TreeRight.Nodes.Add(Ong, Ong)
            strDKCha = strDKOng & " and DocCat='" & Ong & "'"
            tblCha = GetDataTable("select distinct SubCat " & strDKCha)
            For i As Int16 = 0 To tblCha.Rows.Count - 1
                Cha = tblCha.Rows(i)("VAL")
                If Cha <> "" Then Me.TreeRight.Nodes(Ong).Nodes.Add(Cha, Cha)
                strDKCon = strDKCha & " and SubCat='" & Cha & "' and DocType<>'' "
                tblCon = GetDataTable("select distinct DocType " & strDKCon)
                For j As Int16 = 0 To tblCon.Rows.Count - 1
                    Con = tblCon.Rows(j)("VAL")
                    Me.TreeRight.Nodes(Ong).Nodes(Cha).Nodes.Add(Con, Con)
                    strDKChau = strDKCon & " and DocType='" & Con & "' and SubType<>''"
                    tblChau = GetDataTable("select distinct SubType " & strDKChau)
                    If tblChau.Rows.Count > 0 Then
                        For k As Int16 = 0 To tblChau.Rows.Count - 1
                            Chau = tblChau.Rows(k)("VAL")
                            strDKChat = strDKChau & " and SubType='" & Chau & "' and DocName<>'' "
                            Me.TreeRight.Nodes(Ong).Nodes(Cha).Nodes(Con).Nodes.Add(Chau, Chau)
                            tblChat = GetDataTable("select distinct DocName " & strDKChat)
                            If tblChat.Rows.Count > 0 Then
                                For l As Int16 = 0 To tblChat.Rows.Count - 1
                                    Chat = tblChat.Rows(l)("VAL")
                                    If Chat <> "" Then Me.TreeRight.Nodes(Ong).Nodes(Cha).Nodes(Con).Nodes(Chau).Nodes.Add(Chat, Chat)
                                Next
                            End If
                        Next
                    End If
                Next
            Next

        Next
    End Sub
    Private Function DownLoadFile(pNodePath As String, pWhat As String) As String
        Dim fieldName As String = "DocCat_SubCat_DocType_SubType_DocName", fExt As String
        Dim fileKQ As String = ""
        Dim strDK As String = " where CAT='" & Me.CmbCat.Text & "' ", tbl As DataTable
        Dim j As Int16 = pNodePath.Split("\").Length
        For i As Int16 = 0 To j - 1
            strDK = strDK & " and " & fieldName.Split("_")(i) & "='" & pNodePath.Split("\")(i) & "'"
        Next
        If j < fieldName.Split("_").Length Then strDK = strDK & " and " & fieldName.Split("_")(j) & "=''"
        tbl = GetDataTable("select top 1 FileType, DocContent, htmContent from DocList " & strDK)
        On Error Resume Next
        If pWhat = "V" Then Kill("d:\XXX*.*")
        If tbl.Rows.Count = 0 Then GoTo ExitHere
        fExt = tbl.Rows(0)("FileType").ToString.ToLower
        fName = "d:\XXX" & RandomString(8)
        If pWhat = "V" Then
            fileKQ = fName & IIf(fExt.Contains("pdf"), "." & fExt, ".mht")
            File.WriteAllBytes(fileKQ, tbl.Rows(0)("htmContent"))
        Else
            fileKQ = "d:\XXX" & pNodePath.Replace("\", "_") & "." & fExt
            File.WriteAllBytes(fileKQ, tbl.Rows(0)("DocContent"))
        End If
ExitHere:
        On Error GoTo 0
        Return fileKQ
    End Function
    Private Sub TreeRight_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeRight.AfterSelect
        Dim NodePath As String = e.Node.FullPath, FileKQ As String
        FileKQ = DownLoadFile(NodePath, "V")
        If FileKQ = "" Then
            Me.WebDisplay.Navigate("about:blank")
        Else
            If FileKQ.Contains(".pdf") Or FileKQ.Contains(".PDF") Then
                'Process.Start(FileKQ).WaitForExit(4)
                Process.Start(FileKQ)
            Else
                Me.WebDisplay.Navigate(FileKQ)
            End If
        End If
    End Sub
    Private Sub CmbCat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCat.SelectedIndexChanged
        LoadTreeView("")
    End Sub
    Private Sub LblEmail_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblEmail.LinkClicked
        If Me.CmbCat.Text <> "VISA" Or Me.TxtName.Text = "" Or InStr(Me.txtMail.Text, "@") = 0 Then Exit Sub
        FileToBeEmailed = ""
        DuyetNode2FindChecked(Me.TreeRight.Nodes)
        If FileToBeEmailed.Length > 2 Then
            Dim mailMsg As String = "Dear " & Me.TxtName.Text & ";"
            mailMsg = mailMsg & vbCrLf & "Further to our conversation, attached herewith is information for your Visa application. Please feel free to contact us if any clarification needed."
            mailMsg = mailMsg & vbCrLf & "Best regards"
            Try
                Dim myOL As New Microsoft.Office.Interop.Outlook.Application
                Dim oMsg As Microsoft.Office.Interop.Outlook.MailItem
                oMsg = myOL.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
                For i As Int16 = 0 To Me.txtMail.Text.Split(";").Length - 1
                    oMsg.Recipients.Add(Me.txtMail.Text.Split(";")(i))
                Next
                oMsg.Recipients.ResolveAll()
                oMsg.Subject = Me.TxtName.Text & " Visa Requirement"
                oMsg.Body = mailMsg
                For i As Int16 = 1 To FileToBeEmailed.Split("|").Length - 1
                    oMsg.Attachments.Add(FileToBeEmailed.Split("|")(i))
                    File.Delete(FileToBeEmailed.Split("|")(i))
                Next
                oMsg.Send()
            Catch ex As Exception
                MsgBox("Failure Sending Mail To Counter", MsgBoxStyle.Critical, msgTitle)
            End Try
            Me.txtMail.Text = "EMAIL"
            Me.txtMail.ForeColor = Color.DarkGray
            Me.TxtName.Text = ""
        End If
    End Sub
    Private Sub DuyetNode2FindChecked(ByVal pnode As TreeNodeCollection)
        Dim n As TreeNode, KQ As String
        For Each n In pnode
            If n.Checked Then
                KQ = DownLoadFile(n.FullPath, "E")
                If KQ <> "" Then FileToBeEmailed = FileToBeEmailed & "|" & KQ
            End If
            If n.Nodes.Count > 0 Then
                DuyetNode2FindChecked(n.Nodes)
            End If
        Next
    End Sub
    Private Sub txtMail_Enter(sender As Object, e As EventArgs) Handles txtMail.Enter
        If Me.txtMail.Text.ToUpper = "EMAIL" Then
            Me.txtMail.Text = ""
            Me.txtMail.ForeColor = Color.Black
        End If
    End Sub
    Private Sub LblSearch_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSearch.LinkClicked
        Dim strDK As String = ""
        If Me.txtSearch.Text = "" Then Exit Sub
        For i As Int16 = 0 To Me.txtSearch.Text.Split(" ").Length - 1
            strDK = strDK & "or keyword like '%" & Me.txtSearch.Text.Split(" ")(i) & "%' "
        Next
        strDK = " and (" & strDK.Substring(3) & ")"
        LoadTreeView(strDK)
        Me.txtSearch.Text = ""
    End Sub
    Private Function RandomString(ByVal length As Integer) As String
        Dim random As New Random(), AscCode As Int16
        Dim KQ As String = ""
        For i As Integer = 0 To length - 1
            If i / 2 = Int(i / 2) Then
                AscCode = random.Next(65, 90)
            Else
                AscCode = random.Next(97, 122)
            End If
            KQ = KQ & Chr(AscCode)
        Next
        Return KQ
    End Function
End Class