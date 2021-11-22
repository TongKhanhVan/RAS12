Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.IO
Imports System.Drawing.Printing.PrintDocument

Public Class frmQuickRef
    Private mstrCompanyName As String
    Private mblnFirstLoadCompleted
    Private Sub frmQuickRef_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCombo(cboCategory, "Select SubCat as value from DocList where Status='OK'" _
                  & " and Cat='OPS KIT' and DocCat='" & mstrCompanyName & "' order by SubCat", Conn)
        mblnFirstLoadCompleted = True
        Try
            cboCategory.SelectedIndex = 6
        Catch ex As Exception

        End Try


    End Sub

    Public Sub New(intCustId As Integer)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mstrCompanyName = ScalarToString("CompanyInfo", "CompanyName", "Status='OK' and CustId=" _
                                         & intCustId)
    End Sub
    Private Function DownLoadFile(strSubCat As String) As String

        Dim strQuerry As String = "Select top 1 * from DocList where Status='OK' and Cat='OPS KIT'" _
                  & " and DocCat='" & mstrCompanyName & "' and SubCat='" & strSubCat & "'"

        Dim strFileName As String = String.Empty
        Dim strFileExtension As String

        Dim tblResult As DataTable
        tblResult = GetDataTable(strQuerry, Conn)
        On Error Resume Next

        Kill("d:\XXX*.*")
        If tblResult.Rows.Count = 0 Then GoTo ExitHere
        strFileExtension = tblResult.Rows(0)("FileType").ToString.ToLower
        strFileName = "d:\XXX" & RandomString(8)

        strFileName = strFileName & IIf(strFileExtension.Contains("pdf"), "." & strFileExtension, ".mht")
        File.WriteAllBytes(strFileName, tblResult.Rows(0)("htmContent"))
        
ExitHere:
        On Error GoTo 0
        Return strFileName
    End Function
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

    Private Sub cboCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCategory.SelectedIndexChanged
        If mblnFirstLoadCompleted Then
            Dim strFileName As String
            strFileName = DownLoadFile(cboCategory.Text)
            If strFileName = "" Then
                Me.WebDisplay.Navigate("about:blank")
            Else
                If strFileName.Contains(".pdf") Or strFileName.Contains(".PDF") Then
                    'Process.Start(FileKQ).WaitForExit(4)
                    Process.Start(strFileName)
                Else
                    Me.WebDisplay.Navigate(strFileName)
                End If
            End If
        End If
    End Sub
End Class