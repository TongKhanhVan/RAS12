'Imports System.Object
Imports System.IO
Imports System.Data.SqlClient

Public Class frmUploadFile4NonAir
    Private mintFileId As Integer
    Private mstrFileType As String
    Private mobjRow As DataGridViewRow
    Private mblnUpload As Boolean

    ''' <summary>
    ''' Hiện bảng duyệt file để Upload, trả về giá trị integer:
    ''' </summary>
    ''' <returns>
    ''' = -1: upload xảy ra lỗi
    ''' =  0: nếu bấm cancel
    ''' = id: id file vừa upload
    ''' </returns>
    Public Function UploadFile() As Integer
        Try
            Using op As New OpenFileDialog()
                op.Filter = "Format files (*.pdf)|*.pdf|Image Files (*.jpg, *.png)|*.jpg;*.png"
                op.InitialDirectory = "C:"
                If op.ShowDialog() = DialogResult.OK Then
                    Return do_Upload(op.FileName)
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
            Return -1
        End Try
        Return 0
    End Function


    Private Function do_Upload(filename As String) As Integer
        Dim ms As MemoryStream = Nothing
        Dim filetype As String = ""

        Using fs As FileStream = New FileStream(filename, FileMode.Open, FileAccess.Read)
            If CountCharacter(filename, "."c) > 1 Then
                MsgBox("File name must contain NO DOT Character.", MsgBoxStyle.Critical)
                Return -1
            End If
            filetype = filename.Substring(filename.IndexOf("."c) + 1).ToLower()
            Dim bytes(fs.Length) As Byte
            fs.Read(bytes, 0, fs.Length)
            ms = New MemoryStream(bytes)
        End Using

        Dim conn = New SqlConnection(" Data Source=113.161.73.106;Initial Catalog=RAS12;Integrated Security=False;User ID=rasusers;Password=Healthy@FoodR12;Connect Timeout=15; ")
        Dim cmd = New SqlCommand("INSERT INTO [dbo].[Images] ([Image],[FileType]) VALUES (@Image,@FileType);select scope_identity();", conn)
        cmd.Parameters.Add(New SqlParameter("@Image", ms.ToArray))
        cmd.Parameters.Add(New SqlParameter("@FileType", filetype))

        conn.Open()
        Dim id As Integer = cmd.ExecuteScalar()
        conn.Close()
        cmd.Dispose()
        ms.Dispose()
        Return id
    End Function


    Private Sub lbkOK_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkOK.LinkClicked

        If mblnUpload Then
            Dim intFileId As Integer = UploadFile()
            Select Case intFileId
                Case 0, -1
                Case Else
                    Dim strUpdatedField As String = IIf(cboFileType.Text = "Registration", "FileId", "QuotationFile")

                    If ExecuteNonQuerry("Update DuToan_Tour set " & strUpdatedField & "=" & intFileId & " where RecId=" _
                         & mobjRow.Cells("RecId").Value, Conn) Then
                        mintFileId = intFileId
                        mstrFileType = cboFileType.Text
                        Me.DialogResult = Windows.Forms.DialogResult.OK
                    End If

            End Select
        ElseIf cboFileType.Text = "Quotation" AndAlso mobjRow.Cells("QuotationFile").Value > 0 Then
            load_UploadedFile(mobjRow.Cells("QuotationFile").Value)
        ElseIf cboFileType.Text = "Registration" AndAlso mobjRow.Cells("FileId").Value > 0 Then
            load_UploadedFile(mobjRow.Cells("FileId").Value)
        End If


        Me.Dispose()
    End Sub
    Private Sub lbkCancel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkCancel.LinkClicked
        Me.Dispose()
    End Sub

    Public Sub New(objRow As DataGridViewRow, blnUpload As Boolean)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mobjRow = objRow
        mblnUpload = blnUpload
        If Not blnUpload Then
            Me.Text = "View Uploaded File"
            If mobjRow.Cells("QuotationFile").Value > 0 Then
                cboFileType.SelectedIndex = 1
            Else
                cboFileType.SelectedIndex = 0
            End If

        End If
    End Sub

    Public Sub load_UploadedFile(ByVal RecID As Integer)
        Try
            Using file As New FileStream("C:\" & RecID & ".pdf", FileMode.Create, System.IO.FileAccess.Write)
                Dim conn = New SqlConnection(" Data Source=113.161.73.106;Initial Catalog=RAS12;Integrated Security=False;User ID=rasusers;Password=Healthy@FoodR12;Connect Timeout=15; ")

                Dim cmd = New SqlCommand("select [Image] from [dbo].[Images] where recid = @RecID", conn)
                cmd.Parameters.Add(New SqlParameter("@RecID", RecID))
                conn.Open()
                Dim data As Byte() = cmd.ExecuteScalar()
                conn.Close()
                cmd.Parameters.Clear()

                cmd = New SqlCommand("select [FileType] from [dbo].[Images] where recid = @RecID", conn)
                cmd.Parameters.Add(New SqlParameter("@RecID", RecID))
                conn.Open()
                Dim filetype As String = cmd.ExecuteScalar()
                conn.Close()

                cmd.Dispose()

                file.Write(data, 0, data.Length)

                Process.Start("C:\" & RecID & "." & filetype)
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    'Tìm và đếm số ký tự trong chuỗi
    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = ch Then cnt += 1
        Next
        Return cnt
    End Function


    Public Property FileId As Integer
        Get
            Return mintFileId
        End Get
        Set(value As Integer)
            mintFileId = value
        End Set
    End Property
    Public Property FileType As String
        Get
            Return mstrFileType
        End Get
        Set(value As String)
            mstrFileType = value
        End Set
    End Property
End Class