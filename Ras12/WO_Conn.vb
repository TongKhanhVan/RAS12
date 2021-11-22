Imports System.IO
Imports System
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.Drawing
Public Class MySharedFunctions

    Public Shared Function GetColorForToday() As Color
        Dim DayMau() As Color = {Color.YellowGreen, Color.LightGreen, Color.Cyan, Color.DeepSkyBlue, Color.DodgerBlue, _
                                 Color.SkyBlue, Color.Aquamarine, Color.YellowGreen, Color.YellowGreen}
        Return DayMau(Today.DayOfWeek)
    End Function

    Public Shared Function ImageToBytes(ByVal filepath As String) As Byte()
        Dim fs As New IO.FileStream(filepath, IO.FileMode.Open, IO.FileAccess.Read)
        Dim br As New BinaryReader(fs)
        Dim bytes As Byte() = br.ReadBytes(fs.Length)
        br.Close()
        fs.Close()
        Return bytes
    End Function

    Public Shared Function GetNextExeNameAvailable(pPath As String, pAppName As String) As String
        Dim KQ As String
        Dim FileNum As Int16 = FreeFile()
        Dim vStart As Int16 = Now.Day Mod 2
        For i As Int16 = vStart To 128 Step 2
            KQ = pPath & pAppName & "_" & i.ToString.Trim & ".exe"
            Try
                FileOpen(FileNum, KQ, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
                FileClose(FileNum)
                Return KQ
            Catch ex As Exception
            End Try
        Next
        Return ""
    End Function
    Public Shared Function UploadFileToFtp(ByVal pLocalFileName As String, pUploadPath As String, pFtpUserName As String, pFtpPSW As String, pReturnType As String, Optional ReturnPath As String = "") As String
        Dim _UploadPath As String = pUploadPath
        Dim _FTPUser As String = pFtpUserName
        Dim _FTPPass As String = pFtpPSW

        Dim _FileInfo As New System.IO.FileInfo(pLocalFileName)

        _UploadPath = _UploadPath & _FileInfo.Name

        Dim _FtpWebRequest As System.Net.FtpWebRequest = CType(System.Net.FtpWebRequest.Create(New Uri(_UploadPath)), System.Net.FtpWebRequest)
        With _FtpWebRequest
            .Credentials = New System.Net.NetworkCredential(_FTPUser, _FTPPass)
            .KeepAlive = False
            .Timeout = 20000
            .Method = System.Net.WebRequestMethods.Ftp.UploadFile
            .UseBinary = True
            .ContentLength = _FileInfo.Length
        End With
        Dim buffLength As Integer = 2048
        Dim buff(buffLength - 1) As Byte

        Dim _FileStream As System.IO.FileStream = _FileInfo.OpenRead()
        Try
            Dim _Stream As System.IO.Stream = _FtpWebRequest.GetRequestStream()
            Dim contentLen As Integer = _FileStream.Read(buff, 0, buffLength)
            Do While contentLen <> 0
                _Stream.Write(buff, 0, contentLen)
                contentLen = _FileStream.Read(buff, 0, buffLength)
            Loop
            _Stream.Close()
            _Stream.Dispose()
            _FileStream.Close()
            _FileStream.Dispose()
        Catch ex As Exception
            Throw ex
        End Try
        If pReturnType = "WEB" Then
            Return ReturnPath & _FileInfo.Name
        Else
            Return "True"
        End If
    End Function

    Public Shared Function GenDefaultPSW(ByVal pUserCode As String, ByVal AppCode As String) As String
        Return pUserCode.Substring(0, 1) & AppCode.Substring(0, 1) & pUserCode.Substring(1, 1) & _
            AppCode.Substring(1, 1) & pUserCode.Substring(2, 1) & AppCode.Substring(2, 1)
    End Function
    Public Shared Function AddSpace2TKNO(ByVal pDocNo As String) As String
        Dim tmp As String = ""
        If pDocNo = "" Then Return ""
        tmp = pDocNo.Replace(" ", "")
        If pDocNo.Substring(0, 1) = "Z" Then
            tmp = tmp.Substring(0, 3) & " " & tmp.Substring(3, 6) & " " & Strings.Right(tmp, 4)
        Else
            tmp = tmp.Substring(0, 3) & " " & tmp.Substring(3, 4) & " " & Strings.Right(tmp, 6)
        End If
        Return tmp
    End Function

    Public Shared Function CheckRTG(ByVal varRTG As String) As Boolean
        Dim nSeg As Int16, rtgLen As Int16
        varRTG = varRTG.Replace(" ", "")
        On Error GoTo ErrHandler
        rtgLen = varRTG.ToString.Length
        If rtgLen < 8 Then Return False
        nSeg = (rtgLen - 3) / 5
        If rtgLen > 0 Then
            If (rtgLen - 3) / 5 > nSeg OrElse _
                varRTG.Substring(3, 2) = "//" OrElse _
                varRTG.Substring(rtgLen - 4, 2) = "//" Then
                Return False
            End If
            For i As Int16 = 1 To nSeg
                If varRTG.Substring(5 * (i - 1), 3) = varRTG.Substring(5 * i, 3) Then
                    CheckRTG = False
                    Return False
                End If
            Next
        End If
        On Error GoTo 0
        Return True
ErrHandler:
        Return False
        Exit Function
    End Function

    Public Shared Function Deli2InClause(ByVal pDeliStr As String, ByVal Deli As String) As String
        Dim KQ As String = ""
        KQ = pDeliStr.Replace(Deli, "','")
        KQ = "('" & KQ & "')"
        Return KQ
    End Function
    Public Shared Function HashToFixedLen(ByVal pStr As String) As String
        Dim tmpSource() As Byte
        Dim tmpHash() As Byte

        tmpSource = ASCIIEncoding.ASCII.GetBytes(pStr)
        tmpHash = New MD5CryptoServiceProvider().ComputeHash(tmpSource)
        Dim sOutput As New StringBuilder(tmpHash.Length)
        For i As Int16 = 0 To tmpHash.Length - 1
            sOutput.Append(tmpHash(i).ToString("X2"))
        Next
        Return sOutput.ToString()
    End Function
    Public Shared Function checkCharEntered(ByVal varKeyVAL As Int16) As Boolean
        checkCharEntered = False
        If varKeyVAL < 48 OrElse varKeyVAL > 57 Then ' khac 0-9
            If varKeyVAL < 96 OrElse varKeyVAL > 105 Then ' khac numpad  0-9 
                If varKeyVAL < 109 OrElse varKeyVAL > 110 Then ' khac . -
                    If varKeyVAL < 189 OrElse varKeyVAL > 190 Then 'khac . -
                        If varKeyVAL <> 8 Then
                            checkCharEntered = True
                        End If
                    End If
                End If
            End If
        End If
    End Function
    Public Shared Function XDNGayCuoiThangTruoc(ByVal pThangNay As Date) As Date
        Dim KQ As Date
        KQ = DateSerial(Year(pThangNay), Month(pThangNay), 1)
        KQ = KQ.AddDays(-1)
        Return KQ
    End Function
    Public Shared Function EnCode(ByVal abc As String, ByVal pKey As String) As String
        Dim CDE As String = "", CharI As String
        For i As Int16 = 1 To Len(abc)
            CharI = Mid(abc, i, 1)
            If Asc(CharI) <> 13 And Asc(CharI) <> 10 Then
                CDE = CDE & Chr(Asc(CharI) Xor pKey)
            ElseIf Asc(CharI) = 10 Then
                CDE = CDE & vbCrLf
            End If
        Next
        EnCode = CDE
    End Function
    Public Shared Function AddSpace2Rtg(ByVal varRtg As String) As String
        Dim KQ As String = "", i As Int16 = 0, j As Int16 = 3
        Dim tmpStr As String
        tmpStr = varRtg.Replace(" ", "")
        Do While Len(tmpStr) > 2
            KQ = KQ & Left(tmpStr, j) & " "
            tmpStr = Mid(tmpStr, j + 1)
            j = IIf(j = 3, 2, 3)
        Loop
        KQ = KQ.Replace(" NRT ", " TYO ")
        KQ = KQ.Replace(" HND ", " TYO ")
        KQ = KQ.Replace(" EWR ", " NYC ")
        KQ = KQ.Replace(" JFK ", " NYC ")
        KQ = KQ.Replace(" IAH ", " HOU ")
        KQ = KQ.Replace(" ORD ", " CHI ")
        Return KQ.Trim
    End Function
    Public Shared Function Occur(ByVal pWhere As String, ByVal pWhat As String) As Int16
        Dim KQ As Int16 = 0
        For i As Int16 = 1 To Len(pWhere)
            If Mid(pWhere, i, 1) = pWhat Then
                KQ = KQ + 1
            End If
        Next
        Return KQ
    End Function
    Public Shared Function GetCN_String() As String
        Return "server=42.117.5.86;uid=ft;pwd=Healthy@Food;database=FT"
    End Function
    Public Shared Function GetServerName(ByVal pFName As String) As String
        Dim KQ As String = IO.File.ReadAllLines(pFName)(0)
        Return KQ.Split(";")(0).Split("=")(1)
    End Function
    Public Shared Function DefineMailConfig(ByVal pACC As String) As String
        Dim KQ As String
        If InStr(pACC, "@YAHOO.COM") > 0 Then
            KQ = "smtp.mail.yahoo.com|25|0"
        ElseIf InStr(pACC, "@GMAIL.COM") > 0 Then
            KQ = "smtp.gmail.com|587|0"
        ElseIf InStr(pACC, "@TRANSVIET.COM") > 0 Then
            KQ = "smtp.transviet.com|25|0"
        Else
            KQ = "smtp.live.com|587|-1"
        End If
        Return KQ
    End Function
End Class
