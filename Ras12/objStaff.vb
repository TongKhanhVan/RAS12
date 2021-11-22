
Public Class objStaff
    Private _SICode As String, _ShortName As String, _PSW As String, _DAccess As String, _UID As Integer
    Private _UGroup As String, _SupOf As String, _City As String, _App As String, _CreateDate As Date
    Private _AAccess As String, _CAccess As String = "", _Counter As String, _Location As String
    Private _CurrObj As String, _URights As String, _CnStr As String
    Private _TVA As String, _GSA As String, _ALList As String, _emailAddress As String
    Private StrSQL As String
    Private _SelectedDomain As String

    ReadOnly Property TVA() As String
        Get
            Return _TVA
        End Get
    End Property
    ReadOnly Property Location() As String
        Get
            Return _Location
        End Get
    End Property

    ReadOnly Property Counter() As String
        Get
            Return _Counter
        End Get
    End Property

    ReadOnly Property GSA() As String
        Get
            Return _GSA
        End Get
    End Property
    ReadOnly Property EmailAddress() As String
        Get
            Return _emailAddress
        End Get
    End Property

    ReadOnly Property ALList() As String
        Get
            Return _ALList
        End Get
    End Property
    Property App() As String
        Get
            Return _App
        End Get
        Set(ByVal sApp As String)
            _App = sApp
        End Set
    End Property
    Property CurrObj() As String
        Get
            Return _CurrObj
        End Get
        Set(ByVal sCurrObj As String)
            _CurrObj = sCurrObj
            _URights = ""
            Dim dTable As DataTable
            dTable = GetDataTable(String.Format("select SubObject from tblRight where status='OK' and " & _
                " SICode='{0}' and upper(object)='{1}'", _SICode, UCase(_CurrObj)))
            For i As Int16 = 0 To dTable.Rows.Count - 1
                _URights = _URights & "|" & dTable.Rows(i)("SubObject")
            Next
            If _URights.Length > 2 Then _URights = _URights.Substring(1)
        End Set
    End Property

    Property CnStr() As String
        Get
            Return _CnStr
        End Get
        Set(ByVal sCnStr As String)
            _CnStr = sCnStr
        End Set
    End Property
    Public Property SelectedDomain As String
        Get
            Return _SelectedDomain
        End Get
        Set(value As String)
            _SelectedDomain = value
        End Set
    End Property

    Property SICode() As String
        Get
            Return _SICode
        End Get
        Set(ByVal sSICode As String)
            _SICode = sSICode
            If SICode = "" Then
                _UID = 0
                Exit Property
            End If
            _DAccess = "EDU"
            Dim dTable As DataTable, tmpStr As String

            dTable = GetDataTable("select * from tblUser where SICode='" & _SICode & "' and status<>'XX'")
            For i As Int16 = 0 To dTable.Rows.Count - 1
                _UID = dTable.Rows(i)("RecID")
                _shortName = dTable.Rows(i)("SIName")
                _PSW = dTable.Rows(i)("PSW")
                _UGroup = dTable.Rows(i)("Template")
                _City = dTable.Rows(i)("City")
                _SupOf = dTable.Rows(i)("SupOf")
                _CreateDate = dTable.Rows(i)("FstUpdate")
                _Counter = dTable.Rows(i)("Counter")
                _Location = dTable.Rows(i)("Location")
            Next
            If _UID > 0 Then
                dTable = GetDataTable("select * from UserAccess where UserID='" & _UID & "' and status<>'XX' and app='" & _App & "'")
                For i As Int16 = 0 To dTable.Rows.Count - 1
                    If dTable.Rows(i)("CAT") = "ChannelAccess" Then
                        tmpStr = dTable.Rows(i)("VAL")
                        tmpStr = tmpStr.Replace("-", "_")
                        tmpStr = tmpStr.Replace("_", "','")
                        _CAccess = "('" & tmpStr & "')"
                    End If
                    If dTable.Rows(i)("CAT") = "ALAccess" Then
                        tmpStr = dTable.Rows(i)("VAL")
                        tmpStr = tmpStr.Replace("-", "_")
                        tmpStr = tmpStr.Replace("_", "','")
                        _AAccess = "('" & tmpStr & "')"
                    End If
                    If dTable.Rows(i)("CAT") = "DomainAccess" Then _DAccess = dTable.Rows(i)("VAL")
                Next
            End If

            _ALList = "select AL as VAL from Airline where AL not in ('XX','TS') "
            _TVA = "select AL as VAL from Airline where isTVA<>0 "
            _GSA = "select AL as VAL from Airline where isTVA<>0 and AL not in ('XX','TS') "
            If _AAccess = "('01')" Then
                _AAccess = "('01','TS')"
            End If
            If InStr(_AAccess, "YY") = 0 Then
                _ALList = _ALList & " and AL in " & _AAccess
                _TVA = _TVA & " and AL in " & _AAccess
                _GSA = _GSA & " and AL in " & _AAccess
            End If
            _emailAddress = GetEmailAddress()
        End Set
    End Property
    ReadOnly Property ShortName() As String
        Get
            Return _shortName
        End Get
    End Property
    ReadOnly Property PSW() As String
        Get
            Return _PSW
        End Get
    End Property
    ReadOnly Property DAccess() As String
        Get
            Return _DAccess
        End Get
    End Property

    ReadOnly Property CAccess() As String
        Get
            Return _CAccess
        End Get
    End Property

    ReadOnly Property AAccess() As String
        Get
            Return _AAccess
        End Get
    End Property
    ReadOnly Property URights() As String
        Get
            Return _URights
        End Get
    End Property

    ReadOnly Property UGroup() As String
        Get
            Return _UGroup
        End Get
    End Property
    Property City() As String
        Get
            Return _City
        End Get
        Set(ByVal sCity As String)
            _City = sCity
        End Set
    End Property

    ReadOnly Property SupOf() As String
        Get
            Return _SupOf
        End Get
    End Property
    ReadOnly Property UID() As String
        Get
            Return _UID
        End Get
    End Property
    ReadOnly Property CreateDate() As Date
        Get
            Return _CreateDate
        End Get
    End Property
    Public Sub LogOut()
        Dim StaffConn As New SqlClient.SqlConnection
        Dim cmd As SqlClient.SqlCommand = StaffConn.CreateCommand
        StaffConn.ConnectionString = _CnStr
        StaffConn.Open()
        cmd.CommandText = ChangeStatus_ByDK("tblUser", "OK", "where SICode='" & _SICode & "' and status='ON'")
        cmd.ExecuteNonQuery()
        StaffConn.Close()
        StaffConn.Dispose()
        _SICode = ""
    End Sub
    Private Function GetEmailAddress()
        Dim KQ As String = ""
        'If _City <> "SGN" Then Return ""
        'If InStr(_CAccess, "CS") > 0 Then Return ""
        'Try
        '    Dim myOL As New Microsoft.Office.Interop.Outlook.Application
        '    Dim oNS As Microsoft.Office.Interop.Outlook.NameSpace = myOL.GetNamespace("mapi")
        '    Dim oMsg As Microsoft.Office.Interop.Outlook.MailItem
        '    Dim oInbox As Microsoft.Office.Interop.Outlook.MAPIFolder = oNS.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail)
        '    oMsg = oInbox.Items(1)
        '    KQ = oMsg.SenderName
        'Catch ex As Exception
        'End Try
        Return KQ
    End Function
End Class
