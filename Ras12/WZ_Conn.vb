Imports System.IO
Imports SharedFunctions.MySharedFunctions
Imports SKYPE4COMLib
Public Class MySharedFunctionsWzConn
    Private StrSQL As String

    Public Shared Function SendSkypeAlert(pNick As String, pMsg As String) As Boolean
        Try
            If pNick = "" Or pMsg = "" Then Exit Function
            'Dim iSkype As New SKYPE4COMLib.Skype
            'iSkype.SendMessage(pNick, pMsg)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Public Shared Function UName_PSW_match(ByVal pSI As String, ByVal pPSW As String, ByVal pConn As SqlClient.SqlConnection) As Boolean
        Dim KQ As Integer
        Dim cmd As SqlClient.SqlCommand = pConn.CreateCommand
        cmd.CommandText = String.Format("select RecID from tblUser where status <>'XX' and SICode='{0}' and PSW='{1}'", pSI, HashToFixedLen(pPSW))
        KQ = cmd.ExecuteScalar
        Return IIf(KQ > 0, True, False)
    End Function
    Public Function EnCode(ByVal abc As String) As String
        Dim CDE As String = "", CharI As String
        For i As Int16 = 1 To Len(abc)
            CharI = Mid(abc, i, 1)
            If Asc(CharI) <> 13 And Asc(CharI) <> 10 Then
                CDE = CDE & Chr(Asc(CharI) Xor 115)
            ElseIf Asc(CharI) = 10 Then
                CDE = CDE & vbCrLf
            End If
        Next
        EnCode = CDE
    End Function
    Public Shared Function FieldToList(ByVal pField As String, ByVal pTable As String, ByVal pConn As SqlClient.SqlConnection, ByVal pDK As String) As String
        Dim KQ As String
        Dim cmd As SqlClient.SqlCommand = pConn.CreateCommand
        cmd.CommandText = "declare @FieldToList varchar(256); set @FieldToList =''; " & _
            "select @FieldToList = ',' + " & pField & " + @FieldToList from " & pTable & IIf(pDK <> "", " where " & pDK, "") & "; select @FieldToList"
        KQ = cmd.ExecuteScalar
        If KQ.Length > 1 Then KQ = KQ.Substring(1)
        Return KQ
    End Function
    Public Shared Function GetCutOverDate(ByVal pDelay As String, ByVal pConn As SqlClient.SqlConnection) As Date
        Dim KQ As String
        Dim cmd As SqlClient.SqlCommand = pConn.CreateCommand
        cmd.CommandText = String.Format("select VAL1 from MISC where cat='CUTOVER' and VAL='{0}' ", pDelay)
        KQ = cmd.ExecuteScalar
        Return CDate(KQ)
    End Function

    Public Shared Function GetSSMSfrmSQL(ByVal parAction As String, ByVal pConn As SqlClient.SqlConnection) As String
        Dim KQ As String
        Dim cmd As SqlClient.SqlCommand = pConn.CreateCommand
        cmd.CommandText = "select ltrim(rtrim(datepart(ss,getdate()))) +  ltrim(rtrim(datepart(ms,getdate())))  "
        KQ = cmd.ExecuteScalar
        KQ = KQ & parAction
        If KQ.Length > 5 Then KQ = KQ.Substring(0, 5)
        Return KQ
    End Function
    Public Shared Function GetHHMMfrmSQL(ByVal pConn As SqlClient.SqlConnection) As String
        Dim KQ As String
        Dim cmd As SqlClient.SqlCommand = pConn.CreateCommand
        cmd.CommandText = "select ltrim(rtrim(datepart(hh,getdate()))) +  ltrim(rtrim(datepart(mi,getdate()))) +  ltrim(rtrim(datepart(ss,getdate())))  "
        KQ = cmd.ExecuteScalar
        If KQ.Length > 5 Then KQ = KQ.Substring(0, 5)
        Return KQ
    End Function

    Public Shared Function SysDateIsWrong(ByVal pConn As SqlClient.SqlConnection) As Boolean
        Dim KQ As Boolean = False, SQLdate As Date
        Dim MinsSQL As Int16, MinsLocal As Int16
        Dim cmd As SqlClient.SqlCommand = pConn.CreateCommand
        cmd.CommandText = "select getdate() "
        SQLdate = cmd.ExecuteScalar
        MinsSQL = SQLdate.Hour * 60 + SQLdate.Minute
        MinsLocal = Now.Hour * 60 + Now.Minute
        If SQLdate.Date <> Now.Date Then
            KQ = True
        ElseIf MinsSQL - MinsLocal > 5 Or MinsLocal - MinsSQL > 5 Then
            KQ = True
        End If
        If KQ Then MsgBox("System Date/Time is Incorrect", MsgBoxStyle.Critical, "TransViet Travel :: System Message")
        Return KQ
    End Function
    Public Shared Function GetPubHoliday(ByVal pConn As SqlClient.SqlConnection) As String
        Dim cmd As SqlClient.SqlCommand = pConn.CreateCommand
        Dim KQ As String
        cmd.CommandText = "select STR1 from fox_MISC where cat='AMLICH'"
        KQ = cmd.ExecuteScalar
        KQ = KQ & "_01JAN_02SEP_30APR_01MAY"
        Return KQ
    End Function
    Public Shared Sub Insert_Update_WebTable(ByVal pConn_Web As SqlClient.SqlConnection, ByVal pQry As String)
        Dim cmd_Web As SqlClient.SqlCommand = pConn_Web.CreateCommand, i As Int16
        On Error GoTo errHandeler
        If Not pConn_Web.State = ConnectionState.Open Then pConn_Web.Open()
        If Not pConn_Web.State = ConnectionState.Open Then GoTo errHandeler
        cmd_Web.CommandText = pQry
        cmd_Web.ExecuteNonQuery()
        i = 1
errHandeler:
        On Error GoTo 0
    End Sub
    Public Shared Function InsertSMS(ByVal pNbr As String, ByVal pMsg As String, ByVal pPOS As String, ByVal pPRG As String) As String
        Dim SMSConn As New SqlClient.SqlConnection
        Dim smsCmd As SqlClient.SqlCommand = SMSConn.CreateCommand
        'SMSConn.ConnectionString = "server=mssql-transvietcom.transviet.com;uid=transviet;pwd=Abcd1234;database=transvietcom"
        SMSConn.ConnectionString = "server=42.117.5.86;uid=ft;pwd=ftwriter;database=FT"
        Try
            SMSConn.Open()
            smsCmd.CommandText = " Insert SMSLog (CustID, SMSText, Location, MobileNbr, PRG) values (@CustID, @SMSText, @Location, @MobileNbr, @PRG)"
            smsCmd.Parameters.Clear()
            smsCmd.Parameters.Add("@CustID", SqlDbType.Int).Value = -1
            smsCmd.Parameters.Add("@SMSText", SqlDbType.VarChar).Value = pMsg
            smsCmd.Parameters.Add("@Location", SqlDbType.VarChar).Value = pPOS
            smsCmd.Parameters.Add("@MobileNbr", SqlDbType.VarChar).Value = pNbr
            smsCmd.Parameters.Add("@PRG", SqlDbType.VarChar).Value = pPRG
            smsCmd.ExecuteNonQuery()
            SMSConn.Close()
        Catch ex As Exception
            Return "Err Connecting to SMSC"
        End Try
        Return "OK"
    End Function
    Public Shared Function InsertEmail(ByVal pCustID As Integer, ByVal pSubj As String, ByVal pMsg As String, ByVal pPOS As String, ByVal pPRG As String) As String
        Dim EmailConn As New SqlClient.SqlConnection
        Dim EmailCmd As SqlClient.SqlCommand = EmailConn.CreateCommand
        EmailConn.ConnectionString = "server=42.117.5.86;uid=ft;pwd=ftwriter;database=FT"
        Try
            EmailConn.Open()
            EmailCmd.CommandText = "Insert EmailLog (CustID, Subj, MSG, Frm, City, Dept) values (@CustID, @Subj, @MSG, @Frm, @City, @Dept)"
            EmailCmd.Parameters.Clear()
            EmailCmd.Parameters.Add("@CustID", SqlDbType.Int).Value = pCustID
            EmailCmd.Parameters.Add("@Subj", SqlDbType.VarChar).Value = pSubj
            EmailCmd.Parameters.Add("@MSG", SqlDbType.VarChar).Value = pMsg
            EmailCmd.Parameters.Add("@Frm", SqlDbType.VarChar).Value = pPRG
            EmailCmd.Parameters.Add("@City", SqlDbType.VarChar).Value = pPOS
            EmailCmd.Parameters.Add("@Dept", SqlDbType.VarChar).Value = "TKT"
            EmailCmd.ExecuteNonQuery()
            EmailConn.Close()
        Catch ex As Exception
            Return "Err Connecting to Mail Service"
        End Try
        Return "OK"
    End Function
    Public Shared Function Report_IamRunning(ByVal pPRG As String) As String
        Dim RPTConn As New SqlClient.SqlConnection
        Dim RPTCmd As SqlClient.SqlCommand = RPTConn.CreateCommand
        RPTConn.ConnectionString = "server=42.117.5.86;uid=ft;pwd=ftwriter;database=FT"
        Try
            RPTConn.Open()
            RPTCmd.CommandText = String.Format("update fox_MISC set Dat1=Getdate() where cat='TTGT' and VAL='{0}'", pPRG)
            RPTCmd.ExecuteNonQuery()
            RPTConn.Close()
        Catch ex As Exception
            Return "Err Connecting to Mail Service"
        End Try
        Return "OK"
    End Function
End Class
