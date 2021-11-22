Module MdlScalar
    Private Cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Public Function ScalarToInt(ByVal pTbl As String, ByVal pField As String, ByVal pDK_Order As String) As Integer
        Dim KQ As Integer
        Cmd.CommandText = "SELECT " & pField & " from " & pTbl & " where " & Finetune_pDK(pDK_Order)
        KQ = Cmd.ExecuteScalar
        Return KQ
    End Function
    Public Function ScalarToDec(ByVal pTbl As String, ByVal pField As String, ByVal pDK_Order As String) As Decimal
        Dim KQ As Decimal
        Cmd.CommandText = "SELECT " & pField & " from " & pTbl & " where " & Finetune_pDK(pDK_Order)
        KQ = Cmd.ExecuteScalar
        Return KQ
    End Function

    Public Function ScalarToString(ByVal pTbl As String, ByVal pField As String, ByVal pDK_Order As String) As String
        Dim KQ As String
        Cmd.CommandText = "SELECT " & pField & " from " & pTbl & " where " & Finetune_pDK(pDK_Order)
        KQ = Cmd.ExecuteScalar
        Return KQ
    End Function

    Public Function ScalarToDate(ByVal pTbl As String, ByVal pField As String, ByVal pDK_Order As String) As Date
        Dim KQ As Date
        Cmd.CommandText = "SELECT " & pField & " from " & pTbl & " where " & Finetune_pDK(pDK_Order)
        KQ = Cmd.ExecuteScalar
        Return KQ
    End Function
    Public Function ChangeStatus_ByID(ByVal pTable As String, ByVal pStatus As String, ByVal PID As Integer, Optional ByVal pALStatus As String = "") As String
        Dim KQ As String
        KQ = "update " & pTable & " set LstUser='" & myStaff.SICode & "', Lstupdate=cast(getdate() as smalldatetime), status='" & pStatus & "'"
        If pALStatus <> "" Then
            KQ = KQ & ", StatusAL='" & pALStatus & "'"
        End If
        KQ = KQ & " where recID = " & PID
        Return KQ
    End Function
    Public Function ChangeStatus_ByDK(ByVal pTable As String, ByVal pStatus As String, ByVal pDK As String, Optional ByVal pALStatus As String = "") As String
        Dim KQ As String
        KQ = "update " & pTable & " set LstUser='" & myStaff.SICode & "', Lstupdate=getdate(), status='" & pStatus & "'"
        If pALStatus <> "" Then
            KQ = KQ & ", StatusAL='" & pALStatus & "'"
        End If
        KQ = KQ & " where " & Finetune_pDK(pDK)
        Return KQ
    End Function
    Private Function Finetune_pDK(ByVal pDK As String) As String
        If pDK.Trim.Substring(0, 4).ToUpper = "WHER" Then
            Return pDK.Trim.Substring(5)
        End If
        Return pDK
    End Function
    Public Function GetColumnValuesAsString(strTblName As String, strColumn As String, strCondition As String _
                                       , strSeperator As String) As String
        Dim strResult As String
        Cmd.CommandText = "SELECT " & strColumn & "+'" & strSeperator & "' from " _
            & strTblName & " " & strCondition & " for xml path('')"
        strResult = Cmd.ExecuteScalar
        If strResult <> "" Then
            strResult = Mid(strResult, 1, strResult.Length - strSeperator.Length)
        End If
        Return strResult
    End Function
End Module
