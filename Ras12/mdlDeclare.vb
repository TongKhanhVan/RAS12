Module mdlDeclare
    Public myStaff As New objStaff
    Public MySession As New objTerminal
    Public MyAL As New objAL

    Public Conn As New SqlClient.SqlConnection
    Public Conn_Web As New SqlClient.SqlConnection
    Public Conn_TVVN As New SqlClient.SqlConnection
    Public CnStr As String = ""
    Public Const CnStr_CGO As String = "server=42.117.5.86;uid=cosuser;pwd=Healthy@FoodC14;database=COS" ' Update Ti gia cho COS
    Public Const CnStr_TVW As String = "server=42.117.5.86;uid=ft;pwd=Healthy@Food;database=FT" ' chck Balance
    Public Const CnStr_TVVN As String = "server=42.117.5.86;uid=transviet;pwd=Healthy@Food;database=transvietvn" ' Mapp User webReport
    Public Const CnStr_F1S As String = "SERVER=42.117.5.86;uid=f1s;pwd=front1s;database=f1s" ' tao User cho F1S
    Public Const CnStr_FLX As String = "Data Source=42.117.5.86;Initial Catalog=FLX;UID=flxusers;Pwd=Healthy@Food;TimeOut=30" ' Theo doi Qtuan

    Public Const msgTitle As String = "TransViet Travel :: RAS"
    Public pubVarSRV As String = ""
    Public pubVarRCPID_BeingEdited As Integer = 0
    Public pubVarRCPID_BeingCreated As Integer = 0
    Public pubVarBackColor As Color

    Public CutOverDatePPD As Date
    Public CutOverDatePSP As Date
    Public CutOverDateCloseRPT As Date

    Public DDAN As String
    Public pstrPrg As String = "RAS"

    Public pstrVnDomCities As String

    Public Const DKDataConvertMktg_RAS As String = " from TKT t inner join rcp r on t.rcpid=r.recid and r.status not in ('XX','QQ','NA') " & _
        "and t.RecID not in (select TKID from ReportData)  and t.al not in ('XX','01') and t.status<>'XX' and t.srv <>'V' " & _
        "and doctype not in ('GRP','SST') and doi >'19-may-2013'   and (sbu='TVS' or fare+tax+t.charge <>0 )" _
        & " and t.RecID not in (select TKID from TVSGrpOnNH())"
    Public Const DKDataConvertMktg_BSP As String = " from mktg_MIDT.dbo.UA_Hot where ID not in (select TKID from ReportData_BSP) " & _
        "and tdnr <>'' and dais<>''"
End Module
