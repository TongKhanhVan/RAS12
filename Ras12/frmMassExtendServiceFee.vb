Public Class frmMassExtendServiceFee
    Private mstrSbu As String
    Private mstrTable As String
    
    Public Sub New(strSbu As String, strTable As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mstrSbu = strSbu
        mstrTable = strTable
    End Sub
    Private Sub frmMassExtendServiceFee_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Search()
        If Now.Month > 9 Then
            dtpNewValidThru.Value = DateSerial(Now.Year + 1, 12, 31)
        Else
            dtpNewValidThru.Value = DateSerial(Now.Year, 12, 31)
        End If
    End Sub

    Private Function Search() As Boolean
        Dim strSql As String

        strSql = "Select distinct AL, Channel, CustLevel, ValidThru from " & mstrTable _
            & " where sbu='" & mstrSbu _
            & "' and status='OK' and ValidThru between getdate() and '" & Now.AddMonths(3) & "'"
        Me.GridExpireDate.DataSource = GetDataTable(strSQL)
        'Me.GridExpireDate.Left = 597
        'Me.GridExpireDate.Height = 256
        'Me.GridExpireDate.Columns(0).Width = 32
        'Me.GridExpireDate.Columns(1).Width = 32
        'Me.GridExpireDate.Columns(2).Width = 32
        LoadDataGridView(GridExpireDate, strSql, Conn)
        If GridExpireDate.RowCount = 0 Then
            lbkExtend.Visible = False
        End If
        Return True
    End Function

    Private Sub lblExtend_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkExtend.LinkClicked
        Dim lstQuerries As New List(Of String)
        Dim strFilter As String

        With GridExpireDate.CurrentRow
            strFilter = " where Status='OK' and SBU='" & mstrSbu _
                        & "' and AL='" & .Cells("AL").Value _
                        & "' and Channel='" & .Cells("Channel").Value _
                        & "' and CustLevel='" & .Cells("CustLevel").Value _
                        & "' and ValidThru='" & .Cells("ValidThru").Value & "'"
        End With
        If mstrTable = "Cust_Discount" Then
            lstQuerries.Add("insert into " & mstrTable _
                    & "(SBU, AL, COC, Channel, CustLevel, DType, VAL, ValidFrom, Validthru" _
                    & ", FareSource, Base,FstUser,Status, ApproveBy)" _
                    & " Select SBU, AL, COC, Channel, CustLevel, DType, VAL, ValidFrom,'" _
                    & CreateToDate(dtpNewValidThru.Value) _
                    & "', FareSource, Base,'" & myStaff.SICode & "','OK','EXT' from " _
                    & mstrTable & strFilter)
        Else
            lstQuerries.Add("insert into " & mstrTable _
                    & "(ValidFrom, ValidThru, SBU, AL, Channel, CustLevel, TRXType, Base" _
                    & ", G_FIT, FID_fier, ISI, FareFrom, FareThrough, SFType, VAL, MinVal, MaxVal" _
                    & ",Status, FstUser, ApproveBy, Refundable, Area, INF, DestCountries)" _
                    & " Select ValidFrom,'" & CreateToDate(dtpNewValidThru.Value) _
                    & "', SBU, AL, Channel, CustLevel, TRXType, Base" _
                    & ", G_FIT, FID_fier, ISI, FareFrom, FareThrough, SFType, VAL, MinVal, MaxVal" _
                    & ",'OK','" & myStaff.SICode _
                    & "','EXT', Refundable, Area, INF, DestCountries from " _
                    & mstrTable & strFilter)
        End If
        lstQuerries.Add("Update " & mstrTable & " set Status='XX', LstUpdate=Getdate(),LstUser='" & myStaff.SICode _
                        & "'" & strFilter)

        If Not UpdateListOfQuerries(lstQuerries, Conn) Then
            MsgBox("Unable to extend validity!")
        End If
        Search()


    End Sub
End Class