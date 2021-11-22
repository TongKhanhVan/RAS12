Imports System.Text.RegularExpressions
Public Class frmSelectCcNbr
    Private mstrCustShortName As String
    Private mstrPaxName As String
    Private mintCcId As Integer
    Private mstrCounter As String
    Private mblnFirstLoad As Boolean
    Public Sub New(dgrFopRcp As DataGridViewRow)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim rgPaxTitle As New Regex("\sMR$|\sMSTR$|\sMS$|\sMISS$|\sMIST$")
        mstrCustShortName = dgrFopRcp.Cells("CustShortName").Value
        mintCcId = dgrFopRcp.Cells("Ccid").Value
        mstrCounter = dgrFopRcp.Cells("Counter").Value

        If dgrFopRcp.Cells("CcId").Value = 0 Then
            Select Case dgrFopRcp.Cells("Counter").Value
                Case "CWT"
                    mstrPaxName = ScalarToString("tkt", "top 1 PaxName", "Status<>'XX' and RcpNo='" & dgrFopRcp.Cells("RcpNo").Value & "'")
                Case "N-A"
                    mstrPaxName = ScalarToString("dutoan_tour", "top 1 Brief", "Status<>'XX' and Tcode='" & dgrFopRcp.Cells("Document").Value & "'")
                    If mstrPaxName.Contains(" FOR ") Then
                        mstrPaxName = Split(mstrPaxName, " FOR ")(1)
                    End If
            End Select
            mstrPaxName = rgPaxTitle.Replace(mstrPaxName, "")
            mstrPaxName = mstrPaxName.Replace("/", " ")
        End If
        mblnFirstLoad = True
    End Sub
    Private Sub frmSelectCcNbr_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Search()
    End Sub
    Private Function Search() As Boolean
        Dim strQuerry As String
        Dim arrNameBreaks() As String
        Dim i As Integer

        If Not mblnFirstLoad Then Return False



        Dim strCcNbr As String

        If chkHideCcNbr.Checked Then
            strCcNbr = ",m.Details as Last4Digit"
        Else
            strCcNbr = ",(m2.Val1+m.Details) as CardNbr"
        End If

        strQuerry = "Select m.RecId, m.Val as CustShortName,m.Val1 as CardHolder,m.Val2 as CardType" _
                    & strCcNbr & ",m.Description as ExpDate" _
                    & " ,m2.Val2 as Biz,m2.Details as Remark" _
                    & " from  Misc m" _
                    & " LEFT JOIN  Misc m2 on m.RecId =m2.Val" _
                    & " where m.Cat='CcNbr' and m.Status='OK' and m.Val='" & mstrCustShortName _
                    & "' and m2.Cat='PrefixC'"

        If Not chkShowAll.Checked Then
            If mintCcId > 0 Then
                strQuerry = strQuerry & " and m.RecId=" & mintCcId
            ElseIf mstrPaxName <> "" Then
                arrNameBreaks = mstrPaxName.Split(" ")
                For i = 0 To arrNameBreaks.Length - 1
                    arrNameBreaks(i) = " and m.Val1 like '%" & arrNameBreaks(i) & "%'"
                Next
                strQuerry = strQuerry & Join(arrNameBreaks, " ")
                
            End If
        End If

        strQuerry = strQuerry & " order by m.Val1"
        LoadDataGridView(dgCcList, strQuerry, Conn)

        Return True
    End Function

    Private Sub chkHideCcNbr_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideCcNbr.CheckedChanged
        Search()
    End Sub

    Private Sub chkShowAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAll.CheckedChanged
        Search()
    End Sub

    Private Sub lbkSelect_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkSelect.LinkClicked
        With dgCcList
            If .CurrentRow Is Nothing Then
                MsgBox("You must Select a CreditCardNbr")
                Exit Sub
            Else
                mintCcId = .CurrentRow.Cells("RecId").Value
                DialogResult = Windows.Forms.DialogResult.OK
                Me.Dispose()
            End If
        End With

    End Sub
    Public Property CcId As Integer
        Get
            Return mintCcId
        End Get
        Set(value As Integer)
            mintCcId = value
        End Set
    End Property
End Class