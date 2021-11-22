Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions

Public Class frmVatInvoicePrint4Cwt
    Private mblnDataLoadCompleted As Boolean
    Private mobjExcel As New Excel.Application
    Private mobjWbk As Workbook
    Private mobjWsh As Worksheet

    Private Sub frmVatInvoicePrint4Cwt_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        mobjExcel.Quit()
    End Sub

    Private Sub frmVatInvoicePrint4Cwt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cboCustomer.SelectedIndex = 0
    End Sub

    Private Sub lbkLoadFile_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkLoadFile.LinkClicked
        Dim objOfd As New OpenFileDialog
        With objOfd
            .Filter = "excel files (*.xls)|"
            .ShowDialog()
            If .FileName = "" Then
                Exit Sub
            End If

            Dim objWsh As Worksheet
            'Dim i As Integer, j As Integer

            mobjWbk = mobjExcel.Workbooks.Open(.FileName, , True)
            mblnDataLoadCompleted = False

            Select Case cboCustomer.Text
                Case "EY HAN", "EY SGN"
                    For Each objWsh In mobjWbk.Sheets
                        If objWsh.Name.StartsWith("DOM") _
                            Or objWsh.Name.StartsWith("INT") Then
                            cboSheetName.Items.Add(objWsh.Name)
                        End If
                    Next
                    mblnDataLoadCompleted = True
                    cboSheetName.SelectedIndex = 0
                Case "ORACLE"
            End Select

        End With
    End Sub
    Private Function LoadWorksheet2Datagridview(objWsh As Worksheet) As Boolean
        Dim intMaxColumn As Integer
        Dim i As Integer, j As Integer

        dgrTktListing.Rows.Clear()
        dgrTktListing.Columns.Clear()
        dgrTktListing.Columns.Add("SheetName", "SheetName")
        For i = 1 To 54
            If objWsh.Cells(6, i) Is Nothing Or objWsh.Cells(6, i).value = "" Then
                intMaxColumn = i
                Exit For
            Else
                dgrTktListing.Columns.Add(objWsh.Cells(6, i).value, objWsh.Cells(6, i).value)
            End If
        Next

        For i = 7 To 1000
            If objWsh.Range("A" & i).Value Is Nothing Then
                Exit For
            End If
            dgrTktListing.Rows.Add()
            dgrTktListing.Rows(dgrTktListing.Rows.Count - 1).Cells(0).Value = objWsh.Name
            For j = 1 To intMaxColumn - 1
                dgrTktListing.Rows(dgrTktListing.Rows.Count - 1).Cells(j).Value _
                    = objWsh.Cells(i, j).value
            Next
        Next

        Return True
    End Function
    Private Function CheckCorpName(ByRef objWsh As Worksheet, strCustomerName As String _
                                   , strKeyword As String) As Boolean
        Select Case strCustomerName
            Case "EY"
                If Not objWsh.Range("C2").Value.ToString.Contains("ERNST & YOUNG") Then
                    Return False
                End If
            Case "ORACLE"
        End Select

        Return True
    End Function

    Private Sub lbkPrint_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lbkPrint.LinkClicked
        If dgrTktListing.CurrentRow Is Nothing Then
            Exit Sub
        End If
        Select Case cboCustomer.Text
            Case "EY HAN", "EY SGN"
                PrintVatInvoice(My.Application.Info.DirectoryPath & "\VAT Invoice " _
                                & cboCustomer.Text & ".xls")
        End Select


    End Sub
    Private Function PrintVatInvoice(strFileName As String) As Boolean
        Dim objExcel As New Excel.Application
        Dim objWbk As Workbook
        Dim objWsh As Worksheet
        Dim i As Integer, j As Integer
        objWbk = objExcel.Workbooks.Open(strFileName, , True)
        objWsh = objWbk.Sheets("Data")
        'objExcel.Visible = True
        'objExcel.ActiveWindow.Activate()
        With objWsh
            .Range("A2").Value = dtpInvoiceDate.Value.Date
            .Range("F4:H9").ClearContents()

            If dgrTktListing.CurrentRow.Cells("SRV").Value = "S" Then
                .Range("B4").Value = objWbk.Sheets("Temp").Range("A1").value
            ElseIf dgrTktListing.CurrentRow.Cells("SRV").Value = "R" Then
                .Range("B4").Value = objWbk.Sheets("Temp").Range("A2").value
            End If
            .Range("B5").Value = dgrTktListing.CurrentRow.Cells("TKNO").Value
            .Range("B6").Value = ConvertItinerary4EY(dgrTktListing.CurrentRow.Cells("Itinerary").Value)
            .Range("B7").Value = dgrTktListing.CurrentRow.Cells("Pax Name").Value

            If cboSheetName.Text.StartsWith("DOM") Or cboSheetName.Text.StartsWith("MISC") Then
                .Range("F4").Value = dgrTktListing.CurrentRow.Cells("Net Fare").Value
                .Range("H4").Value = dgrTktListing.CurrentRow.Cells("VAT4FARE").Value
                .Range("F8").Value = Math.Round(dgrTktListing.CurrentRow.Cells("AL Charge").Value)
                .Range("H8").Value = Math.Round(dgrTktListing.CurrentRow.Cells("VAT AL Charge").Value)
                .Range("F9").Value = Math.Round(dgrTktListing.CurrentRow.Cells("Tax").Value)
                .Range("H9").Value = Math.Round(dgrTktListing.CurrentRow.Cells("VAT Tax").Value)
            Else
                Dim decRoe As Decimal
                If dgrTktListing.CurrentRow.Cells("Currency").Value = "VND" Then
                    decRoe = 1
                Else
                    decRoe = dgrTktListing.CurrentRow.Cells("ROE (USD/VND)").Value
                End If
                .Range("F4").Value = Math.Round(dgrTktListing.CurrentRow.Cells("Net Fare").Value _
                    * decRoe)
                .Range("F8").Value = Math.Round(dgrTktListing.CurrentRow.Cells("AL Charge").Value _
                    * decRoe)
                .Range("F9").Value = Math.Round(dgrTktListing.CurrentRow.Cells("Tax").Value _
                    * decRoe)

            End If
            .Range("F11").Value = Math.Round(dgrTktListing.CurrentRow.Cells("Service Fee (No VAT)").Value)
            .Range("H11").Value = Math.Round(dgrTktListing.CurrentRow.Cells("VAT (SF)").Value)

            If cboSheetName.Text.StartsWith("INT") Then
                .Range("G4").Value = 0
            ElseIf .Range("H4").Value <> 0 AndAlso cboSheetName.Text.StartsWith("DOM") Then
                .Range("G4").Value = 10
            Else
                .Range("G4").Value = ""
            End If
            If cboSheetName.Text.StartsWith("INT") Then

            ElseIf .Range("H8").Value <> 0 Then
                .Range("G8").Value = 10
            Else
                .Range("G8").Value = ""
            End If
            If cboSheetName.Text.StartsWith("INT") Then
                .Range("G9").Value = 0
            ElseIf .Range("H9").Value <> 0 Then
                .Range("G9").Value = 10
            Else
                .Range("G9").Value = ""
            End If

            If .Range("H11").Value <> 0 Then
                .Range("G11").Value = 10
            Else
                .Range("G11").Value = ""
            End If
            .Activate()

        End With
        objExcel.Visible = True
        'objExcel.ActiveWindow.Activate()
        Return True
    End Function

    Private Sub cboSheetName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSheetName.SelectedIndexChanged
        If mblnDataLoadCompleted Then
            For Each objWsh As Worksheet In mobjWbk.Sheets
                If objWsh.Name = cboSheetName.Text Then
                    LoadWorksheet2Datagridview(objWsh)
                End If
            Next
        End If
    End Sub
    Private Function ConvertItinerary4EY(strOldItinerary As String) As String
        Dim rgCar As New Regex("\s\w\w\s")
        Dim arrCities As String()
        Dim i As Integer

        strOldItinerary = rgCar.Replace(strOldItinerary, "-")
        strOldItinerary = Replace(strOldItinerary, " // ", "--")
        arrCities = strOldItinerary.Split("-")
        For i = 0 To arrCities.Length - 1
            Select Case arrCities(i)
                Case "SGN"
                    arrCities(i) = "HO CHI MINH"
                Case "HAN"
                    arrCities(i) = "HA NOI"
                Case "DAD"
                    arrCities(i) = "DA NANG"
                Case Else
                    arrCities(i) = ScalarToString("CityCode", "CityName" _
                                               , "Airport='" & arrCities(i) & "'")
            End Select
        Next
        Return Replace(Join(arrCities, "-"), "--", "//")
    End Function

    Private Sub cboCustomer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCustomer.SelectedIndexChanged
        If mblnDataLoadCompleted Then
            cboSheetName.Items.Clear()
        End If

    End Sub
End Class