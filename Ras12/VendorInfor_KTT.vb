Public Class VendorInfor_KTT
    Private WhoIs As String
    Private Cmd As SqlClient.SqlCommand = Conn.CreateCommand
    Sub New(pWho As String)
        WhoIs = pWho
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Private Sub VendorInfor_KTT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CmbCAT.Text = Me.CmbCAT.Items(0).ToString
        If WhoIs = "KTT" Then
            Me.LblChangeFOP.Enabled = False
            Me.LblChangeCat.Enabled = False
            Me.LblChangeName.Enabled = False
            Me.LblDeleteCompany.Enabled = False
            Me.Panel1.Enabled = False
            Me.CmbCAT.Text = "NH"
            Me.TabControl1.SelectTab("TabPage2")
        Else
            LoadCmb_VAL(CmbCountry, "select Country as VAL, CountryName as DIS from lib.dbo.Country order by DIS")
            Me.CmbCountry.Text = "VIETNAM"
            Me.CmbNewCat.Text = Me.CmbNewCat.Items(0).ToString
            Me.CmbFOPedit.Text = "PPD"
            Me.CmbFOP1New.Text = "BTF"
        End If
    End Sub
    Private Sub LblFilter_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblFilter.LinkClicked
        LoadGridVendor()
    End Sub

    Private Sub GridVendor_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridVendor.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If Me.ChckCompanyOKonly.Checked Then
            Me.LblDeleteCompany.Visible = True
            Me.LblChangeCat.Visible = True
            Me.LblChangeName.Visible = True
            Me.LblUpdateHD.Visible = True
            Me.LblChangeFOP.Visible = True
        End If
        Me.TxtNewShortName.Text = Me.GridVendor.CurrentRow.Cells("ShortName").Value
    End Sub

    Private Sub LblUpdate_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblUpdateHD.LinkClicked
        ChangeInfor("HD", Me.CmbINV.Text)
    End Sub
    Private Sub ChangeInfor(pField As String, pVal As String)
        Cmd.CommandText = "update UNC_Company set " & pField & "='" & pVal & "' where recID=" & Me.GridVendor.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        LoadGridVendor()
    End Sub
    Private Sub LoadGridVendor()
        Dim StrSQL As String = "select RecID, HD, FOP, ShortName from UNC_Company where cat='" & Me.CmbCAT.Text.Substring(0, 2) & "' and status<>'XX'"
        If Me.TxtNameFilter.Text <> "" Then StrSQL = StrSQL & " and shortname like '%" & Me.TxtNameFilter.Text & "%'"
        Me.GridVendor.DataSource = GetDataTable(StrSQL & " order by ShortName")
        Me.LblUpdateHD.Visible = False
        Me.LblChangeFOP.Visible = False

        Me.GridVendor.Columns(0).Visible = False
        Me.GridVendor.Columns(1).Width = 32
        Me.GridVendor.Columns(2).Width = 32
        Me.GridVendor.Columns(3).Width = 256
    End Sub

    Private Sub LblChangeFOP_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblChangeFOP.LinkClicked
        ChangeInfor("FOP", Me.CmbFOPedit.Text)
        ChangeInfor("FOP1", Me.CmbFOP1edit.Text)
    End Sub

    Private Sub LblChangeCat_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblChangeCat.LinkClicked
        Cmd.CommandText = "update UNC_Company set cat='" & Me.CmbNewCat.Text.Substring(0, 2) & _
            "' where recid=" & Me.GridVendor.CurrentRow.Cells("RecID").Value
        Cmd.ExecuteNonQuery()
        LoadGridVendor()
    End Sub

    Private Sub LblChangeName_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblChangeName.LinkClicked
        Cmd.CommandText = "update UNC_Company set ShortName='" & Me.TxtNewShortName.Text & _
            "' where recid=" & Me.GridVendor.CurrentRow.Cells("RecID").Value
        Cmd.ExecuteNonQuery()
        Me.LblChangeName.Visible = False
    End Sub

    Private Sub LblDeleteCompany_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblDeleteCompany.LinkClicked
        Dim tmpChkRecID As Integer
        Cmd.Parameters.Clear()
        cmd.CommandText = "select City from unc_company where RecID=@VendorID"
        cmd.Parameters.AddWithValue("@VendorID", Me.GridVendor.CurrentRow.Cells("recid").Value)
        If cmd.ExecuteScalar <> myStaff.City Then Exit Sub

        cmd.CommandText = "select top 1 RecID from dutoan_Item where vendorID=@VendorID" & _
                " and DutoanID in (select recid from dutoan_Tour where status not in ('XX','RR'))"
        tmpChkRecID = cmd.ExecuteScalar
        If tmpChkRecID > 0 Then
            MsgBox("There is Pending NonAir Booking on This Vendor", MsgBoxStyle.Critical, msgTitle)
            Exit Sub
        End If

        Cmd.CommandText = ChangeStatus_ByID("UNC_Company", "XX", Me.GridVendor.CurrentRow.Cells("recid").Value) & _
            ";" & ChangeStatus_ByDK("UNC_Accounts", "XX", "status ='OK' and CompanyID=" & Me.GridVendor.CurrentRow.Cells("recid").Value)
        Cmd.CommandText = Cmd.CommandText & "; Update Lib.dbo.Supplier set VendorID=0 where VendorID=@VendorID"
        cmd.ExecuteNonQuery()
        LoadGridVendor()
    End Sub
    Private Function AddVendor() As Integer
        Dim KQ As Integer = 0
        If Me.TxtShortName.Text = "" Then Exit Function
        Cmd.CommandText = "insert into UNC_Company (cat, shortname, FOP, fstuser, city, FOP1) values ('" & _
            Me.CmbCAT.Text.Substring(0, 2) & "','" & Me.TxtShortName.Text.Replace("--", "") & "','" & _
            Me.CmbFOPNew.Text & "','" & myStaff.SICode & "','" & myStaff.City & "','" & Me.CmbFOP1New.Text & "') "
        cmd.ExecuteNonQuery()
        cmd.CommandText = "select top 1 RecID from uNC_Company order by recID Desc"
        KQ = cmd.ExecuteScalar
        LoadGridVendor()

        Me.TxtShortName.Text = ""
        Return KQ

    End Function
    Private Sub LblAddNew_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAddNewVendor.LinkClicked
        If AddVendor() = 0 Then
            MsgBox("Invalid Input:" & Err.Description, MsgBoxStyle.Critical, msgTitle)
        End If
    End Sub
    Private Sub LblCreateVendorAndSupplier_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblCreateVendorAndSupplier.LinkClicked
        Dim NewVendorID As Integer = AddVendor()
        If NewVendorID = 0 Then Exit Sub
        If Me.TxtSupplierName.Text = "" Or Me.TxtAddress.Text = "" Or Me.CmbCity.Text = "" Then Exit Sub
        If Me.txtEmail.Text <> "" And Not Me.txtEmail.Text.Contains("@") Then Exit Sub
        If Me.txtPhone.Text <> "" And Me.txtPhone.Text.Length < 8 Then Exit Sub
        Cmd.CommandText = "insert Supplier (FullName, Address, Address_CityCode, Address_CountryCode, Tel, Email, VendorID, FstUser, City)" & _
            " values (@FullName, @Address, @Address_CityCode, @Address_CountryCode, @Tel, @Email, @VendorID, @FstUser, @City)"
        Cmd.Parameters.Clear()
        cmd.Parameters.Add("@FullName", SqlDbType.NVarChar).Value = Me.TxtSupplierName.Text
        cmd.Parameters.Add("@Address", SqlDbType.NVarChar).Value = Me.TxtAddress.Text
        cmd.Parameters.Add("@Address_CityCode", SqlDbType.VarChar).Value = Me.CmbCity.Text
        cmd.Parameters.Add("@Address_CountryCode", SqlDbType.VarChar).Value = Me.CmbCountry.SelectedValue
        cmd.Parameters.Add("@Tel", SqlDbType.VarChar).Value = Me.txtPhone.Text
        cmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = Me.txtEmail.Text
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@City", SqlDbType.VarChar).Value = myStaff.City
        cmd.Parameters.Add("@VendorID", SqlDbType.Int).Value = NewVendorID
        cmd.ExecuteNonQuery()
        Me.TxtSupplierName.Text = ""
        Me.TxtAddress.Text = ""
        Me.txtPhone.Text = ""
        Me.txtEmail.Text = ""
        Me.CmbCountry.SelectedValue = "VN"
    End Sub

    Private Sub CmbCountry_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCountry.SelectedIndexChanged
        Try
            If Me.CmbCountry.SelectedValue = "VN" Then
                Me.CmbCity.DropDownStyle = ComboBoxStyle.DropDownList
            Else
                Me.CmbCity.DropDownStyle = ComboBoxStyle.DropDown
                Me.CmbCity.Text = ""
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter
        Me.TxtSupplierName.Text = Me.TxtShortName.Text
    End Sub
    Private Sub LblChangeHD_Tong_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblChangeHD_Tong.LinkClicked
        ChangeInfor("HD_TachGop", Me.CmbHD_tong.Text)
    End Sub

    Private Sub CmbCAT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCAT.SelectedIndexChanged
        LoadGridVendor()
    End Sub
End Class