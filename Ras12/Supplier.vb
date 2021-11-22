Public Class Supplier
    Private Sub OptDontKnowAccountNo_CheckedChanged(sender As Object, e As EventArgs) Handles OptDontKnowAccountNo.CheckedChanged
        Me.GroupBox1.Visible = Me.OptDontKnowAccountNo.Checked
        Me.GroupBox2.Visible = Me.OptDontKnowAccountNo.Checked
        Me.OptCreateNew.Checked = Me.OptDontKnowAccountNo.Checked
        Me.GroupBox2.Enabled = False
    End Sub
    Private Sub LblSearch_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSearch.LinkClicked
        Me.GroupBox2.Visible = True
        Me.OptNewSearch.Checked = False
        Me.OptCreatedNewOfSelectedVendor.Checked = False
        Me.OptCreateNew.Checked = False
        Me.OptCreatedNewOfSelectedVendor.Enabled = False
        Me.GroupBox2.Enabled = True
        Me.GridSupplier.DataSource = GetDataTable("Select VendorID, FullName, Address from Supplier where status='OK' and " & _
            "VendorID in (select CompanyID from UNC_Accounts where replace(accountNumber,' ','') like '%" & _
            Me.TxtAccountNo.Text.Replace(" ", "") & "' and status <>'XX')")
        Me.GridSupplier.Columns(0).Visible = False
        Me.GridSupplier.Columns(1).Width = 128
        Me.GridSupplier.Columns(2).Width = 256

    End Sub
    Private Sub GridSupplier_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridSupplier.CellContentClick
        Me.OptCreatedNewOfSelectedVendor.Enabled = True
    End Sub


    Private Sub OptCreateNew_Click(sender As Object, e As EventArgs) Handles OptCreateNew.Click, OptCreatedNewOfSelectedVendor.Click
        Me.GroupBox1.Visible = True
    End Sub

    Private Sub OptNewSearch_Click(sender As Object, e As EventArgs) Handles OptNewSearch.Click
        Me.GroupBox1.Visible = False
    End Sub

    Private Sub Supplier_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()

    End Sub

    Private Sub Supplier_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCmb_VAL(CmbCountry, "select Country as VAL, CountryName as DIS from lib.dbo.Country order by DIS")
        LoadCmb_VAL(CmbNewCountry, "select Country as VAL, CountryName as DIS from lib.dbo.Country order by DIS")
        Me.CmbCountry.Text = "VIETNAM"
        Me.CmbNewCountry.Text = "VIETNAM"
        loadGridSupplier_Edit()
    End Sub

    Private Sub CmbCountry_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCountry.SelectedIndexChanged
        If Me.CmbCountry.SelectedValue = "VN" Then
            Me.CmbCity.DropDownStyle = ComboBoxStyle.DropDownList
        Else
            Me.CmbCity.DropDownStyle = ComboBoxStyle.DropDown
            Me.CmbCity.Text = ""
        End If
    End Sub
    Private Sub LblAdd_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblAdd.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        Dim vVendorID As Integer = 0
        If Me.TxtName.Text = "" Or Me.TxtAddress.Text = "" Or Me.CmbCity.Text = "" Then Exit Sub
        If Me.txtEmail.Text <> "" And Not Me.txtEmail.Text.Contains("@") Then Exit Sub
        If Me.txtPhone.Text <> "" And Me.txtPhone.Text.Length < 8 Then Exit Sub
        If Me.OptCreatedNewOfSelectedVendor.Checked Then
            vVendorID = Me.GridSupplier.CurrentRow.Cells("VendorID").Value
        End If
        cmd.CommandText = "insert Supplier (FullName, Address, Address_CityCode, Address_CountryCode, Tel, Email, VendorID, FstUser, City)" & _
            " values (@FullName, @Address, @Address_CityCode, @Address_CountryCode, @Tel, @Email, @VendorID,@FstUser, @City)"
        cmd.Parameters.Add("@FullName", SqlDbType.NVarChar).Value = Me.TxtName.Text
        cmd.Parameters.Add("@Address", SqlDbType.NVarChar).Value = Me.TxtAddress.Text
        cmd.Parameters.Add("@Address_CityCode", SqlDbType.VarChar).Value = Me.CmbCity.Text
        cmd.Parameters.Add("@Address_CountryCode", SqlDbType.VarChar).Value = Me.CmbCountry.SelectedValue
        cmd.Parameters.Add("@Tel", SqlDbType.VarChar).Value = Me.txtPhone.Text
        cmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = Me.txtEmail.Text
        cmd.Parameters.Add("@FstUser", SqlDbType.VarChar).Value = myStaff.SICode
        cmd.Parameters.Add("@City", SqlDbType.VarChar).Value = myStaff.City
        cmd.Parameters.Add("@VendorID", SqlDbType.Int).Value = vVendorID
        Try
            cmd.ExecuteNonQuery()
            Me.TxtName.Text = ""
            Me.TxtAddress.Text = ""
            Me.txtPhone.Text = ""
            Me.txtEmail.Text = ""
            Me.CmbCountry.SelectedValue = "VN"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub LoadGridSupplier_Edit()
        Me.LblSave.Visible = False
        Me.GridSupplier_Edit.DataSource = GetDataTable("select RecID, FullName, Tel, Email, Address, Address_CountryCode as Country from Supplier where status <>'XX'")
        Me.GridSupplier_Edit.Columns(0).Visible = False
        Me.GridSupplier_Edit.Columns(1).Width = 128
        Me.GridSupplier_Edit.Columns(3).Width = 256
        Me.GridSupplier_Edit.Columns(4).Width = 256
    End Sub

    Private Sub GridSupplier_Edit_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridSupplier_Edit.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        Me.txtNewName.Text = Me.GridSupplier_Edit.CurrentRow.Cells("FullName").Value
        Me.TxtNewPhone.Text = Me.GridSupplier_Edit.CurrentRow.Cells("Tel").Value
        Me.txtNewMail.Text = Me.GridSupplier_Edit.CurrentRow.Cells("eMail").Value
        Me.txtNewAddress.Text = Me.GridSupplier_Edit.CurrentRow.Cells("Address").Value
        Me.CmbNewCountry.SelectedValue = Me.GridSupplier_Edit.CurrentRow.Cells("Country").Value
        Me.LblSave.Visible = True
    End Sub

    Private Sub LblSave_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LblSave.LinkClicked
        Dim cmd As SqlClient.SqlCommand = Conn.CreateCommand
        If Me.txtNewName.Text = "" Or Me.txtNewAddress.Text = "" Then Exit Sub
        If Me.txtNewMail.Text <> "" And Not Me.txtNewMail.Text.Contains("@") Then Exit Sub
        If Me.TxtNewPhone.Text <> "" And Me.TxtNewPhone.Text.Length < 8 Then Exit Sub

        cmd.CommandText = "Update Supplier set FullName=@FullName, Address=@Address, Tel=@Tel, Email=@Email, " & _
            " Address_CountryCode=@CountryCode where RecID=@RecID"
        cmd.Parameters.Add("@FullName", SqlDbType.NVarChar).Value = Me.txtNewName.Text
        cmd.Parameters.Add("@Address", SqlDbType.NVarChar).Value = Me.txtNewAddress.Text
        cmd.Parameters.Add("@Tel", SqlDbType.VarChar).Value = Me.TxtNewPhone.Text
        cmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = Me.txtNewMail.Text
        cmd.Parameters.Add("@CountryCode", SqlDbType.VarChar).Value = Me.CmbNewCountry.SelectedValue
        cmd.Parameters.Add("@RecID", SqlDbType.Int).Value = Me.GridSupplier_Edit.CurrentRow.Cells("RecID").Value
        cmd.ExecuteNonQuery()
        Me.txtNewName.Text = ""
        Me.txtNewAddress.Text = ""
        Me.TxtNewPhone.Text = ""
        Me.txtNewMail.Text = ""
        LoadGridSupplier_Edit()
    End Sub

    Private Sub GroupBox1_VisibleChanged(sender As Object, e As EventArgs) Handles GroupBox1.VisibleChanged
        Me.TxtName.Text = ""
        Me.TxtAddress.Text = ""
        Me.txtPhone.Text = ""
        Me.txtEmail.Text = ""
        Me.CmbCountry.SelectedValue = "VN"
    End Sub
End Class