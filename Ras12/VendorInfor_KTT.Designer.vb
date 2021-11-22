<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VendorInfor_KTT
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.GridVendor = New System.Windows.Forms.DataGridView()
        Me.CmbCAT = New System.Windows.Forms.ComboBox()
        Me.CmbINV = New System.Windows.Forms.ComboBox()
        Me.LblUpdateHD = New System.Windows.Forms.LinkLabel()
        Me.TxtNameFilter = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LblFilter = New System.Windows.Forms.LinkLabel()
        Me.CmbFOPedit = New System.Windows.Forms.ComboBox()
        Me.LblChangeFOP = New System.Windows.Forms.LinkLabel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.CmbFOP1New = New System.Windows.Forms.ComboBox()
        Me.TxtShortName = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.LblCreateVendorAndSupplier = New System.Windows.Forms.LinkLabel()
        Me.CmbCity = New System.Windows.Forms.ComboBox()
        Me.CmbCountry = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.TxtAddress = New System.Windows.Forms.TextBox()
        Me.TxtSupplierName = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LblAddNewVendor = New System.Windows.Forms.LinkLabel()
        Me.CmbFOPNew = New System.Windows.Forms.ComboBox()
        Me.LblAddCompany = New System.Windows.Forms.LinkLabel()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.CmbFOP1edit = New System.Windows.Forms.ComboBox()
        Me.CmbHD_tong = New System.Windows.Forms.ComboBox()
        Me.TxtNewShortName = New System.Windows.Forms.TextBox()
        Me.LblChangeHD_Tong = New System.Windows.Forms.LinkLabel()
        Me.LblChangeName = New System.Windows.Forms.LinkLabel()
        Me.CmbNewCat = New System.Windows.Forms.ComboBox()
        Me.LblDeleteCompany = New System.Windows.Forms.LinkLabel()
        Me.LblChangeCat = New System.Windows.Forms.LinkLabel()
        Me.ChckCompanyOKonly = New System.Windows.Forms.CheckBox()
        CType(Me.GridVendor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GridVendor
        '
        Me.GridVendor.AllowUserToAddRows = False
        Me.GridVendor.AllowUserToDeleteRows = False
        Me.GridVendor.BackgroundColor = System.Drawing.Color.Pink
        Me.GridVendor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridVendor.Location = New System.Drawing.Point(5, 4)
        Me.GridVendor.Name = "GridVendor"
        Me.GridVendor.RowHeadersVisible = False
        Me.GridVendor.Size = New System.Drawing.Size(346, 469)
        Me.GridVendor.TabIndex = 0
        '
        'CmbCAT
        '
        Me.CmbCAT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCAT.FormattingEnabled = True
        Me.CmbCAT.Items.AddRange(New Object() {"AR - Air", "NH - Restaurant", "KS - Hotel", "XE - Coach", "BO - Boat", "LD - Land", "OS - OtherTour Service", "TO - TO/TA", "OT - Other", "US - Utility Supplier", "AL - Airlines", "AP - Airport", "YT - So YT"})
        Me.CmbCAT.Location = New System.Drawing.Point(32, 22)
        Me.CmbCAT.Name = "CmbCAT"
        Me.CmbCAT.Size = New System.Drawing.Size(120, 21)
        Me.CmbCAT.TabIndex = 1
        '
        'CmbINV
        '
        Me.CmbINV.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbINV.FormattingEnabled = True
        Me.CmbINV.Items.AddRange(New Object() {"VAT", "B.H", "N/A"})
        Me.CmbINV.Location = New System.Drawing.Point(84, 10)
        Me.CmbINV.Name = "CmbINV"
        Me.CmbINV.Size = New System.Drawing.Size(55, 21)
        Me.CmbINV.TabIndex = 2
        '
        'LblUpdateHD
        '
        Me.LblUpdateHD.AutoSize = True
        Me.LblUpdateHD.Location = New System.Drawing.Point(6, 13)
        Me.LblUpdateHD.Name = "LblUpdateHD"
        Me.LblUpdateHD.Size = New System.Drawing.Size(72, 13)
        Me.LblUpdateHD.TabIndex = 3
        Me.LblUpdateHD.TabStop = True
        Me.LblUpdateHD.Text = "ChangeHD to"
        Me.LblUpdateHD.Visible = False
        '
        'TxtNameFilter
        '
        Me.TxtNameFilter.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNameFilter.Location = New System.Drawing.Point(160, 475)
        Me.TxtNameFilter.Name = "TxtNameFilter"
        Me.TxtNameFilter.Size = New System.Drawing.Size(156, 20)
        Me.TxtNameFilter.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(75, 478)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(79, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Name Contains"
        '
        'LblFilter
        '
        Me.LblFilter.AutoSize = True
        Me.LblFilter.Location = New System.Drawing.Point(322, 478)
        Me.LblFilter.Name = "LblFilter"
        Me.LblFilter.Size = New System.Drawing.Size(29, 13)
        Me.LblFilter.TabIndex = 6
        Me.LblFilter.TabStop = True
        Me.LblFilter.Text = "Filter"
        '
        'CmbFOPedit
        '
        Me.CmbFOPedit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbFOPedit.FormattingEnabled = True
        Me.CmbFOPedit.Items.AddRange(New Object() {"PSP", "PPD"})
        Me.CmbFOPedit.Location = New System.Drawing.Point(84, 37)
        Me.CmbFOPedit.Name = "CmbFOPedit"
        Me.CmbFOPedit.Size = New System.Drawing.Size(55, 21)
        Me.CmbFOPedit.TabIndex = 9
        '
        'LblChangeFOP
        '
        Me.LblChangeFOP.AutoSize = True
        Me.LblChangeFOP.Location = New System.Drawing.Point(6, 40)
        Me.LblChangeFOP.Name = "LblChangeFOP"
        Me.LblChangeFOP.Size = New System.Drawing.Size(77, 13)
        Me.LblChangeFOP.TabIndex = 11
        Me.LblChangeFOP.TabStop = True
        Me.LblChangeFOP.Text = "ChangeFOP to"
        Me.LblChangeFOP.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(354, 4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(426, 498)
        Me.TabControl1.TabIndex = 12
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Panel1)
        Me.TabPage1.Controls.Add(Me.LblAddCompany)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(418, 472)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Add Vendor"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.CmbFOP1New)
        Me.Panel1.Controls.Add(Me.TxtShortName)
        Me.Panel1.Controls.Add(Me.CmbCAT)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.LblAddNewVendor)
        Me.Panel1.Controls.Add(Me.CmbFOPNew)
        Me.Panel1.Location = New System.Drawing.Point(5, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(408, 454)
        Me.Panel1.TabIndex = 15
        '
        'CmbFOP1New
        '
        Me.CmbFOP1New.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbFOP1New.FormattingEnabled = True
        Me.CmbFOP1New.Items.AddRange(New Object() {"BTF", "CSH"})
        Me.CmbFOP1New.Location = New System.Drawing.Point(224, 44)
        Me.CmbFOP1New.Name = "CmbFOP1New"
        Me.CmbFOP1New.Size = New System.Drawing.Size(56, 21)
        Me.CmbFOP1New.TabIndex = 15
        '
        'TxtShortName
        '
        Me.TxtShortName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtShortName.ForeColor = System.Drawing.Color.Blue
        Me.TxtShortName.Location = New System.Drawing.Point(224, 22)
        Me.TxtShortName.Name = "TxtShortName"
        Me.TxtShortName.Size = New System.Drawing.Size(180, 20)
        Me.TxtShortName.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LblCreateVendorAndSupplier)
        Me.GroupBox1.Controls.Add(Me.CmbCity)
        Me.GroupBox1.Controls.Add(Me.CmbCountry)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtPhone)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtEmail)
        Me.GroupBox1.Controls.Add(Me.TxtAddress)
        Me.GroupBox1.Controls.Add(Me.TxtSupplierName)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(4, 109)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(401, 129)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Supplier Infor"
        '
        'LblCreateVendorAndSupplier
        '
        Me.LblCreateVendorAndSupplier.AutoSize = True
        Me.LblCreateVendorAndSupplier.Location = New System.Drawing.Point(250, 108)
        Me.LblCreateVendorAndSupplier.Name = "LblCreateVendorAndSupplier"
        Me.LblCreateVendorAndSupplier.Size = New System.Drawing.Size(137, 13)
        Me.LblCreateVendorAndSupplier.TabIndex = 3
        Me.LblCreateVendorAndSupplier.TabStop = True
        Me.LblCreateVendorAndSupplier.Text = "Create Vendor and Supplier"
        '
        'CmbCity
        '
        Me.CmbCity.FormattingEnabled = True
        Me.CmbCity.Items.AddRange(New Object() {"An Giang", "Bac Giang", "Bac Kan", "Bac Lieu", "Bac Ninh", "Ba Ria Vung Tau", "Ben Tre", "Binh Dinh", "Binh Duong", "Binh Phuoc", "Binh Thuan", "Ca Mau", "Can Tho", "Cao Bang", "Da Nang", "Dak Lack", "Dak Nong", "Dien Bien", "Dong Nai", "Dong Thap", "Gia Lai", "Ha Giang", "Ha Nam", "Ha Noi", "Ha Tay", "Ha Tinh", "Hai Duong", "Hai Phong", "Hau Giang", "HCMC", "Hoa Binh", "Hung Yen", "Khanh Hoa", "Kien Giang", "Kon Tum", "Lai Châu", "Lam Dong", "Lang Son", "Lao Cai", "Long An", "Nam Dinh", "Nghe An", "Ninh Binh", "Ninh Thuan", "Phu Tho", "Phu Yen", "Quang Binh", "Quang Nam", "Quang Ngai", "Quang Ninh", "Quang Tri", "Soc Trang", "Son La", "Tay Ninh", "Thai Binh", "Thai Nguyen", "Thanh Hoa", "Thua Thien - Hue", "Tien Giang", "Tra Vinh", "Tuyen Quang", "Vinh Long", "Vinh Phuc", "Yen Bai"})
        Me.CmbCity.Location = New System.Drawing.Point(52, 38)
        Me.CmbCity.Name = "CmbCity"
        Me.CmbCity.Size = New System.Drawing.Size(193, 21)
        Me.CmbCity.TabIndex = 2
        '
        'CmbCountry
        '
        Me.CmbCountry.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbCountry.FormattingEnabled = True
        Me.CmbCountry.Location = New System.Drawing.Point(246, 16)
        Me.CmbCountry.Name = "CmbCountry"
        Me.CmbCountry.Size = New System.Drawing.Size(154, 21)
        Me.CmbCountry.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 108)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Phone"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Address"
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(52, 105)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(193, 20)
        Me.txtPhone.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 41)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(24, 13)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "City"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 19)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(35, 13)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Name"
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(52, 84)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(349, 20)
        Me.txtEmail.TabIndex = 0
        '
        'TxtAddress
        '
        Me.TxtAddress.Location = New System.Drawing.Point(52, 61)
        Me.TxtAddress.Name = "TxtAddress"
        Me.TxtAddress.Size = New System.Drawing.Size(349, 20)
        Me.TxtAddress.TabIndex = 0
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Location = New System.Drawing.Point(52, 16)
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.Size = New System.Drawing.Size(193, 20)
        Me.TxtSupplierName.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 87)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 13)
        Me.Label7.TabIndex = 1
        Me.Label7.Text = "eMail"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.Blue
        Me.Label9.Location = New System.Drawing.Point(3, 47)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(28, 13)
        Me.Label9.TabIndex = 3
        Me.Label9.Text = "FOP"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Blue
        Me.Label1.Location = New System.Drawing.Point(3, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(23, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Cat"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.ForeColor = System.Drawing.Color.Blue
        Me.Label10.Location = New System.Drawing.Point(160, 47)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 13)
        Me.Label10.TabIndex = 3
        Me.Label10.Text = "DefaultFOP"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Blue
        Me.Label3.Location = New System.Drawing.Point(160, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "ShortName"
        '
        'LblAddNewVendor
        '
        Me.LblAddNewVendor.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAddNewVendor.AutoSize = True
        Me.LblAddNewVendor.Location = New System.Drawing.Point(305, 93)
        Me.LblAddNewVendor.Name = "LblAddNewVendor"
        Me.LblAddNewVendor.Size = New System.Drawing.Size(99, 13)
        Me.LblAddNewVendor.TabIndex = 13
        Me.LblAddNewVendor.TabStop = True
        Me.LblAddNewVendor.Text = "Create Vendor Only"
        '
        'CmbFOPNew
        '
        Me.CmbFOPNew.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbFOPNew.FormattingEnabled = True
        Me.CmbFOPNew.Items.AddRange(New Object() {"PSP", "PPD"})
        Me.CmbFOPNew.Location = New System.Drawing.Point(32, 44)
        Me.CmbFOPNew.Name = "CmbFOPNew"
        Me.CmbFOPNew.Size = New System.Drawing.Size(55, 21)
        Me.CmbFOPNew.TabIndex = 12
        '
        'LblAddCompany
        '
        Me.LblAddCompany.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAddCompany.AutoSize = True
        Me.LblAddCompany.Location = New System.Drawing.Point(-51, 6)
        Me.LblAddCompany.Name = "LblAddCompany"
        Me.LblAddCompany.Size = New System.Drawing.Size(26, 13)
        Me.LblAddCompany.TabIndex = 1
        Me.LblAddCompany.TabStop = True
        Me.LblAddCompany.Text = "Add"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.CmbFOP1edit)
        Me.TabPage2.Controls.Add(Me.CmbHD_tong)
        Me.TabPage2.Controls.Add(Me.TxtNewShortName)
        Me.TabPage2.Controls.Add(Me.CmbINV)
        Me.TabPage2.Controls.Add(Me.LblChangeHD_Tong)
        Me.TabPage2.Controls.Add(Me.LblChangeFOP)
        Me.TabPage2.Controls.Add(Me.LblUpdateHD)
        Me.TabPage2.Controls.Add(Me.LblChangeName)
        Me.TabPage2.Controls.Add(Me.CmbFOPedit)
        Me.TabPage2.Controls.Add(Me.CmbNewCat)
        Me.TabPage2.Controls.Add(Me.LblDeleteCompany)
        Me.TabPage2.Controls.Add(Me.LblChangeCat)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(418, 472)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Edit"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'CmbFOP1edit
        '
        Me.CmbFOP1edit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbFOP1edit.FormattingEnabled = True
        Me.CmbFOP1edit.Items.AddRange(New Object() {"BTF", "CSH"})
        Me.CmbFOP1edit.Location = New System.Drawing.Point(84, 59)
        Me.CmbFOP1edit.Name = "CmbFOP1edit"
        Me.CmbFOP1edit.Size = New System.Drawing.Size(55, 21)
        Me.CmbFOP1edit.TabIndex = 13
        '
        'CmbHD_tong
        '
        Me.CmbHD_tong.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbHD_tong.FormattingEnabled = True
        Me.CmbHD_tong.Items.AddRange(New Object() {"TONG", "CHITIET"})
        Me.CmbHD_tong.Location = New System.Drawing.Point(228, 58)
        Me.CmbHD_tong.Name = "CmbHD_tong"
        Me.CmbHD_tong.Size = New System.Drawing.Size(139, 21)
        Me.CmbHD_tong.TabIndex = 12
        '
        'TxtNewShortName
        '
        Me.TxtNewShortName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtNewShortName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtNewShortName.Location = New System.Drawing.Point(228, 37)
        Me.TxtNewShortName.Name = "TxtNewShortName"
        Me.TxtNewShortName.Size = New System.Drawing.Size(139, 20)
        Me.TxtNewShortName.TabIndex = 10
        '
        'LblChangeHD_Tong
        '
        Me.LblChangeHD_Tong.AutoSize = True
        Me.LblChangeHD_Tong.Location = New System.Drawing.Point(142, 61)
        Me.LblChangeHD_Tong.Name = "LblChangeHD_Tong"
        Me.LblChangeHD_Tong.Size = New System.Drawing.Size(85, 13)
        Me.LblChangeHD_Tong.TabIndex = 11
        Me.LblChangeHD_Tong.TabStop = True
        Me.LblChangeHD_Tong.Text = "HD Tong_chitiet"
        '
        'LblChangeName
        '
        Me.LblChangeName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblChangeName.AutoSize = True
        Me.LblChangeName.Location = New System.Drawing.Point(140, 40)
        Me.LblChangeName.Name = "LblChangeName"
        Me.LblChangeName.Size = New System.Drawing.Size(88, 13)
        Me.LblChangeName.TabIndex = 9
        Me.LblChangeName.TabStop = True
        Me.LblChangeName.Text = "ChangeName To"
        Me.LblChangeName.Visible = False
        '
        'CmbNewCat
        '
        Me.CmbNewCat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CmbNewCat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbNewCat.FormattingEnabled = True
        Me.CmbNewCat.Items.AddRange(New Object() {"AR - Air", "NH - Restaurant", "KS - Hotel", "XE - Coach", "BO - Boat", "LD - Land", "OS - OtherTour Service", "TO - TO/TA", "OT - Other", "US - Utility Supplier", "AL - Airlines", "AP - Airport", "YT- So YT"})
        Me.CmbNewCat.Location = New System.Drawing.Point(228, 10)
        Me.CmbNewCat.Name = "CmbNewCat"
        Me.CmbNewCat.Size = New System.Drawing.Size(139, 21)
        Me.CmbNewCat.TabIndex = 8
        '
        'LblDeleteCompany
        '
        Me.LblDeleteCompany.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblDeleteCompany.AutoSize = True
        Me.LblDeleteCompany.Location = New System.Drawing.Point(373, 13)
        Me.LblDeleteCompany.Name = "LblDeleteCompany"
        Me.LblDeleteCompany.Size = New System.Drawing.Size(38, 13)
        Me.LblDeleteCompany.TabIndex = 2
        Me.LblDeleteCompany.TabStop = True
        Me.LblDeleteCompany.Text = "Delete"
        Me.LblDeleteCompany.Visible = False
        '
        'LblChangeCat
        '
        Me.LblChangeCat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblChangeCat.AutoSize = True
        Me.LblChangeCat.Location = New System.Drawing.Point(142, 13)
        Me.LblChangeCat.Name = "LblChangeCat"
        Me.LblChangeCat.Size = New System.Drawing.Size(80, 13)
        Me.LblChangeCat.TabIndex = 7
        Me.LblChangeCat.TabStop = True
        Me.LblChangeCat.Text = "Change CAT to"
        Me.LblChangeCat.Visible = False
        '
        'ChckCompanyOKonly
        '
        Me.ChckCompanyOKonly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChckCompanyOKonly.AutoSize = True
        Me.ChckCompanyOKonly.Checked = True
        Me.ChckCompanyOKonly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChckCompanyOKonly.Location = New System.Drawing.Point(4, 477)
        Me.ChckCompanyOKonly.Name = "ChckCompanyOKonly"
        Me.ChckCompanyOKonly.Size = New System.Drawing.Size(65, 17)
        Me.ChckCompanyOKonly.TabIndex = 1
        Me.ChckCompanyOKonly.Text = "OK Only"
        Me.ChckCompanyOKonly.UseVisualStyleBackColor = True
        '
        'VendorInfor_KTT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 496)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.GridVendor)
        Me.Controls.Add(Me.ChckCompanyOKonly)
        Me.Controls.Add(Me.TxtNameFilter)
        Me.Controls.Add(Me.LblFilter)
        Me.Controls.Add(Me.Label2)
        Me.Location = New System.Drawing.Point(0, 56)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "VendorInfor_KTT"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "TransViet Travel :: RAS 12 :. Vendor Infor Update"
        CType(Me.GridVendor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GridVendor As System.Windows.Forms.DataGridView
    Friend WithEvents CmbCAT As System.Windows.Forms.ComboBox
    Friend WithEvents CmbINV As System.Windows.Forms.ComboBox
    Friend WithEvents LblUpdateHD As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtNameFilter As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LblFilter As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbFOPedit As System.Windows.Forms.ComboBox
    Friend WithEvents LblChangeFOP As System.Windows.Forms.LinkLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TxtNewShortName As System.Windows.Forms.TextBox
    Friend WithEvents CmbFOPNew As System.Windows.Forms.ComboBox
    Friend WithEvents LblChangeName As System.Windows.Forms.LinkLabel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CmbNewCat As System.Windows.Forms.ComboBox
    Friend WithEvents TxtShortName As System.Windows.Forms.TextBox
    Friend WithEvents LblChangeCat As System.Windows.Forms.LinkLabel
    Friend WithEvents ChckCompanyOKonly As System.Windows.Forms.CheckBox
    Friend WithEvents LblAddCompany As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDeleteCompany As System.Windows.Forms.LinkLabel
    Friend WithEvents LblAddNewVendor As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LblCreateVendorAndSupplier As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbCity As System.Windows.Forms.ComboBox
    Friend WithEvents CmbCountry As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents TxtAddress As System.Windows.Forms.TextBox
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CmbHD_tong As System.Windows.Forms.ComboBox
    Friend WithEvents LblChangeHD_Tong As System.Windows.Forms.LinkLabel
    Friend WithEvents CmbFOP1New As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CmbFOP1edit As System.Windows.Forms.ComboBox
End Class
