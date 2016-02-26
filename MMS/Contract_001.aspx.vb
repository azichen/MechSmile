Imports System.Data

Partial Class MMS_Contract_001
    Inherits System.Web.UI.Page
    Dim EditingMode As String = ""


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else
                Dim sFileName As String = Request.Path.Substring(Request.Path.LastIndexOf("/") + 1, Request.Path.Length - (Request.Path.LastIndexOf("/") + 1))
                Dim sRol As String = HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
                'ViewState("ROL") = modUtil.GetRolData(sFileName, sRol)
                modUtil.GetRolData(Me.Request, Me.ViewState("ROL"), Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()

            End If

            Dim tmpsql As String
            Me.DropDownList5.DataSourceID = Nothing
            tmpsql = "SELECT AreaCode, AreaCode+'_'+AreaName as AreaName FROM MMSArea WHERE (Effective = 'Y')"
            Me.DropDownList5.DataMember = "AreaCode"
            Me.DropDownList5.DataTextField = "AreaName"
            Me.DropDownList5.DataSource = get_DataTable(tmpsql)
            Me.DropDownList5.DataBind()
            EditingMode = ""

            Select Case Request("FormMode")
                Case "add"
                    EditingMode = "add"
                    CheckBox1.Checked = True
                    CheckBox1.Visible = True
                    'modUtil.SetObjReadOnly(Me, "ContractNoTextBox")
                    Me.TitleLabel.Text = "合約基本資料 - 新增"
                    Me.ButtonSave.Visible = True
                    Me.Button2.Visible = False
                    Me.Button5.Visible = False
                    CustomerNoTextBox.Focus()
                    If GetAreaCode() = "" Then
                        Me.DropDownList5.Enabled = True
                    Else
                        Me.DropDownList5.SelectedValue = GetAreaCode()
                        Me.dispEmployee()
                    End If
                    Me.Button7.Visible = False
                    If ViewState("ROL").Substring(1, 1) = "N" Then
                        Me.ButtonSave.Visible = False
                    Else
                        Me.ButtonSave.Visible = True
                    End If
                    Me.Button7.Visible = False
                    Me.Button8.Visible = True
                    '20140905
                    If GetAreaCode() <> "" Then
                        Dim i As Integer
                        For i = 0 To Me.DropDownList5.Items.Count - 1
                            If Me.DropDownList5.Items(i).Value = GetAreaCode() Then
                                Me.DropDownList5.SelectedIndex = i
                            End If
                        Next
                        Me.DropDownList5.Enabled = False
                    End If
                Case "edit"
                    dispDetail()
            End Select
            'Try
            '    If ViewState("ROL").Substring(2, 1) = "N" Then Me.ButtonSave.Visible = False
            'Catch ex As Exception

            'End Try
            '設定日期輸入及檢查
            modUtil.SetDateObj(Me.StartDateCTextBox, False, Me.EndDateCTextBox, False)
            modUtil.SetDateObj(Me.EndDateCTextBox, False, Nothing, False)
            modUtil.SetDateObj(Me.StartDateITextBox, False, Me.EndDateITextBox, False)
            modUtil.SetDateObj(Me.EndDateITextBox, False, Nothing, False)
            modUtil.SetDateObj(Me.ArchiveDateMTextBox, False, Nothing, False)
            modUtil.SetDateObj(Me.ArchiveDateATextBox, False, Nothing, False)
            modUtil.SetDateObj(Me.CancelDateTextBox, False, Nothing, False)
            modUtil.SetDateObj(Me.SealDateTextBox, False, Nothing, False)

            '設定唯讀
            modUtil.SetObjReadOnly(Me, "CustomerNameTextBox")
            'modUtil.SetObjReadOnly(Me, "ContractNoTextBox")
            modUtil.SetObjReadOnly(Me, "OldContractNoTextBox")

            '設定只可輸入數字
            Me.QuantityTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TotalQuantityTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.PricePerMonthTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.UnitPricePerMonthTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.AmountOfContractTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.DaysAllowedTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.PeriodMaintenanceAmountTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.MaintenanceAmountTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.AdjAmountOfContractTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.GUINumberTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TotalOfMonthsTextBox.Attributes.Add("onkeypress", "return CheckKeyNumber()")


            Me.OldContractNoTextBox.Attributes.Add("onkeypress", "return ToUpper()")
            Me.ContractNoTextBox.Attributes.Add("onkeypress", "return ToUpper()")
            Me.CustomerNoTextBox.Attributes.Add("onkeypress", "return ToUpper()")

        End If
        If Me.txtRet.Text.Trim <> "" Then
            Try
                Dim splitStr As String() = Me.txtRet.Text.Split(",")
                CustomerNoTextBox.Text = splitStr(0)
                CustomerNameTextBox.Text = splitStr(1)
                Me.popCmp.Hide()
            Catch ex As Exception
                Dim ss As String = ""
            End Try
        End If
        countdate()

        If GetAreaCode() <> "" Then
            modUtil.SetObjReadOnly(Me, "ArchiveDateATextBox")
            modUtil.SetObjReadOnly(Me, "SealDateTextBox")
            modUtil.SetObjReadOnly(Me, "CancelDateTextBox")
        End If

    End Sub

    Private Sub dispDetail()
        If ViewState("ROL").Substring(1, 1) = "N" Then
            Me.Button5.Visible = False
        Else
            Me.Button5.Visible = True
        End If

        modUtil.SetObjReadOnly(Me, "ContractNoTextBox")
        'Me.TitleLabel.Text = "合約基本資料 - 修改"
        Me.TitleLabel.Text = "合約基本資料 - 瀏覽"
        Dim tmpsql As String
        tmpsql = "select * from  MMSContract where ContractNo='" + Request("STNID").Trim + "'"
        Me.SqlDataSource1.SelectCommand = tmpsql
        Me.SqlDataSource1.DataBind()
        Me.ButtonSave.Visible = False
        Me.Button2.Visible = True

        Me.Button4.Visible = False

        modUtil.SetObjReadOnly(Me, "CustomerNoTextBox") '不可修改
        DropDownList5.Enabled = True


        tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
        Me.SqlDataSource2.SelectCommand = tmpsql
        Me.SqlDataSource2.DataBind()
        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee" ' WHERE Effective = 'Y' "
        Me.SqlDataSource3.SelectCommand = tmpsql
        Me.SqlDataSource3.DataBind()
        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee" ' WHERE Effective = 'Y' "
        Me.SqlDataSource4.SelectCommand = tmpsql
        Me.SqlDataSource4.DataBind()

        Dim dv As Data.DataView
        dv = Me.SqlDataSource1.Select(New DataSourceSelectArguments())
        '第一筆第一個欄位得資料 
        'Me.ContractNoTextBox.Text = dv(0)("CustomerNo").ToString()
        Try
            ContractNoTextBox.Text = dv(0)("ContractNo").ToString()
            CustomerNoTextBox.Text = dv(0)("CustomerNo").ToString()
            CustomerNameTextBox.Text = dv(0)("CustomerName").ToString()
            TelNoTextBox.Text = dv(0)("TelNo").ToString()
            ContactTextBox.Text = dv(0)("Contact").ToString()
            DropDownList5.SelectedValue = dv(0)("AreaCode").ToString()
            Me.DropDownList6.SelectedValue = dv(0)("Salesman").ToString()
            Me.DropDownList7.SelectedValue = dv(0)("Cashier").ToString()
            Me.TextBox1.Text = dv(0)("Salesman").ToString()
            Me.TextBox2.Text = dv(0)("Cashier").ToString()
            OldContractNoTextBox.Text = dv(0)("OldContractNo").ToString()
            RadioButtonList2.SelectedValue = dv(0)("ContractType").ToString()
            BuildingNameTextBox.Text = dv(0)("BuildingName").ToString()
            AddressTextBox.Text = dv(0)("Address").ToString()
            SpecificationNoTextBox.Text = dv(0)("SpecificationNo").ToString()
            QuantityTextBox.Text = dv(0)("Quantity").ToString()
            StartDateCTextBox.Text = dv(0)("StartDateC").ToString()
            EndDateCTextBox.Text = dv(0)("EndDateC").ToString()
            TotalOfMonthsTextBox.Text = dv(0)("TotalOfMonths").ToString()
            UnitPricePerMonthTextBox.Text = dv(0)("UnitPricePerMonth").ToString()
            TotalQuantityTextBox.Text = dv(0)("TotalQuantity").ToString()
            PricePerMonthTextBox.Text = dv(0)("PricePerMonth").ToString()
            AmountOfContractTextBox.Text = dv(0)("AmountOfContract").ToString()
            PeriodMaintenanceAmountTextBox.Text = dv(0)("PeriodMaintenanceAmount").ToString()
            MaintenanceAmountTextBox.Text = dv(0)("MaintenanceAmount").ToString()
            AdjAmountOfContractTextBox.Text = dv(0)("AdjAmountOfContract").ToString()
            DropDownList8.SelectedValue = dv(0)("PaymentMethod").ToString()
            DaysAllowedTextBox.Text = dv(0)("DaysAllowed").ToString()
            GUINumberTextBox.Text = dv(0)("GUINumber").ToString()
            TitleOfInvoiceTextBox.Text = dv(0)("TitleOfInvoice").ToString()
            AddressOfInvoiceTextBox.Text = dv(0)("AddressOfInvoice").ToString()
            StartDateITextBox.Text = dv(0)("StartDateI").ToString()
            EndDateITextBox.Text = dv(0)("EndDateI").ToString()
            DropDownList4.SelectedValue = dv(0)("InvoiceCycle").ToString()
            ItemNameTextBox.Text = dv(0)("ItemName").ToString()
            RadioButtonList3.SelectedValue = dv(0)("DatePlus").ToString()
            PayerTextBox.Text = dv(0)("Payer").ToString()
            PayerTelTextBox.Text = dv(0)("PayerTel").ToString()
            PayerAdderssTextBox.Text = dv(0)("PayerAdderss").ToString()
            ItemNoTextBox.Text = dv(0)("ItemNo").ToString()
            SealDateTextBox.Text = dv(0)("SealDate").ToString()
            ArchiveDateMTextBox.Text = dv(0)("ArchiveDateM").ToString()
            ArchiveDateATextBox.Text = dv(0)("ArchiveDateA").ToString()
            CancelDateTextBox.Text = dv(0)("CancelDate").ToString()
            MemoTextBox.Text = dv(0)("Memo").ToString()
            Memo2TextBox.Text = dv(0)("Memo2").ToString()
        Catch ex As Exception

        End Try
        If dv(0)("ArchiveDateM").ToString() <> "" And dv(0)("ArchiveDateA").ToString() <> "" Then
            DisableAll()
        Else
            If dv(0)("LastInvoiceDate").ToString() <> "" Then
                DisableAll1()
            End If
        End If
        If CancelDateTextBox.Text <> "" Then
            Button2.Visible = False
        End If
        TelNoTextBox.Focus()
        DisableFull()


        If ViewState("ROL").Substring(2, 1) = "N" Then
            Me.Button7.Visible = False
        Else
            Me.Button7.Visible = True
        End If
        Me.Button8.Visible = False
    End Sub

    Private Function GetAreaCode() As String
        '    Dim tmpsql As String
        '    Dim STNID As String = Request.Cookies("STNID").Value
        '    tmpsql = "select * from MECHSTNM where STNID='" + STNID + "'"
        '    If get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "    " Then 'Or get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "4601"
        '        tmpsql = "select * from MMSArea where ProfitNo='" + get_DataTable(tmpsql).Rows(0).Item("STNUID") + "'"
        '        Dim tmparea As String
        '        Try
        '            GetAreaCode = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
        '        Catch ex As Exception
        '            GetAreaCode = ""
        '        End Try
        '    Else
        '        GetAreaCode = ""
        '    End If
        '    'GetAreaCode = "C"
        Dim Areacode As String = ""
        GetAreaCode = modUnset.GetAreaCode(My.User.Name).Trim
    End Function

    '修改時用續約
    Private Sub EnableAll1()

        modUtil.SetObjEnabled(ContractNoTextBox)
        modUtil.SetObjEnabled(CustomerNoTextBox)
        'modUtil.SetObjEnabled(CustomerNameTextBox) 'mark in 20140729
        modUtil.SetObjEnabled(TelNoTextBox)
        modUtil.SetObjEnabled(ContactTextBox)
        DropDownList5.Enabled = True
        DropDownList4.Enabled = True
        DropDownList6.Enabled = True
        DropDownList7.Enabled = True
        modUtil.SetObjEnabled(TextBox1)
        modUtil.SetObjEnabled(TextBox2)
        modUtil.SetObjEnabled(OldContractNoTextBox)
        RadioButtonList2.Enabled = True
        modUtil.SetObjEnabled(BuildingNameTextBox)
        modUtil.SetObjEnabled(AddressTextBox)
        modUtil.SetObjEnabled(SpecificationNoTextBox)
        modUtil.SetObjEnabled(QuantityTextBox)
        modUtil.SetObjEnabled(StartDateCTextBox)
        modUtil.SetObjEnabled(EndDateCTextBox)
        modUtil.SetObjEnabled(TotalOfMonthsTextBox)
        modUtil.SetObjEnabled(UnitPricePerMonthTextBox)
        modUtil.SetObjEnabled(TotalQuantityTextBox)
        modUtil.SetObjEnabled(PricePerMonthTextBox)
        modUtil.SetObjEnabled(AmountOfContractTextBox)
        modUtil.SetObjEnabled(PeriodMaintenanceAmountTextBox)
        modUtil.SetObjEnabled(MaintenanceAmountTextBox)
        modUtil.SetObjEnabled(AdjAmountOfContractTextBox)
        DropDownList8.Enabled = True
        modUtil.SetObjEnabled(DaysAllowedTextBox)
        modUtil.SetObjEnabled(GUINumberTextBox)
        modUtil.SetObjEnabled(TitleOfInvoiceTextBox)
        modUtil.SetObjEnabled(AddressOfInvoiceTextBox)
        modUtil.SetObjEnabled(StartDateITextBox)
        modUtil.SetObjEnabled(EndDateITextBox)
        DropDownList4.Enabled = True
        modUtil.SetObjEnabled(ItemNameTextBox)
        RadioButtonList3.Enabled = True
        modUtil.SetObjEnabled(PayerTextBox)
        modUtil.SetObjEnabled(PayerTelTextBox)
        modUtil.SetObjEnabled(PayerAdderssTextBox)
        modUtil.SetObjEnabled(ItemNoTextBox)
        modUtil.SetObjEnabled(SealDateTextBox)
        modUtil.SetObjEnabled(ArchiveDateMTextBox)
        modUtil.SetObjEnabled(ArchiveDateATextBox)
        modUtil.SetObjEnabled(CancelDateTextBox)
        modUtil.SetObjEnabled(MemoTextBox)
        modUtil.SetObjEnabled(Memo2TextBox)

        Me.Image1.Visible = True
        Me.Image2.Visible = True
        Me.Image3.Visible = True
        Me.Image4.Visible = True
        Me.Image5.Visible = True
        Me.Image6.Visible = True
        Me.Image7.Visible = True
        Me.Image8.Visible = True
    End Sub

    '修改時用發票已開立
    Private Sub DisableAll1()
        modUtil.SetObjReadOnly(Me, "ContractNoTextBox")
        modUtil.SetObjReadOnly(Me, "CustomerNoTextBox")
        modUtil.SetObjReadOnly(Me, "CustomerNameTextBox")
        modUtil.SetObjReadOnly(Me, "TelNoTextBox")
        modUtil.SetObjReadOnly(Me, "ContactTextBox")
        DropDownList5.Enabled = False
        DropDownList4.Enabled = False
        'DropDownList7.Enabled = False
        modUtil.SetObjReadOnly(Me, "TextBox1")
        modUtil.SetObjReadOnly(Me, "TextBox2")
        modUtil.SetObjReadOnly(Me, "OldContractNoTextBox")
        RadioButtonList2.Enabled = False
        modUtil.SetObjReadOnly(Me, "BuildingNameTextBox")
        modUtil.SetObjReadOnly(Me, "AddressTextBox")
        modUtil.SetObjReadOnly(Me, "SpecificationNoTextBox")
        modUtil.SetObjReadOnly(Me, "QuantityTextBox")
        modUtil.SetObjReadOnly(Me, "StartDateCTextBox")
        modUtil.SetObjReadOnly(Me, "EndDateCTextBox")
        modUtil.SetObjReadOnly(Me, "TotalOfMonthsTextBox")
        modUtil.SetObjReadOnly(Me, "UnitPricePerMonthTextBox")
        modUtil.SetObjReadOnly(Me, "TotalQuantityTextBox")
        modUtil.SetObjReadOnly(Me, "PricePerMonthTextBox")
        modUtil.SetObjReadOnly(Me, "AmountOfContractTextBox")
        modUtil.SetObjReadOnly(Me, "PeriodMaintenanceAmountTextBox")
        modUtil.SetObjReadOnly(Me, "MaintenanceAmountTextBox")
        AdjAmountOfContractTextBox.Enabled = True
        DropDownList8.Enabled = False
        modUtil.SetObjReadOnly(Me, "DaysAllowedTextBox")
        modUtil.SetObjReadOnly(Me, "GUINumberTextBox")
        modUtil.SetObjReadOnly(Me, "TitleOfInvoiceTextBox")
        modUtil.SetObjReadOnly(Me, "AddressOfInvoiceTextBox")
        'MARK IN 20150901 開放發票開例日期起迄，因為修改率很高，相關修改已完成
        'modUtil.SetObjReadOnly(Me, "StartDateITextBox")
        'modUtil.SetObjReadOnly(Me, "EndDateITextBox")
        'Me.Image3.Visible = False
        'Me.Image4.Visible = False
        DropDownList4.Enabled = False
        modUtil.SetObjReadOnly(Me, "ItemNameTextBox")
        RadioButtonList3.Enabled = False
        modUtil.SetObjReadOnly(Me, "PayerTextBox")
        modUtil.SetObjReadOnly(Me, " PayerTelTextBox")
        modUtil.SetObjReadOnly(Me, " PayerAdderssTextBox")
        modUtil.SetObjReadOnly(Me, " ItemNoTextBox")
        modUtil.SetObjReadOnly(Me, " SealDateTextBox")
        modUtil.SetObjReadOnly(Me, "ArchiveDateMTextBox")
        modUtil.SetObjReadOnly(Me, "ArchiveDateATextBox")
        modUtil.SetObjReadOnly(Me, "CancelDateTextBox")
        modUtil.SetObjReadOnly(Me, "MemoTextBox")
        modUtil.SetObjReadOnly(Me, "Memo2TextBox")
        Me.Image1.Visible = False
        Me.Image2.Visible = False
        
    End Sub

    '修改時機械\會計已歸檔
    Private Sub DisableAll()
        modUtil.SetObjReadOnly(Me, "ContractNoTextBox")
        modUtil.SetObjReadOnly(Me, "CustomerNoTextBox")
        modUtil.SetObjReadOnly(Me, "CustomerNameTextBox")
        TelNoTextBox.Enabled = True
        ContactTextBox.Enabled = True
        DropDownList5.Enabled = False
        'DropDownList6.Enabled = False
        'DropDownList7.Enabled = False
        modUtil.SetObjReadOnly(Me, "TextBox1")
        modUtil.SetObjReadOnly(Me, "TextBox2")
        modUtil.SetObjReadOnly(Me, "OldContractNoTextBox")
        RadioButtonList2.Enabled = False
        modUtil.SetObjReadOnly(Me, "BuildingNameTextBox")
        modUtil.SetObjReadOnly(Me, "AddressTextBox")
        modUtil.SetObjReadOnly(Me, "SpecificationNoTextBox")
        modUtil.SetObjReadOnly(Me, "QuantityTextBox")
        modUtil.SetObjReadOnly(Me, "StartDateCTextBox")
        modUtil.SetObjReadOnly(Me, "EndDateCTextBox")
        modUtil.SetObjReadOnly(Me, "TotalOfMonthsTextBox")
        modUtil.SetObjReadOnly(Me, "UnitPricePerMonthTextBox")
        modUtil.SetObjReadOnly(Me, "TotalQuantityTextBox")
        modUtil.SetObjReadOnly(Me, "PricePerMonthTextBox")
        modUtil.SetObjReadOnly(Me, "AmountOfContractTextBox")
        modUtil.SetObjReadOnly(Me, "PeriodMaintenanceAmountTextBox")
        modUtil.SetObjReadOnly(Me, "MaintenanceAmountTextBox")
        modUtil.SetObjReadOnly(Me, "AdjAmountOfContractTextBox")
        DropDownList8.Enabled = False
        DropDownList4.Enabled = False
        modUtil.SetObjReadOnly(Me, "DaysAllowedTextBox")
        modUtil.SetObjReadOnly(Me, "GUINumberTextBox")
        modUtil.SetObjReadOnly(Me, "TitleOfInvoiceTextBox")
        modUtil.SetObjReadOnly(Me, "AddressOfInvoiceTextBox")
        modUtil.SetObjReadOnly(Me, "StartDateITextBox")
        modUtil.SetObjReadOnly(Me, "EndDateITextBox")
        modUtil.SetObjReadOnly(Me, "DropDownList4")
        modUtil.SetObjReadOnly(Me, "ItemNameTextBox")
        RadioButtonList3.Enabled = False
        modUtil.SetObjReadOnly(Me, "PayerTextBox")
        modUtil.SetObjReadOnly(Me, "PayerTelTextBox")
        modUtil.SetObjReadOnly(Me, "PayerAdderssTextBox")
        modUtil.SetObjReadOnly(Me, "ItemNoTextBox")
        modUtil.SetObjReadOnly(Me, "SealDateTextBox")
        modUtil.SetObjReadOnly(Me, "ArchiveDateMTextBox")
        modUtil.SetObjReadOnly(Me, "ArchiveDateATextBox")
        modUtil.SetObjReadOnly(Me, "CancelDateTextBox")
        modUtil.SetObjReadOnly(Me, "MemoTextBox")
        modUtil.SetObjReadOnly(Me, "Memo2TextBox")
        modUtil.SetObjReadOnly(Me, "StartDateCTextBox")
        modUtil.SetObjReadOnly(Me, "EndDateCTextBox")

        Me.Image1.Visible = False
        Me.Image2.Visible = False
        Me.Image3.Visible = False
        Me.Image4.Visible = False
        Me.Image5.Visible = False
        Me.Image6.Visible = False
        Me.Image8.Visible = False
    End Sub

    '瀏覽資料
    Private Sub DisableFull()
        modUtil.SetObjReadOnly(Me, "ContractNoTextBox")
        modUtil.SetObjReadOnly(Me, "CustomerNoTextBox")
        modUtil.SetObjReadOnly(Me, "CustomerNameTextBox")
        modUtil.SetObjReadOnly(Me, "TelNoTextBox")
        modUtil.SetObjReadOnly(Me, "ContactTextBox")
        DropDownList5.Enabled = False
        DropDownList6.Enabled = False
        DropDownList7.Enabled = False
        modUtil.SetObjReadOnly(Me, "TextBox1")
        modUtil.SetObjReadOnly(Me, "TextBox2")
        modUtil.SetObjReadOnly(Me, "OldContractNoTextBox")
        RadioButtonList2.Enabled = False
        modUtil.SetObjReadOnly(Me, "BuildingNameTextBox")
        modUtil.SetObjReadOnly(Me, "AddressTextBox")
        modUtil.SetObjReadOnly(Me, "SpecificationNoTextBox")
        modUtil.SetObjReadOnly(Me, "QuantityTextBox")
        modUtil.SetObjReadOnly(Me, "StartDateCTextBox")
        modUtil.SetObjReadOnly(Me, "EndDateCTextBox")
        modUtil.SetObjReadOnly(Me, "TotalOfMonthsTextBox")
        modUtil.SetObjReadOnly(Me, "UnitPricePerMonthTextBox")
        modUtil.SetObjReadOnly(Me, "TotalQuantityTextBox")
        modUtil.SetObjReadOnly(Me, "PricePerMonthTextBox")
        modUtil.SetObjReadOnly(Me, "AmountOfContractTextBox")
        modUtil.SetObjReadOnly(Me, "PeriodMaintenanceAmountTextBox")
        modUtil.SetObjReadOnly(Me, "MaintenanceAmountTextBox")
        modUtil.SetObjReadOnly(Me, "AdjAmountOfContractTextBox")
        DropDownList8.Enabled = False
        DropDownList4.Enabled = False
        modUtil.SetObjReadOnly(Me, "DaysAllowedTextBox")
        modUtil.SetObjReadOnly(Me, "GUINumberTextBox")
        modUtil.SetObjReadOnly(Me, "TitleOfInvoiceTextBox")
        modUtil.SetObjReadOnly(Me, "AddressOfInvoiceTextBox")
        modUtil.SetObjReadOnly(Me, "StartDateITextBox")
        modUtil.SetObjReadOnly(Me, "EndDateITextBox")
        modUtil.SetObjReadOnly(Me, "DropDownList4")
        modUtil.SetObjReadOnly(Me, "ItemNameTextBox")
        RadioButtonList3.Enabled = False
        modUtil.SetObjReadOnly(Me, "PayerTextBox")
        modUtil.SetObjReadOnly(Me, "PayerTelTextBox")
        modUtil.SetObjReadOnly(Me, "PayerAdderssTextBox")
        modUtil.SetObjReadOnly(Me, "ItemNoTextBox")
        modUtil.SetObjReadOnly(Me, "SealDateTextBox")
        modUtil.SetObjReadOnly(Me, "ArchiveDateMTextBox")
        modUtil.SetObjReadOnly(Me, "ArchiveDateATextBox")
        modUtil.SetObjReadOnly(Me, "CancelDateTextBox")
        modUtil.SetObjReadOnly(Me, "MemoTextBox")
        modUtil.SetObjReadOnly(Me, "Memo2TextBox")
        modUtil.SetObjReadOnly(Me, "StartDateCTextBox")
        modUtil.SetObjReadOnly(Me, "EndDateCTextBox")

        Me.Image1.Visible = False
        Me.Image2.Visible = False
        Me.Image3.Visible = False
        Me.Image4.Visible = False
        Me.Image5.Visible = False
        Me.Image6.Visible = False
        Me.Image7.Visible = False
        Me.Image8.Visible = False


        Me.Button2.Visible = False
    End Sub


    '檢查輸入資料
    Private Function CheckData_detail() As String
        CheckData_detail = ""

        If DropDownList5.SelectedValue = "" Then
            CheckData_detail += "請選擇區域! \n"
        End If
        If DropDownList6.SelectedValue = "" Then
            CheckData_detail += "請選擇業務員! \n"
        End If
        If DropDownList7.SelectedValue = "" Then
            CheckData_detail += "請選擇收款員! \n"
        End If
        If CustomerNoTextBox.Text.Trim = "" Then
            CheckData_detail += "客戶代號不可空白! \n"
        End If
        If CustomerNameTextBox.Text.Trim = "" Then
            CheckData_detail += "客戶名稱不可空白! \n"
        End If
        If BuildingNameTextBox.Text.Trim = "" Then
            CheckData_detail += "大樓名稱不可空白! \n"
        End If
        If AddressTextBox.Text.Trim = "" Then
            CheckData_detail += "地址不可空白! \n"
        End If
        'If QuantityTextBox.Text.Trim = "" Then
        '    CheckData_detail += "台數不可空白或為0! \n"
        'Else
        '    If Int(QuantityTextBox.Text.Trim) = 0 Then
        '        CheckData_detail += "台數不可空白或為0! \n"
        '    End If
        'End If
        'If TotalQuantityTextBox.Text.Trim = "" Then
        '    CheckData_detail += "總台數不可空白或為0! \n"
        'Else
        '    If Int(TotalQuantityTextBox.Text.Trim) = 0 Then
        '        CheckData_detail += "總台數不可空白或為0! \n"
        '    End If
        'End If
        If AmountOfContractTextBox.Text.Trim = "" Then
            CheckData_detail += "合約總金額不可空白或為0! \n"
        Else
            If Int(AmountOfContractTextBox.Text.Trim) = 0 Then
                CheckData_detail += "合約總金額不可空白或為0! \n"
            End If
        End If
        If StartDateCTextBox.Text.Trim = "" Then
            CheckData_detail += "合約期間起不可空白! \n"
        End If
        If EndDateCTextBox.Text.Trim = "" Then
            CheckData_detail += "合約期間迄不可空白! \n"
        End If
        'If UnitPricePerMonthTextBox.Text.Trim = "" Then
        '    CheckData_detail += "每台每月金額不可空白或為0! \n"
        'Else
        '    If Int(UnitPricePerMonthTextBox.Text.Trim) = 0 Then
        '        CheckData_detail += "每台每月金額不可空白或為0! \n"
        '    End If

        'End If
        If DaysAllowedTextBox.Text.Trim = "" Then
            CheckData_detail += "收款日數不可空白或為0! \n"
        Else
            If Int(DaysAllowedTextBox.Text.Trim) = 0 Then
                CheckData_detail += "收款日數不可空白或為0! \n"
            End If

        End If
        If PeriodMaintenanceAmountTextBox.Text.Trim = "" Then
            CheckData_detail += "每期保養金額不可空白或為0! \n"
        Else
            If Int(PeriodMaintenanceAmountTextBox.Text.Trim) = 0 Then
                CheckData_detail += "每期保養金額不可空白或為0! \n"
            End If
        End If
        If MaintenanceAmountTextBox.Text.Trim = "" Then
            CheckData_detail += "總保養金額不可空白或為0! \n"
        Else
            If Int(MaintenanceAmountTextBox.Text.Trim) = 0 Then
                CheckData_detail += "總保養金額不可空白或為0! \n"
            End If
        End If
        If StartDateITextBox.Text.Trim = "" Then
            CheckData_detail += "開立期間起不可空白! \n"
        End If
        If EndDateITextBox.Text.Trim = "" Then
            CheckData_detail += "開立期間迄不可空白! \n"
        End If
        If ItemNameTextBox.Text.Trim = "" Then
            CheckData_detail += "發票品名不可空白! \n"
        End If
        If PayerTextBox.Text.Trim = "" Then
            CheckData_detail += "付款人不可空白! \n"
        End If
        If ItemNoTextBox.Text.Trim = "" Then
            CheckData_detail += "產品別不可空白! \n"
        End If

        If StartDateCTextBox.Text.Trim > EndDateCTextBox.Text.Trim Then
            CheckData_detail += "合約期間迄之結束日期不可大於起始日期! \n"
        End If

        If StartDateITextBox.Text.Trim > EndDateITextBox.Text.Trim Then
            CheckData_detail += "開立期間迄之結束日期不可大於起始日期! \n"
        End If

        If modUtil.BANCheck(GUINumberTextBox.Text) = False And GUINumberTextBox.Text.Trim <> "" Then
            CheckData_detail += "統一編號檢查錯誤! \n"
        End If
        '
    End Function

    '顯示區域所屬業務員收款員
    Private Sub dispEmployee()
        Dim tmpsql As String
        If Me.DropDownList5.SelectedValue = "" And Me.DropDownList5.SelectedIndex = -1 Then
            Try
                tmpsql = "SELECT AreaCode, AreaCode+'_'+AreaName  as AreaName FROM MMSArea WHERE (Effective = 'Y')"
                Me.SqlDataSource2.SelectCommand = tmpsql
                Me.SqlDataSource2.DataBind()
                Me.DropDownList5.SelectedIndex = 0
            Catch ex As Exception
                Dim sss As String = ""
            End Try
        End If

        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee"
        tmpsql += " WHERE Effective = 'Y' and AreaCode='" + Me.DropDownList5.SelectedValue + "'"
        tmpsql += " and Salesman = 'Y'"
        Me.SqlDataSource3.SelectCommand = tmpsql
        Me.SqlDataSource3.DataBind()

        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee"
        tmpsql += " WHERE Effective = 'Y' and AreaCode='" + Me.DropDownList5.SelectedValue + "'"
        tmpsql += " and Cashier = 'Y'"
        Me.SqlDataSource4.SelectCommand = tmpsql
        Me.SqlDataSource4.DataBind()
        'Me.DropDownList6.Enabled = True
        'Me.DropDownList7.Enabled = True
    End Sub

    'Protected Sub DropDownList5_DataBound(sender As Object, e As System.EventArgs) Handles DropDownList5.DataBound
    '    dispEmployee()
    'End Sub

    Protected Sub DropDownList5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList5.SelectedIndexChanged
        dispEmployee()
    End Sub

    '取消
    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Select Case Button3.Text
            Case "回主畫面"
                Response.Redirect("Contract.aspx?Returnflag=1&msg=")
            Case "取消設定"
                Me.Button3.Text = "回主畫面"
                dispDetail()
        End Select

    End Sub

    Private Function CheckContractNo() As String
        CheckContractNo = ""
        If Me.ContractNoTextBox.Text.Trim.Length <> 14 Then
            CheckContractNo += "合約編號長度不足!! \n"
            Exit Function
        End If
        If Mid(ContractNoTextBox.Text.Trim, 1, 1) <> "S" Then
            CheckContractNo += "合約編號第一碼錯誤!! \n"
        End If
        If Mid(ContractNoTextBox.Text.Trim, 2, 1) <> "E" And Mid(ContractNoTextBox.Text.Trim, 2, 1) <> "P" Then
            CheckContractNo += "合約編號第二碼錯誤!! \n"
        End If
        If Mid(ContractNoTextBox.Text.Trim, 3, 6) <> CustomerNoTextBox.Text.Trim Then
            CheckContractNo += "合約編號第3至8碼錯誤!! \n"
        End If
        If Mid(ContractNoTextBox.Text.Trim, 9, 1) <> "-" Or Mid(ContractNoTextBox.Text.Trim, 12, 1) <> "-" Then
            CheckContractNo += "合約編號第9或12碼錯誤!! \n"
        End If
        Dim tmpsql As String
        tmpsql = "select * from MMSContract where ContractNo='" + Me.ContractNoTextBox.Text + "'"
        If Me.get_DataTable(tmpsql).Rows.Count > 0 Then
            CheckContractNo += "合約編號已經存在!! \n"
        End If
        '
    End Function

    '新增
    Protected Sub ButtonSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        'countdate()
        Dim sMsg As String = ""
        Dim getMonSE As String = ""
        Dim tmpsql As String
        'ADD IN 20141103
        Dim sSql As String
        Dim sRow As Collection = New Collection

        sMsg = CheckData_detail()

        If sMsg.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & sMsg)
            Exit Sub
        End If

        '20141127 增加合約編號檢查 (新增才檢查)
        If CheckBox1.Checked Then
            'modDB.InsertSignRecord("azitest", "contractno check : " + Mid(StartDateCTextBox.Text.Trim, 3, 2) + " VS " + Mid(ContractNoTextBox.Text.Trim, 10, 2), My.User.Name)
            If Mid(StartDateCTextBox.Text.Trim, 3, 2) <> Mid(ContractNoTextBox.Text.Trim, 10, 2) Then '合約編號年度不同
                modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "合約編號年度 與 合約日期 不同! \n")
                Exit Sub
            End If
        End If

        Dim cno, ccheck As String
        If ContractNoTextBox.Text.Trim = "" Then
            cno = getContractNo("")
        Else
            ccheck = CheckContractNo.Trim
            If ccheck <> "" Then
                modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & ccheck)
                Exit Sub
            Else
                cno = ContractNoTextBox.Text
            End If
        End If
        tmpsql = "INSERT INTO [SMILE_HQ].[dbo].[MMSContract]([ContractNo],[CustomerNo],[CustomerName],[TelNo],[Contact],[Salesman]"
        tmpsql += " ,[Cashier],[AreaCode],[OldContractNo],[ContractType],[BuildingName],[Address],[SpecificationNo],[Quantity]"
        tmpsql += " ,[StartDateC],[EndDateC],[TotalOfMonths],[UnitPricePerMonth],[TotalQuantity],[PricePerMonth],[AmountOfContract]"
        tmpsql += " ,[PeriodMaintenanceAmount],[MaintenanceAmount],[AdjAmountOfContract],[PaymentMethod],[DaysAllowed],[GUINumber]"
        tmpsql += " ,[TitleOfInvoice],[AddressOfInvoice],[StartDateI],[EndDateI],[InvoiceCycle],[ItemName],[DatePlus],[Payer]"
        tmpsql += " ,[PayerTel],[PayerAdderss],[ItemNo],[SealDate],[ArchiveDateM],[ArchiveDateA],[CancelDate],[Memo],[Memo2])"
        tmpsql += " VALUES("
        tmpsql += "'" + cno + "'," '[ContractNo]
        tmpsql += "'" + CustomerNoTextBox.Text + "'," '[CustomerNo]
        tmpsql += "'" + CustomerNameTextBox.Text + "'," '[CustomerName]
        tmpsql += "'" + TelNoTextBox.Text + "'," '[TelNo]
        tmpsql += "'" + ContactTextBox.Text + "'," '[Contact]
        tmpsql += "'" + DropDownList6.SelectedValue + "'," '[Salesman]
        tmpsql += "'" + DropDownList7.SelectedValue + "'," '[Cashier]
        tmpsql += "'" + DropDownList5.SelectedValue + "'," '[AreaCode]
        tmpsql += "'" + OldContractNoTextBox.Text + "'," '[OldContractNo]
        tmpsql += "'" + RadioButtonList2.SelectedValue + "'," '[ContractType]
        tmpsql += "'" + BuildingNameTextBox.Text + "'," '[BuildingName]
        tmpsql += "'" + AddressTextBox.Text + "'," '[Address]
        tmpsql += "'" + SpecificationNoTextBox.Text + "'," '[SpecificationNo]
        tmpsql += "'" + QuantityTextBox.Text + "'," '[Quantity]
        tmpsql += "'" + StartDateCTextBox.Text + "'," '[StartDateC]
        tmpsql += "'" + EndDateCTextBox.Text + "'," '[EndDateC]
        tmpsql += "'" + TotalOfMonthsTextBox.Text + "'," '[TotalOfMonths]
        tmpsql += "'" + UnitPricePerMonthTextBox.Text + "'," '[UnitPricePerMonth]
        tmpsql += "'" + TotalQuantityTextBox.Text + "'," '[TotalQuantity]
        tmpsql += "'" + PricePerMonthTextBox.Text + "'," '[PricePerMonth]
        tmpsql += "'" + AmountOfContractTextBox.Text + "'," '[AmountOfContract]
        tmpsql += "'" + PeriodMaintenanceAmountTextBox.Text + "'," '[PeriodMaintenanceAmount]
        tmpsql += "'" + MaintenanceAmountTextBox.Text + "'," '[MaintenanceAmount]
        tmpsql += "'" + AdjAmountOfContractTextBox.Text + "'," '[AdjAmountOfContract]
        tmpsql += "'" + DropDownList8.SelectedValue + "'," '[PaymentMethod]
        tmpsql += "'" + DaysAllowedTextBox.Text + "'," '[DaysAllowed]
        tmpsql += "'" + GUINumberTextBox.Text + "'," '[GUINumber]
        tmpsql += "'" + TitleOfInvoiceTextBox.Text + "'," '[TitleOfInvoice]
        tmpsql += "'" + AddressOfInvoiceTextBox.Text + "'," '[AddressOfInvoice]
        tmpsql += "'" + StartDateITextBox.Text + "'," '[StartDateI]
        tmpsql += "'" + EndDateITextBox.Text + "'," '[EndDateI]
        tmpsql += "'" + DropDownList4.SelectedValue + "'," '[InvoiceCycle]
        tmpsql += "'" + ItemNameTextBox.Text + "'," '[ItemName]
        tmpsql += "'" + RadioButtonList3.SelectedValue + "'," '[DatePlus]
        tmpsql += "'" + PayerTextBox.Text + "'," '[Payer]
        tmpsql += "'" + PayerTelTextBox.Text + "'," '[PayerTel]
        tmpsql += "'" + PayerAdderssTextBox.Text + "'," '[PayerAdderss]
        tmpsql += "'" + ItemNoTextBox.Text + "'," '[ItemNo]
        tmpsql += "'" + SealDateTextBox.Text + "'," '[SealDate]
        tmpsql += "'" + ArchiveDateMTextBox.Text + "'," '[ArchiveDateM]
        tmpsql += "'" + ArchiveDateATextBox.Text + "'," '[ArchiveDateA]
        tmpsql += "'" + CancelDateTextBox.Text + "'," '[CancelDate]
        tmpsql += "'" + MemoTextBox.Text + "'," '[Memo])
        tmpsql += "'" + Memo2TextBox.Text + "'" '[Memo2]) 'add in 20160201
        tmpsql += ")"
        Me.SqlDataSource1.InsertCommand = tmpsql
        Me.SqlDataSource1.Insert()

        tmpsql = "delete from MMSInvoiceA where ContractNo='" + cno + "' AND (YN<>'Y') AND (INVOICENO='') "
        Me.SqlDataSource1.DeleteCommand = tmpsql
        Me.SqlDataSource1.Delete()
        Dim tcycle, cyc1Y, cyc1M, cyc2, cyct As Integer

        Select Case DropDownList4.SelectedValue.ToString
            Case "1"
                tcycle = 12
            Case "2"
                tcycle = 3
            Case "3"
                tcycle = 6
            Case "4"
                tcycle = 2
            Case "5"
                tcycle = 1
            Case "6"
                tcycle = 24
        End Select
        '
        'modDB.InsertSignRecord("azitest", "Contract001_test: tcycle=" + tcycle.ToString, My.User.Name)
        '
        Dim cyc1 As Date
        cyc1 = StartDateITextBox.Text
        '
        cyc1Y = cyc1.Year
        cyc1M = cyc1.Month
        cyc2 = Mid(EndDateITextBox.Text, 1, 7).Replace("/", "") 'Int(Mid(GET_W_DATE(EndDateITextBox.Text), 1, 4).ToString + Mid(GET_W_DATE(EndDateITextBox.Text), 6, 2))
        cyct = Mid(StartDateITextBox.Text, 1, 7).Replace("/", "") 'Int(Mid(GET_W_DATE(StartDateITextBox.Text), 1, 4).ToString + Mid(GET_W_DATE(StartDateITextBox.Text), 6, 2))

        Do While cyct <= cyc2
            sSql = "select * from [MMSInvoiceA] where [CONTRACTNO]='" + cno + "' AND [InvoicePeriod]='" + cyct.ToString + "'"
            sRow = modDB.GetRowData(sSql)
            If sRow.Count = 0 Then
                tmpsql = "insert into MMSInvoiceA(ContractNo,InvoicePeriod,YN,INVOICEDATE,INVOICENO,updatetime) values('" + cno + "','" + cyct.ToString + "','N','','',getdate())"
                Me.SqlDataSource1.InsertCommand = tmpsql
                Me.SqlDataSource1.Insert()
            End If
            '
            If cyc1M + tcycle > 12 Then
                cyc1Y = cyc1Y + ((tcycle + cyc1M) \ 12)
                cyc1M = cyc1M + tcycle - ((tcycle + cyc1M) \ 12) * 12
            Else
                cyc1M += tcycle
            End If
            '
            If cyc1M < 10 Then
                cyct = Int(cyc1Y.ToString.Trim + "0" + cyc1M.ToString.Trim)
            Else
                cyct = Int(cyc1Y.ToString.Trim + cyc1M.ToString.Trim)
            End If
        Loop

        tmpsql = "update MMSCustomers set Used='Y' where CustomerNo='" + CustomerNoTextBox.Text + "'"
        EXE_SQL(tmpsql)

        Dim sFields As New Hashtable
        sFields.Add("txtContractID", cno)
        sFields.Add("txtContractID0", cno)
        Me.ViewState("QryField") = sFields
        Session("QryField") = Me.ViewState("QryField")

        Response.Redirect("Contract.aspx?Returnflag=1&msg=資料新增成功!!")
    End Sub


    '取得西元年月
    Public Function GET_W_DATE(ByVal dd As String) As String
        Dim tdd As Date
        GET_W_DATE = LTrim(RTrim(tdd.Year.ToString))
        If tdd.Month < 10 Then GET_W_DATE += "0"
        GET_W_DATE += LTrim(RTrim(tdd.Month.ToString))
        If tdd.Day < 10 Then GET_W_DATE += "0"
        GET_W_DATE += LTrim(RTrim(tdd.Month.ToString))
    End Function

    '修改
    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim tmpsql As String
        tmpsql = " UPDATE [SMILE_HQ].[dbo].[MMSContract]"
        tmpsql += " SET [CustomerNo] = '" + CustomerNoTextBox.Text + "'"
        tmpsql += " ,[CustomerName] = '" + CustomerNameTextBox.Text + "'"
        tmpsql += " ,[TelNo] = '" + TelNoTextBox.Text + "'"
        tmpsql += " ,[Contact] = '" + ContactTextBox.Text + "'"
        tmpsql += " ,[Salesman] = '" + DropDownList6.SelectedValue + "'"
        tmpsql += " ,[Cashier] = '" + DropDownList7.SelectedValue + "'"
        tmpsql += " ,[AreaCode] = '" + DropDownList5.SelectedValue + "'"
        tmpsql += " ,[OldContractNo] = '" + OldContractNoTextBox.Text + "'"
        tmpsql += " ,[ContractType] = '" + RadioButtonList2.SelectedValue + "'"
        tmpsql += " ,[BuildingName] = '" + BuildingNameTextBox.Text + "'"
        tmpsql += " ,[Address] = '" + AddressTextBox.Text + "'"
        tmpsql += " ,[SpecificationNo] = '" + SpecificationNoTextBox.Text + "'"
        tmpsql += " ,[Quantity] = '" + QuantityTextBox.Text + "'"
        tmpsql += " ,[StartDateC] = '" + StartDateCTextBox.Text + "'"
        tmpsql += " ,[EndDateC] = '" + EndDateCTextBox.Text + "'"
        tmpsql += " ,[TotalOfMonths] = '" + TotalOfMonthsTextBox.Text + "'"
        tmpsql += " ,[UnitPricePerMonth] = '" + UnitPricePerMonthTextBox.Text + "'"
        tmpsql += " ,[TotalQuantity] = '" + TotalQuantityTextBox.Text + "'"
        tmpsql += " ,[PricePerMonth] = '" + PricePerMonthTextBox.Text + "'"
        tmpsql += " ,[AmountOfContract] = '" + AmountOfContractTextBox.Text + "'"
        tmpsql += " ,[PeriodMaintenanceAmount] = '" + PeriodMaintenanceAmountTextBox.Text + "'"
        tmpsql += " ,[MaintenanceAmount] = '" + MaintenanceAmountTextBox.Text + "'"
        tmpsql += " ,[AdjAmountOfContract] = '" + AdjAmountOfContractTextBox.Text + "'"
        tmpsql += " ,[PaymentMethod] = '" + DropDownList8.SelectedValue + "'"
        tmpsql += " ,[DaysAllowed] = '" + DaysAllowedTextBox.Text + "'"
        tmpsql += " ,[GUINumber] = '" + GUINumberTextBox.Text + "'"
        tmpsql += " ,[TitleOfInvoice] = '" + TitleOfInvoiceTextBox.Text + "'"
        tmpsql += " ,[AddressOfInvoice] = '" + AddressOfInvoiceTextBox.Text + "'"
        tmpsql += " ,[StartDateI] = '" + StartDateITextBox.Text + "'"
        tmpsql += " ,[EndDateI] = '" + EndDateITextBox.Text + "'"
        tmpsql += " ,[InvoiceCycle] = '" + DropDownList4.SelectedValue + "'"
        tmpsql += " ,[ItemName] = '" + ItemNameTextBox.Text + "'"
        tmpsql += " ,[DatePlus] = '" + RadioButtonList3.SelectedValue + "'"
        tmpsql += " ,[Payer] = '" + PayerTextBox.Text + "'"
        tmpsql += " ,[PayerTel] = '" + PayerTelTextBox.Text + "'"
        tmpsql += " ,[PayerAdderss] = '" + PayerAdderssTextBox.Text + "'"
        tmpsql += " ,[ItemNo] = '" + ItemNoTextBox.Text + "'"
        tmpsql += " ,[SealDate] = '" + SealDateTextBox.Text + "'"
        tmpsql += " ,[ArchiveDateM] = '" + ArchiveDateMTextBox.Text + "'"
        tmpsql += " ,[ArchiveDateA] = '" + ArchiveDateATextBox.Text + "'"
        tmpsql += " ,[CancelDate] = '" + CancelDateTextBox.Text + "'"
        tmpsql += " ,[Memo] = '" + MemoTextBox.Text + "'"
        tmpsql += " ,[Memo2] = '" + Memo2TextBox.Text + "'"
        tmpsql += "  WHERE ContractNo='" + ContractNoTextBox.Text + "'"
        Me.SqlDataSource1.UpdateCommand = tmpsql
        Me.SqlDataSource1.Update()
        '
        modDB.InsertSignRecord("(CONTRACT001)-合約修改", "合約編號:" + ContractNoTextBox.Text, My.User.Name)
        'MARK IN 20141028
        tmpsql = "Delete from MMSInvoiceA where ContractNo='" + ContractNoTextBox.Text + "'" _
               + " AND YN='N'"
        Me.SqlDataSource1.DeleteCommand = tmpsql
        Me.SqlDataSource1.Delete()
        'add in 20141124
        Dim INVYM As String = ""
        tmpsql = "Select ISNULL(Max(InvoicePeriod),'')  AS InvoicePeriod from MMSInvoiceA where ContractNo='" + ContractNoTextBox.Text + "'"
        INVYM = Trim(get_DataTable(tmpsql).Rows(0).Item("InvoicePeriod"))
        Dim vINVYM As Integer = 0
        If INVYM <> "" Then
            vINVYM = Convert.ToInt32(INVYM)
        End If
        '
        Dim tcycle, cyc1Y, cyc1M, cyc2, cyct As Integer
        Select Case DropDownList4.SelectedValue.ToString
            Case "1"
                tcycle = 12
            Case "2"
                tcycle = 3
            Case "3"
                tcycle = 6
            Case "4"
                tcycle = 2
            Case "5"
                tcycle = 1
            Case "6"
                tcycle = 24
        End Select

        Dim dd As Date = StartDateITextBox.Text
        cyc1Y = dd.Year ' Int(Mid(StartDateITextBox.Text, 1, 4))
        cyc1M = dd.Month ' Int(Mid(StartDateITextBox.Text, 6, 2))
        cyc2 = Mid(EndDateITextBox.Text, 1, 7).Replace("/", "") 'Int(Mid(GET_W_DATE(EndDateITextBox.Text), 1, 4).ToString + Mid(GET_W_DATE(EndDateITextBox.Text), 6, 2))
        cyct = Mid(StartDateITextBox.Text, 1, 7).Replace("/", "") 'Int(Mid(GET_W_DATE(StartDateITextBox.Text), 1, 4).ToString + Mid(GET_W_DATE(StartDateITextBox.Text), 6, 2))
        'modDB.InsertSignRecord("azitest", "Contract001_test: cyc2=" + cyc2.ToString + " cyc1M=" + cyc1M.ToString, My.User.Name)
        Do While cyct <= cyc2
            If (vINVYM = 0) Or ((vINVYM > 0) And (cyct > vINVYM)) Then
                tmpsql = "Insert Into MMSInvoiceA(ContractNo,InvoicePeriod,YN,INVOICEDATE,INVOICENO,updatetime) values('" + ContractNoTextBox.Text + "','" + cyct.ToString + "','N','','',Getdate())"
                Me.SqlDataSource1.InsertCommand = tmpsql
                Me.SqlDataSource1.Insert()
            End If
            'modDB.InsertSignRecord("azitest", "Contract001_test: tcycle=" + tcycle.ToString + " cyc1M=" + cyc1M.ToString, My.User.Name)

            If cyc1M + tcycle > 12 Then
                cyc1Y = cyc1Y + ((tcycle + cyc1M) \ 12)
                cyc1M = cyc1M + tcycle - ((tcycle + cyc1M) \ 12) * 12
                'modDB.InsertSignRecord("azitest", "Contract001_test: >12 :" + tcycle.ToString + " cyc1M=" + cyc1M.ToString, My.User.Name)
            Else
                cyc1M += tcycle
            End If
            If cyc1M < 10 Then
                cyct = Int(cyc1Y.ToString.Trim + "0" + cyc1M.ToString.Trim)
            Else
                cyct = Int(cyc1Y.ToString.Trim + cyc1M.ToString.Trim)
            End If
            'modDB.InsertSignRecord("azitest", "Contract001_test: cyct =" + cyct.ToString, My.User.Name)
        Loop
        dispDetail()
        Me.Button3.Text = "回主畫面"
    End Sub

    Private Function getFulldate(ByVal dd As String) As String
        Dim tmpdd As Date = dd
        getFulldate = LTrim(RTrim(tmpdd.Year.ToString)) + "/"
        If tmpdd.Month < 10 Then getFulldate += "0"
        getFulldate += LTrim(RTrim(tmpdd.Month.ToString)) + "/"
        If tmpdd.Day < 10 Then getFulldate += "0"
        getFulldate += LTrim(RTrim(tmpdd.Day.ToString))
    End Function

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.popCmp.Show()

    End Sub


    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.popCmp.Hide()
    End Sub

    '******************************************************************************************************
    '* 開始查詢
    '******************************************************************************************************
    Protected Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim sSql As String
        Me.txtQCmpID.Text = Me.txtQCmpID.Text.Trim
        Me.txtQSName.Text = Me.txtQSName.Text.Trim

        Me.lblMsg.Text = ""
        If Me.txtQCmpID.Text & Me.txtQSName.Text = "" Then
            Me.lblMsg.Text = "至少須輸入一個查詢條件！"
        ElseIf Me.txtQSName.Text = "" And Me.txtQCmpID.Text.Length < 2 Then
            Me.lblMsg.Text = "代號至少須輸入兩碼！"
        End If
        If Me.lblMsg.Text <> "" Then Me.lblMsg.Visible = True : Exit Sub
        Me.lblMsg.Visible = False

        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        sSql = "SELECT [CustomerNo],[CustomerName] FROM [MMSCustomers] where Effective='Y'"
        If Me.txtQCmpID.Text.Trim <> "" Then
            sSql += " and CustomerNo like '%" + Me.txtQCmpID.Text.Trim + "%'"
        End If
        If Me.txtQSName.Text.Trim <> "" Then
            sSql += " and CustomerName like '%" + Me.txtQSName.Text.Trim + "%'"
        End If
        Me.dscCmp.SelectCommand = sSql

        '--------------------------------------------------* 設定GridView資料來源ID
        Me.grdCmp.DataSourceID = Me.dscCmp.ID
        Me.grdCmp.DataBind()

        If Me.grdCmp.Rows.Count = 0 Then
            Me.lblMsg.Text = "查無資料！"
            Me.lblMsg.Visible = True
        Else
            Me.ViewState("SQL") = sSql
        End If
    End Sub


    'Protected Sub grdCmp_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles grdCmp.RowCommand

    '    If e.CommandName = "Select" Then
    '        Dim ss As String = ""
    '        CustomerNoTextBox.Text = grdCmp.SelectedRow.Cells(0).Text
    '    End If
    'End Sub


    'Protected Sub grdCmp_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles grdCmp.SelectedIndexChanged
    '    CustomerNoTextBox.Text = Me.grdCmp.SelectedRow.Cells(1).Text
    '    Me.popCmp.Hide()
    'End Sub

    '******************************************************************************************************
    '* 設定 回傳值Script
    '******************************************************************************************************
    Protected Sub grdCmp_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdCmp.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sBtn As Button = e.Row.FindControl("Button1")
            Dim sStr As String
            sStr = "javascript: var sObj = 'txtRet'; " _
                 & "$get(sObj).value = '" & e.Row.DataItem.Item(0) & "," & e.Row.DataItem.Item(1) & "' ;"
            sBtn.Attributes.Add("onclick", sStr)
            'DataBinder.Eval(e.Row.DataItem,  "CmpID")
            CallByName(CustomerNoTextBox, "Text", CallType.Set, e.Row.DataItem.Item(0))
            CallByName(CustomerNameTextBox, "Text", CallType.Set, e.Row.DataItem.Item(1))
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        CallByName(Me.Page, "CustomerNoTextBox_TextChanged", CallType.Method, Me, Nothing)
    End Sub

    Public Sub CustomerNoTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CustomerNoTextBox.TextChanged
        Dim sSql As String
        Dim sRow As Collection = New Collection
        sSql = "select * from [MMSCustomers] where [CustomerNo]='" + CustomerNoTextBox.Text + "'"
        sRow = modDB.GetRowData(sSql)
        If sRow.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "查無客戶資料!!")
            CustomerNameTextBox.Text = ""
            CustomerNoTextBox.Text = ""
            Exit Sub
        End If
        CustomerNameTextBox.Text = sRow.Item("CustomerName").ToString
        CustomerNoTextBox.Text = sRow.Item("CustomerNo").ToString
        Me.DropDownList5.SelectedValue = sRow.Item("AreaCode").ToString
        Me.TextBox1.Text = sRow.Item("Salesman").ToString
        Me.TextBox2.Text = sRow.Item("Cashier").ToString
        Me.dispEmployee()
        'Me.DropDownList6.Enabled = True
        'Me.DropDownList7.Enabled = True
        Me.TelNoTextBox.Text = sRow.Item("TelNo").ToString
        Me.ContactTextBox.Text = sRow.Item("Contact").ToString
        'Me.DropDownList6.SelectedValue = sRow.Item("Salesman").ToString
        'Me.DropDownList7.SelectedValue = sRow.Item("Cashier").ToString
        Me.DropDownList5.Enabled = True

        Me.GUINumberTextBox.Text = sRow.Item("GUINumber").ToString
        Me.TitleOfInvoiceTextBox.Text = sRow.Item("TitleOfInvoice").ToString
        Me.AddressOfInvoiceTextBox.Text = sRow.Item("Address").ToString
        Me.PayerAdderssTextBox.Text = sRow.Item("Address").ToString
        '
        If (Request("FormMode") = "add") And (ContractNoTextBox.Text.Trim = "") And (StartDateCTextBox.Text <> "") Then
            Me.ContractNoTextBox.Text = getContractNo("")
        End If
    End Sub



    ''ContractNo取得
    'Private Function getContractNo() As String
    '    Dim seg1, seg2, seg3 As String
    '    seg1 = "SE" + Me.DropDownList5.SelectedValue + "A"
    '    seg2 = Mid(Now.Year.ToString, 3, 2)
    '    seg3 = 1

    '    Dim sSql As String
    '    Dim sRow As Collection = New Collection
    '    sSql = "select isnull(max(ContractNo),'') as ContractNo from [MMSContract] where substring([ContractNo],1,4)='" + seg1 + "'"
    '    sSql += " and substring([ContractNo],9,2)='" + Mid(Now.Year.ToString, 3, 2) + "'"
    '    sRow = modDB.GetRowData(sSql)
    '    If sRow.Item("ContractNo").ToString = "" Then
    '        seg1 += "001"
    '    Else
    '        Select Case Str(Int(Mid(sRow.Item("ContractNo").ToString, 5, 3)) + 1).Trim.Length
    '            Case 1
    '                seg1 += "00"
    '            Case 2
    '                seg1 += "0"
    '        End Select
    '        seg1 += Str(Int(Mid(sRow.Item("ContractNo").ToString, 5, 3)) + 1).Trim
    '    End If

    '    getContractNo = seg1 + "-" + seg2 + "-" + seg3
    'End Function

    'ContractNo取得
    Private Function getContractNo(ByVal tt As String) As String
        If (Me.ContractNoTextBox.Text = "") And (Me.StartDateCTextBox.Text <> "") Then
            Dim seg1, seg2, seg3 As String
            If tt = "old" Then
                seg1 = Mid(OldContractNoTextBox.Text, 1, 2) + CustomerNoTextBox.Text
                seg2 = Mid(OldContractNoTextBox.Text, 10, 2)
            Else
                seg1 = "SE" + CustomerNoTextBox.Text
                '20141124 改抓合約日期
                'seg2 = Mid(Now.Year.ToString, 3, 2)
                seg2 = Mid(StartDateCTextBox.Text, 3, 2)
            End If
            seg3 = ""

            Dim sSql As String
            Dim sRow As Collection = New Collection
            sSql = "select isnull(max(ContractNo),'') as ContractNo from [MMSContract] where substring([ContractNo],1,8)='" + seg1 + "'"
            'sSql += " and substring([ContractNo],10,2)='" + Mid(Now.Year.ToString, 3, 2) + "'"
            sSql += " and substring([ContractNo],10,2)='" + seg2 + "'"
            sRow = modDB.GetRowData(sSql)
            If sRow.Item("ContractNo").ToString = "" Then
                seg3 += "01"
            Else
                If (Int(Mid(sRow.Item("ContractNo").ToString, 13, 2)) + 1) < 10 Then
                    seg3 = "0"
                End If
                seg3 += Str(Int(Mid(sRow.Item("ContractNo").ToString, 13, 2)) + 1).Trim
            End If

            getContractNo = seg1 + "-" + seg2 + "-" + seg3
        Else
            getContractNo = ""
        End If

    End Function

    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.TitleLabel.Text = "合約基本資料 - 新增"
        Me.ButtonSave.Visible = True
        Me.Button2.Visible = False
        Me.Button5.Visible = False
        Me.OldContractNoTextBox.Text = Me.ContractNoTextBox.Text
        Me.ContractNoTextBox.Text = ""
        CheckBox1.Checked = True
        CheckBox1.Visible = True
        EditingMode = "add"

        Dim tmpsql, sMsg As String
        sMsg = ""
        tmpsql = "select * from MMSArea where AreaCode='" + Me.DropDownList5.SelectedValue + "'"
        If get_DataTable(tmpsql).Rows(0).Item("Effective") = "N" Then
            sMsg += "此客戶所屬區域已經停用! \n"
        End If
        tmpsql = "select * from MMSEmployee where EmployeeNo='" + Me.TextBox1.Text + "'"
        tmpsql += " and Effective='Y' and Salesman='Y'"
        tmpsql += " and AreaCode='" + Me.DropDownList5.SelectedValue + "'"
        If get_DataTable(tmpsql).Rows.Count = 0 Then
            sMsg += "此客戶所選業務員已經不在此區域! \n"
            Me.TextBox1.Text = ""
        End If
        tmpsql = "select * from MMSEmployee where EmployeeNo='" + Me.TextBox2.Text + "'"
        tmpsql += " and Effective='Y' and Cashier='Y'"
        tmpsql += " and AreaCode='" + Me.DropDownList5.SelectedValue + "'"
        If get_DataTable(tmpsql).Rows.Count = 0 Then
            sMsg += "此客戶所選收款員已經不在此區域! \n"
            Me.TextBox2.Text = ""
        End If
        If sMsg.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", sMsg + "請修正上述問題後點選更新!")
            tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
            Me.SqlDataSource2.SelectCommand = tmpsql
            Me.SqlDataSource2.DataBind()
            Me.dispEmployee()
        End If

        EnableAll1()

        Me.ArchiveDateATextBox.Text = ""
        Me.ArchiveDateMTextBox.Text = ""
        Me.CancelDateTextBox.Text = ""
        Me.SealDateTextBox.Text = ""

        Me.ContractNoTextBox.Text = getContractNo("old")
        Me.ContractNoTextBox.ReadOnly = False
        modUtil.SetObjEnabled(ContractNoTextBox)

        Me.Button3.Text = "取消設定"
        Me.Button7.Visible = False
    End Sub

    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        'StartDateCTextBox
        'EndDateCTextBox
        If Me.StartDateCTextBox.Text = "" Or Me.EndDateCTextBox.Text = "" Then
            Me.TotalOfMonthsTextBox.Text = "0"
            Exit Sub
        End If
        'Dim y1, y2, m1, m2 As Integer
        'y1 = Mid(StartDateCTextBox.Text, 1, 4)
        'm1 = Mid(StartDateCTextBox.Text, 6, 2)
        'y2 = Mid(EndDateCTextBox.Text, 1, 4)
        'm2 = Mid(EndDateCTextBox.Text, 6, 2)
        ''Me.TotalOfMonthsTextBox.Text = ((y2 - y1) * 12) + (m2 - m1 + 1)
        'Me.TotalOfMonthsTextBox.Text = ((y2 - y1) * 12) + (m2 - m1)

        Dim dstart, dend As Date
        dstart = StartDateCTextBox.Text
        dend = EndDateCTextBox.Text
        Me.TotalOfMonthsTextBox.Text = DateDiff(DateInterval.Month, dstart, dend)
        If dend.Day > dstart.Day Then
            Me.TotalOfMonthsTextBox.Text = Int(Me.TotalOfMonthsTextBox.Text) + 1
        End If
        '
        If (EditingMode = "add") And (ContractNoTextBox.Text.Trim = "") And (StartDateCTextBox.Text <> "") Then
            Me.ContractNoTextBox.Text = getContractNo("")
        End If
    End Sub


#Region "DB function"

    '設定區域資料已使用
    Public Function SetEmployeeUsed(ByVal ss As String)
        Dim tmpsql As String
        tmpsql = "update MMSEmployee set Used='Y' where EmployeeNo='" + ss + "'"
        EXE_SQL(tmpsql)
    End Function

    '設定區域資料已使用
    Public Function SetAreaUsed(ByVal ss As String)
        Dim tmpsql As String
        tmpsql = "update MMSArea set Used='Y' where AreaCode='" + ss + "'"
        EXE_SQL(tmpsql)
    End Function

    '連接資料庫取得datatable
    Public Function get_DataTable(ByVal SQL1 As String) As System.Data.DataTable
        Dim CONNSTR1 As String = modDB.GetDataSource.ConnectionString
        Dim CONN1 As New System.Data.SqlClient.SqlConnection(CONNSTR1)
        Dim DATAADAPTER1 As System.Data.SqlClient.SqlDataAdapter
        Try
            DATAADAPTER1 = New System.Data.SqlClient.SqlDataAdapter(SQL1, CONN1)
            DATAADAPTER1.SelectCommand.CommandTimeout = 1600
            Dim MYDATASET1 As New System.Data.DataSet
            MYDATASET1.Reset()
            DATAADAPTER1.Fill(MYDATASET1)
            get_DataTable = MYDATASET1.Tables(0)
        Catch ex As Exception
            get_DataTable = Nothing
        End Try
    End Function

    '執行quary不回傳值
    Public Function EXE_SQL(ByVal TMPSQL1 As String) As Boolean
        Dim AA As Integer
        Dim CONNSTR1 As String = modDB.GetDataSource.ConnectionString
        Dim CONN1 As New System.Data.SqlClient.SqlConnection(CONNSTR1)
        Dim SQLCOMMAND1 As New System.Data.SqlClient.SqlCommand
        Try
            SQLCOMMAND1.CommandText = TMPSQL1
            If CONN1.State.Closed = ConnectionState.Open Then CONN1.Close()
            SQLCOMMAND1.Connection = CONN1
            CONN1.Open()
            SQLCOMMAND1.ExecuteNonQuery()
            CONN1.Close()
        Catch ex As Exception
            EXE_SQL = False
            Exit Function
        End Try
        EXE_SQL = True
    End Function
#End Region



    Protected Sub DropDownList6_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList6.DataBound
        Try
            Me.DropDownList6.SelectedValue = Me.TextBox1.Text
        Catch ex As Exception

        End Try

    End Sub

    Protected Sub DropDownList7_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList7.DataBound
        Try
            Me.DropDownList7.SelectedValue = Me.TextBox2.Text
        Catch ex As Exception

        End Try

    End Sub

    Protected Sub GUINumberTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GUINumberTextBox.TextChanged

        If modUtil.BANCheck(GUINumberTextBox.Text) = False And GUINumberTextBox.Text.Trim <> "" Then
            modUtil.showMsg(Me.Page, "訊息", "統一編號檢查錯誤! \n")
        End If
    End Sub

    Protected Sub StartDateCTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles StartDateCTextBox.TextChanged
        'countdate()
        If (EditingMode = "add") And (ContractNoTextBox.Text.Trim = "") And (StartDateCTextBox.Text <> "") Then
            Me.ContractNoTextBox.Text = getContractNo("")
        End If
    End Sub

    Private Sub countdate()
        'StartDateCTextBox
        'EndDateCTextBox
        If Me.StartDateCTextBox.Text = "" Or Me.EndDateCTextBox.Text = "" Then
            Me.TotalOfMonthsTextBox.Text = "0"
            Exit Sub
        End If

        Dim dstart, dend As Date
        dstart = StartDateCTextBox.Text
        dend = EndDateCTextBox.Text
        Me.TotalOfMonthsTextBox.Text = DateDiff(DateInterval.Month, dstart, dend)
        If dend.Day > dstart.Day Then
            Me.TotalOfMonthsTextBox.Text = Int(Me.TotalOfMonthsTextBox.Text) + 1
        End If
    End Sub

    Protected Sub EndDateCTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles EndDateCTextBox.TextChanged, StartDateCTextBox.TextChanged
        countdate()
    End Sub

    Protected Sub PricePerMonthTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PricePerMonthTextBox.TextChanged
        Dim i As Integer
        'If Me.UnitPricePerMonthTextBox.Text.Trim <> "" Or Me.UnitPricePerMonthTextBox.Text <> "0" Then Exit Sub
        If TotalQuantityTextBox.Text.Trim = "" Or TotalQuantityTextBox.Text = "0" Then Exit Sub
        If Me.StartDateCTextBox.Text <> "" And Me.EndDateCTextBox.Text <> "" Then
            countdate()
        Else
            Exit Sub
        End If
        UnitPricePerMonthTextBox.Text = Int(Int(PricePerMonthTextBox.Text) / Int(TotalQuantityTextBox.Text))
        AmountOfContractTextBox.Text = Int(PricePerMonthTextBox.Text) * Int(TotalOfMonthsTextBox.Text)
    End Sub

    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        '
        EnableAll1()
        modUtil.SetObjReadOnly(Me, "ContractNoTextBox")
        Me.TitleLabel.Text = "合約基本資料 - 修改"
        EditingMode = ""
        Dim tmpsql As String
        tmpsql = "select * from  MMSContract where ContractNo='" + Request("STNID").Trim + "'"
        Me.SqlDataSource1.SelectCommand = tmpsql
        Me.SqlDataSource1.DataBind()
        Me.ButtonSave.Visible = False
        Me.Button2.Visible = True

        Me.Button4.Visible = False

        modUtil.SetObjReadOnly(Me, "CustomerNoTextBox") '不可修改
        DropDownList5.Enabled = True


        tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
        Me.SqlDataSource2.SelectCommand = tmpsql
        Me.SqlDataSource2.DataBind()
        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee" ' WHERE Effective = 'Y' "
        Me.SqlDataSource3.SelectCommand = tmpsql
        Me.SqlDataSource3.DataBind()
        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee" ' WHERE Effective = 'Y' "
        Me.SqlDataSource4.SelectCommand = tmpsql
        Me.SqlDataSource4.DataBind()

        Dim dv As Data.DataView
        dv = Me.SqlDataSource1.Select(New DataSourceSelectArguments())
        '第一筆第一個欄位得資料 
        'Me.ContractNoTextBox.Text = dv(0)("CustomerNo").ToString()
        Try
            ContractNoTextBox.Text = dv(0)("ContractNo").ToString()
            CustomerNoTextBox.Text = dv(0)("CustomerNo").ToString()
            CustomerNameTextBox.Text = dv(0)("CustomerName").ToString()
            TelNoTextBox.Text = dv(0)("TelNo").ToString()
            ContactTextBox.Text = dv(0)("Contact").ToString()
            DropDownList5.SelectedValue = dv(0)("AreaCode").ToString()
            Me.DropDownList6.SelectedValue = dv(0)("Salesman").ToString()
            Me.DropDownList7.SelectedValue = dv(0)("Cashier").ToString()
            Me.TextBox1.Text = dv(0)("Salesman").ToString()
            Me.TextBox2.Text = dv(0)("Cashier").ToString()
            OldContractNoTextBox.Text = dv(0)("OldContractNo").ToString()
            RadioButtonList2.SelectedValue = dv(0)("ContractType").ToString()
            BuildingNameTextBox.Text = dv(0)("BuildingName").ToString()
            AddressTextBox.Text = dv(0)("Address").ToString()
            SpecificationNoTextBox.Text = dv(0)("SpecificationNo").ToString()
            QuantityTextBox.Text = dv(0)("Quantity").ToString()
            StartDateCTextBox.Text = dv(0)("StartDateC").ToString()
            EndDateCTextBox.Text = dv(0)("EndDateC").ToString()
            TotalOfMonthsTextBox.Text = dv(0)("TotalOfMonths").ToString()
            UnitPricePerMonthTextBox.Text = dv(0)("UnitPricePerMonth").ToString()
            TotalQuantityTextBox.Text = dv(0)("TotalQuantity").ToString()
            PricePerMonthTextBox.Text = dv(0)("PricePerMonth").ToString()
            AmountOfContractTextBox.Text = dv(0)("AmountOfContract").ToString()
            PeriodMaintenanceAmountTextBox.Text = dv(0)("PeriodMaintenanceAmount").ToString()
            MaintenanceAmountTextBox.Text = dv(0)("MaintenanceAmount").ToString()
            AdjAmountOfContractTextBox.Text = dv(0)("AdjAmountOfContract").ToString()
            DropDownList8.SelectedValue = dv(0)("PaymentMethod").ToString()
            DaysAllowedTextBox.Text = dv(0)("DaysAllowed").ToString()
            GUINumberTextBox.Text = dv(0)("GUINumber").ToString()
            TitleOfInvoiceTextBox.Text = dv(0)("TitleOfInvoice").ToString()
            AddressOfInvoiceTextBox.Text = dv(0)("AddressOfInvoice").ToString()
            StartDateITextBox.Text = dv(0)("StartDateI").ToString()
            EndDateITextBox.Text = dv(0)("EndDateI").ToString()
            DropDownList4.SelectedValue = dv(0)("InvoiceCycle").ToString()
            ItemNameTextBox.Text = dv(0)("ItemName").ToString()
            RadioButtonList3.SelectedValue = dv(0)("DatePlus").ToString()
            PayerTextBox.Text = dv(0)("Payer").ToString()
            PayerTelTextBox.Text = dv(0)("PayerTel").ToString()
            PayerAdderssTextBox.Text = dv(0)("PayerAdderss").ToString()
            ItemNoTextBox.Text = dv(0)("ItemNo").ToString()
            SealDateTextBox.Text = dv(0)("SealDate").ToString()
            ArchiveDateMTextBox.Text = dv(0)("ArchiveDateM").ToString()
            ArchiveDateATextBox.Text = dv(0)("ArchiveDateA").ToString()
            CancelDateTextBox.Text = dv(0)("CancelDate").ToString()
            MemoTextBox.Text = dv(0)("Memo").ToString()
            Memo2TextBox.Text = dv(0)("Memo2").ToString()
        Catch ex As Exception

        End Try
        If dv(0)("ArchiveDateM").ToString() <> "" And dv(0)("ArchiveDateA").ToString() <> "" Then
                DisableAll()
        Else
            If dv(0)("LastInvoiceDate").ToString() <> "" Then
                DisableAll1()
            End If
        End If
        'ADD IN 20151020
        'modDB.InsertSignRecord("AziTest", "CONTRACT001_ROL=" + Me.ViewState("ROL"), My.User.Name) 'AZITEST
        If (Mid(Me.ViewState("ROL"), 5, 2) = "01") Or (Mid(Me.ViewState("ROL"), 5, 2) = "05") Then 'MODIFY IN 20151020
            EnableAll1()
        End If
        '
        If CancelDateTextBox.Text <> "" Then
            Button2.Visible = False
        End If
        TelNoTextBox.Focus()
        Me.Button7.Visible = False
        Me.Button2.Visible = True
        Me.Button3.Text = "取消設定"

        If GetAreaCode() = "" Then
            Me.Image6.Visible = True
            modUtil.SetObjEnabled(ArchiveDateATextBox)
        Else
            Me.Image6.Visible = False
            modUtil.SetObjReadOnly(Me, "ArchiveDateATextBox")
        End If

        modUtil.SetDateObj(Me.StartDateCTextBox, False, Me.EndDateCTextBox, False)
        modUtil.SetDateObj(Me.EndDateCTextBox, False, Nothing, False)
        modUtil.SetDateObj(Me.StartDateITextBox, False, Me.EndDateITextBox, False)
        modUtil.SetDateObj(Me.EndDateITextBox, False, Nothing, False)
        modUtil.SetDateObj(Me.ArchiveDateMTextBox, False, Nothing, False)
        modUtil.SetDateObj(Me.ArchiveDateATextBox, False, Nothing, False)
        modUtil.SetDateObj(Me.CancelDateTextBox, False, Nothing, False)
        modUtil.SetDateObj(Me.SealDateTextBox, False, Nothing, False)
    End Sub

    Protected Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.Click
        Response.Redirect("Contract_001.aspx?FormMode=add")
    End Sub
End Class
