Imports System.Data

Partial Class MMS_Cust_001
    Inherits System.Web.UI.Page

    Private Sub disable_edit()
        CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = False
        CType(FormView1.FindControl("DropDownList2"), DropDownList).Enabled = False
        CType(FormView1.FindControl("DropDownList3"), DropDownList).Enabled = False
        CType(FormView1.FindControl("CustomerNameTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("AddressTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("EMAILTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("TelNoTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("ContactTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("GUINumberTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("Date1TextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("TitleOfInvoiceTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("MemoTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("FaxNoTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("Date2TextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("DaysAllowedTextBox"), TextBox).ReadOnly = True
        'CType(FormView1.FindControl("Button5"), Button).Visible = False
        CType(FormView1.FindControl("Button3"), Button).Visible = False
    End Sub


    Private Sub enable_edit()
        CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = True
        CType(FormView1.FindControl("DropDownList2"), DropDownList).Enabled = True
        CType(FormView1.FindControl("DropDownList3"), DropDownList).Enabled = True
        CType(FormView1.FindControl("CustomerNameTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("AddressTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("EMAILTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("TelNoTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("ContactTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("GUINumberTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("Date1TextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("TitleOfInvoiceTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("MemoTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("FaxNoTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("Date2TextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("DaysAllowedTextBox"), TextBox).ReadOnly = False
        'CType(FormView1.FindControl("Button5"), Button).Visible = True
        CType(FormView1.FindControl("Button3"), Button).Visible = True
    End Sub

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else
                Dim sFileName As String = Request.Path.Substring(Request.Path.LastIndexOf("/") + 1, Request.Path.Length - (Request.Path.LastIndexOf("/") + 1))
                Dim sRol As String = HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
                ViewState("ROL") = modUtil.GetRolData(sFileName, sRol)
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
            End If

            Select Case Request("FormMode")
                Case "add"
                    Me.FormView1.ChangeMode(FormViewMode.Insert)
                    Me.TitleLabel.Text = "客戶基本資料 - 新增"
                    dispEmployee()
                    CType(FormView1.FindControl("CustomerNoTextBox"), TextBox).Focus()
                Case "edit"
                    Me.FormView1.ChangeMode(FormViewMode.Edit)
                    Me.TitleLabel.Text = "客戶基本資料 - 修改"
                    Dim tmpsql As String
                    tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
                    Me.SqlDataSource2.SelectCommand = tmpsql
                    Me.SqlDataSource2.DataBind()
                    tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee" ' WHERE Effective = 'Y' "
                    Me.SqlDataSource3.SelectCommand = tmpsql
                    Me.SqlDataSource3.DataBind()
                    tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee" ' WHERE Effective = 'Y' "
                    Me.SqlDataSource4.SelectCommand = tmpsql
                    Me.SqlDataSource4.DataBind()
                    tmpsql = "select * from  MMSCustomers where [CustomerNo]='" + Request("STNID").Trim + "'"
                    Me.SqlDataSource1.SelectCommand = tmpsql
                    Me.SqlDataSource1.DataBind()
                    CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue = CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Text
                    CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue = CType(FormView1.FindControl("SalesmanTextBox"), TextBox).Text
                    CType(FormView1.FindControl("DropDownList2"), DropDownList).SelectedValue = CType(FormView1.FindControl("CashierTextBox"), TextBox).Text
                    '
                    If CType(FormView1.FindControl("EffectiveTextBox"), TextBox).Text = "N" Then
                        CType(FormView1.FindControl("Button5"), Button).Text = "取消停用"
                        CType(FormView1.FindControl("Button3"), Button).Visible = False
                        CType(FormView1.FindControl("DropDownList3"), DropDownList).Enabled = False
                        CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = False
                        CType(FormView1.FindControl("DropDownList2"), DropDownList).Enabled = False
                        CType(FormView1.FindControl("CustomerNameTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("AddressTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("EMAILTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("TelNoTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("ContactTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("Date1TextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("Date2TextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("DaysAllowedTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("TitleOfInvoiceTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("MemoTextBox"), TextBox).Enabled = False
                    Else
                        tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea] WHERE Effective = 'Y' "
                        Me.SqlDataSource2.SelectCommand = tmpsql
                        Me.SqlDataSource2.DataBind()

                        Me.SqlDataSource4.DataBind()
                        CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue = CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Text
                        dispEmployee()
                    End If
                    'dispEmployee()
                    CType(FormView1.FindControl("DropDownList3"), DropDownList).Focus()

                    disable_edit()

            End Select
            Try
                CType(FormView1.FindControl("Date1TextBox"), TextBox).Attributes.Add("onkeypress", "return CheckKeyNumber()")
                CType(FormView1.FindControl("Date2TextBox"), TextBox).Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Catch ex As Exception

            End Try
            Try

                If ViewState("ROL").Substring(2, 1) = "N" Then CType(FormView1.FindControl("Button6"), Button).Visible = False
                If ViewState("ROL").Substring(3, 1) = "N" Then CType(FormView1.FindControl("Button5"), Button).Visible = False
            Catch ex As Exception

            End Try


        End If
    End Sub


    '檢查輸入資料
    Private Function CheckData_detail(ByVal ttype As String) As String
        CheckData_detail = ""
        Dim getCustomerNo As String
        If Request("FormMode") = "add" Then
            getCustomerNo = CType(FormView1.FindControl("CustomerNoTextBox"), TextBox).Text
            If getCustomerNo = "" Then
                CheckData_detail += "客戶代號不可空白! \n"
            End If
        Else
            getCustomerNo = CType(FormView1.FindControl("CustomerNoLabel1"), Label).Text
        End If

        Dim getCustomerName As String = CType(FormView1.FindControl("CustomerNameTextBox"), TextBox).Text
        Dim getGUINumber As String = CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Text
        Dim getDate1 As String = CType(FormView1.FindControl("Date1TextBox"), TextBox).Text
        Dim getDate2 As String = CType(FormView1.FindControl("Date2TextBox"), TextBox).Text
        Dim getDaysAllowed As String = CType(FormView1.FindControl("DaysAllowedTextBox"), TextBox).Text
        Dim getTitleOfInvoice As String = CType(FormView1.FindControl("TitleOfInvoiceTextBox"), TextBox).Text
        Dim getAreaCode As String = CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue
        Dim getSalesman As String = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
        Dim getCashier As String = CType(FormView1.FindControl("DropDownList2"), DropDownList).SelectedValue

        If getCustomerName = "" Then
            CheckData_detail += "客戶名稱不可空白! \n"
        End If
        If getDate1 = "" Then
            CheckData_detail += "有統編催收天數不可空白! \n"
        End If
        If getDate2 = "" Then
            CheckData_detail += "無統編催收天數不可空白! \n"
        End If
        If getDaysAllowed = "" Then
            CheckData_detail += "收款天數天數不可空白! \n"
        End If
        If getTitleOfInvoice = "" Then
            CheckData_detail += "發票抬頭不可空白! \n"
        End If
        If getAreaCode = "" Then
            CheckData_detail += "區域不可空白! \n"
        End If
        If getSalesman = "" Then
            CheckData_detail += "業務員不可空白! \n"
        End If
        If getCashier = "" Then
            CheckData_detail += "收款員不可空白! \n"
        End If
        If ttype = "add" Then
            Dim sSql As String
            Dim sRow As Collection = New Collection
            sSql = "select * from [MMSCustomers] where [CustomerNo]='" + getCustomerNo + "'"
            sRow = modDB.GetRowData(sSql)
            If sRow.Count > 0 Then
                CheckData_detail += "客戶代號已經被使用! \n"
            End If
        End If
        If modUtil.BANCheck(CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Text) = False And CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Text <> "" Then
            CheckData_detail += "統一編號檢查錯誤! \n"
        End If
    End Function

    '新增
    Protected Sub Button1_Click(sender As Object, e As System.EventArgs)
        Dim sMsg As String = ""
        Dim getMonSE As String = ""

        sMsg = CheckData_detail("add")

        If sMsg.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & sMsg)
            Exit Sub
        End If
        Dim getCustomerNo As String = CType(FormView1.FindControl("CustomerNoTextBox"), TextBox).Text
        Dim getCustomerName As String = CType(FormView1.FindControl("CustomerNameTextBox"), TextBox).Text
        Dim getAddress As String = CType(FormView1.FindControl("AddressTextBox"), TextBox).Text
        Dim getAreaCode As String = CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue
        Dim getEMAIL As String = CType(FormView1.FindControl("EMAILTextBox"), TextBox).Text
        Dim getTelNo As String = CType(FormView1.FindControl("TelNoTextBox"), TextBox).Text
        Dim getContact As String = CType(FormView1.FindControl("ContactTextBox"), TextBox).Text
        Dim getFaxNo As String = CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Text
        Dim getGUINumber As String = CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Text
        Dim getTitleOfInvoice As String = CType(FormView1.FindControl("TitleOfInvoiceTextBox"), TextBox).Text
        Dim getSalesman As String = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
        Dim getCashier As String = CType(FormView1.FindControl("DropDownList2"), DropDownList).SelectedValue
        Dim getMemo As String = CType(FormView1.FindControl("MemoTextBox"), TextBox).Text
        Dim getDate1 As String = CType(FormView1.FindControl("Date1TextBox"), TextBox).Text
        Dim getDate2 As String = CType(FormView1.FindControl("Date2TextBox"), TextBox).Text
        Dim getDaysAllowed As String = CType(FormView1.FindControl("DaysAllowedTextBox"), TextBox).Text

        Dim tmpsql As String
        '
        modDB.InsertSignRecord("MMS_CUST001", "客戶基本資料新增:" + getCustomerNo, My.User.Name) 'azitest
        '
        tmpsql = "insert into [MMSCustomers]([CustomerNo],[CustomerName],[Address],[AreaCode],[EMAIL],[TelNo],[Contact],[FaxNo],[GUINumber],[TitleOfInvoice],[Salesman],[Cashier],[Memo],[Date1],[Date2],[DaysAllowed],[Effective])"
        tmpsql += " values("
        tmpsql += "'" + getCustomerNo + "',"
        tmpsql += "'" + getCustomerName + "',"
        tmpsql += "'" + getAddress + "',"
        tmpsql += "'" + getAreaCode + "',"
        tmpsql += "'" + getEMAIL + "',"
        tmpsql += "'" + getTelNo + "',"
        tmpsql += "'" + getContact + "',"
        tmpsql += "'" + getFaxNo + "',"
        tmpsql += "'" + getGUINumber + "',"
        tmpsql += "'" + getTitleOfInvoice + "',"
        tmpsql += "'" + getSalesman + "',"
        tmpsql += "'" + getCashier + "',"
        tmpsql += "'" + getMemo + "',"
        tmpsql += getDate1 + ","
        tmpsql += getDate2 + ","
        tmpsql += getDaysAllowed + ","
        tmpsql += "'Y')"
        Me.SqlDataSource1.InsertCommand = tmpsql
        Me.SqlDataSource1.Insert()
        SetAreaUsed(getAreaCode)
        SetEmployeeUsed(getSalesman)
        SetEmployeeUsed(getCashier)

        Dim sFields As New Hashtable
        sFields.Add("TextBox1", getCustomerNo)

        sFields.Add("TextBox2", getCustomerNo)
        Me.ViewState("QryField") = sFields
        Session("QryField") = Me.ViewState("QryField")

        Response.Redirect("Cust.aspx?Returnflag=1&msg=資料新增成功!!")
    End Sub

    '修改刪除
    Protected Sub Button5_Click(sender As Object, e As System.EventArgs)
        Dim getCustomerNo As String = CType(FormView1.FindControl("CustomerNoLabel1"), Label).Text
        Dim getTextBox1 As String = CType(FormView1.FindControl("TextBox1"), TextBox).Text
        Dim tmpsql, sMsg As String
        '
        modDB.InsertSignRecord("MMS_CUST001", "客戶基本資料刪除: " + getCustomerNo, My.User.Name) 'azitest
        '
        sMsg = ""
        If CType(FormView1.FindControl("Button5"), Button).Text = "取消停用" Then
            tmpsql = "select * from MMSArea where AreaCode='" + CType(FormView1.FindControl("DropDownList3"), DropDownList).Text + "'"
            If get_DataTable(tmpsql).Rows(0).Item("Effective") = "N" Then
                sMsg += "此客戶所屬區域已經停用! \n"
            End If
            tmpsql = "select * from MMSEmployee where EmployeeNo='" + CType(FormView1.FindControl("DropDownList1"), DropDownList).Text + "'"
            tmpsql += " and Effective='Y' and Salesman='Y'"
            If get_DataTable(tmpsql).Rows.Count = 0 Then
                sMsg += "此客戶所選業務員已經停用! \n"
            End If
            tmpsql = "select * from MMSEmployee where EmployeeNo='" + CType(FormView1.FindControl("DropDownList2"), DropDownList).Text + "'"
            tmpsql += " and Effective='Y' and Cashier='Y'"
            If get_DataTable(tmpsql).Rows.Count = 0 Then
                sMsg += "此客戶所選收款員已經停用! \n"
            End If

            If sMsg.Length > 0 Then
                modUtil.showMsg(Me.Page, "訊息", sMsg + "請修正上述問題後點選更新, \n 以完成取消停用!")
                CType(FormView1.FindControl("DropDownList3"), DropDownList).Enabled = True
                CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = True
                CType(FormView1.FindControl("DropDownList2"), DropDownList).Enabled = True
                CType(FormView1.FindControl("CustomerNameTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("AddressTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("EMAILTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("TelNoTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("ContactTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("Date1TextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("Date2TextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("DaysAllowedTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("TitleOfInvoiceTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("MemoTextBox"), TextBox).Enabled = True

                tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
                Me.SqlDataSource2.SelectCommand = tmpsql
                Me.SqlDataSource2.DataBind()
                Me.dispEmployee()

                Exit Sub
            End If

            tmpsql = "update MMSCustomers set "
            tmpsql += " Effective='Y'"
            tmpsql += " where CustomerNo='" + getCustomerNo + "'"
            Me.SqlDataSource1.UpdateCommand = tmpsql
            Me.SqlDataSource1.Update()
            Response.Redirect("Cust.aspx?Returnflag=1&msg=該客戶資料已取消停用!!")
        Else
            If getTextBox1 = "Y" Then
                tmpsql = "update MMSCustomers set "
                tmpsql += " Effective='N'"
                tmpsql += " where CustomerNo='" + getCustomerNo + "'"
                Me.SqlDataSource1.UpdateCommand = tmpsql
                Me.SqlDataSource1.Update()
                Response.Redirect("Cust.aspx?Returnflag=1&msg=該客戶資料已使用過不可刪除,已將其停用!!")
            End If

            If getTextBox1 = "N" Then
                tmpsql = "delete MMSCustomers "
                tmpsql += " where CustomerNo='" + getCustomerNo + "'"
                Me.SqlDataSource1.DeleteCommand = tmpsql
                Me.SqlDataSource1.Delete()
                Response.Redirect("Cust.aspx?Returnflag=1&msg=該客戶資料已刪除!!")
            End If
        End If
    End Sub

    '修改更新
    Protected Sub Button3_Click(sender As Object, e As System.EventArgs)

        Dim sMsg As String = ""
        Dim getMonSE As String = ""

        sMsg = CheckData_detail("edit")

        If sMsg.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & sMsg)
            Exit Sub
        End If

        Dim getCustomerNo As String = CType(FormView1.FindControl("CustomerNoLabel1"), Label).Text
        Dim getCustomerName As String = CType(FormView1.FindControl("CustomerNameTextBox"), TextBox).Text
        Dim getAddress As String = CType(FormView1.FindControl("AddressTextBox"), TextBox).Text
        Dim getAreaCode As String = CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue
        Dim getEMAIL As String = CType(FormView1.FindControl("EMAILTextBox"), TextBox).Text
        Dim getTelNo As String = CType(FormView1.FindControl("TelNoTextBox"), TextBox).Text
        Dim getContact As String = CType(FormView1.FindControl("ContactTextBox"), TextBox).Text
        Dim getFaxNo As String = CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Text
        Dim getGUINumber As String = CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Text
        Dim getTitleOfInvoice As String = CType(FormView1.FindControl("TitleOfInvoiceTextBox"), TextBox).Text
        Dim getSalesman As String = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
        Dim getCashier As String = CType(FormView1.FindControl("DropDownList2"), DropDownList).SelectedValue
        Dim getMemo As String = CType(FormView1.FindControl("MemoTextBox"), TextBox).Text
        Dim getDate1 As String = CType(FormView1.FindControl("Date1TextBox"), TextBox).Text
        Dim getDate2 As String = CType(FormView1.FindControl("Date2TextBox"), TextBox).Text
        Dim getDaysAllowed As String = CType(FormView1.FindControl("DaysAllowedTextBox"), TextBox).Text
        Dim getEffective As String = CType(FormView1.FindControl("EffectiveTextBox"), TextBox).Text
        Dim tmpsql As String
        '留記錄
        modDB.InsertSignRecord("MMS_CUST001", "客戶基本資料修改:" + getCustomerNo + ",統編=" + getGUINumber + ",抬頭=" + getTitleOfInvoice, My.User.Name) 'azitest
        '
        tmpsql = "update MMSCustomers set "
        tmpsql += " [CustomerName]='" + getCustomerName + "',"
        tmpsql += " [Address]='" + getAddress + "',"
        tmpsql += " [AreaCode]='" + getAreaCode + "',"
        tmpsql += " [EMAIL]='" + getEMAIL + "',"
        tmpsql += " [TelNo]='" + getTelNo + "',"
        tmpsql += " [Contact]='" + getContact + "',"
        tmpsql += " [FaxNo]='" + getFaxNo + "',"
        tmpsql += " [GUINumber]='" + getGUINumber + "',"
        tmpsql += " [TitleOfInvoice]='" + getTitleOfInvoice + "',"
        tmpsql += " [Salesman]='" + getSalesman + "',"
        tmpsql += " [Cashier]='" + getCashier + "',"
        tmpsql += " [Memo]='" + getMemo + "',"
        tmpsql += " [Date1]=" + getDate1 + ","
        tmpsql += " [Date2]=" + getDate2 + ","
        tmpsql += " [DaysAllowed]=" + getDaysAllowed + ","
        tmpsql += " [Effective]='" + getEffective + "'"
        tmpsql += " where CustomerNo='" + getCustomerNo + "'"
        Me.SqlDataSource1.UpdateCommand = tmpsql
        Me.SqlDataSource1.Update()
        SetAreaUsed(getAreaCode)
        SetEmployeeUsed(getSalesman)
        SetEmployeeUsed(getCashier)
        '20150828 增加異動時一併修改仍然有效合約的統編與抬頭
        tmpsql = "UPDATE MMSCONTRACT SET GUINUMBER='" + getGUINumber + "',TITLEOFINVOICE='" + getTitleOfInvoice + "'" _
               + " WHERE CONTRACTNO IN (Select CONTRACTNO From MMSContract " _
               + "                       Where CUSTOMERNO='" + getCustomerNo + "'" _
               + "                         AND EnddateI>=CONVERT(VARCHAR(10),GETDATE(),111)) "
        '
        Me.SqlDataSource1.UpdateCommand = tmpsql
        Me.SqlDataSource1.Update()
        '
        Response.Redirect("Cust.aspx?Returnflag=1&msg=資料修改成功!!")
    End Sub

    '修改取消
    Protected Sub Button4_Click(sender As Object, e As System.EventArgs)
        Response.Redirect("Cust.aspx?Returnflag=1&msg=")
    End Sub

    '新增取消
    Protected Sub Button2_Click(sender As Object, e As System.EventArgs)
        Response.Redirect("Cust.aspx?Returnflag=1&msg=")
    End Sub

    Protected Sub DropDownList3_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        dispEmployee()
    End Sub

    '顯示區域所屬業務員收款員
    Private Sub dispEmployee()
        Dim tmpsql As String
        If CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue = "" And CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedIndex = -1 Then
            Try
                tmpsql = "SELECT AreaCode, AreaCode+'-'+AreaName as AreaName FROM MMSArea WHERE (Effective = 'Y')"
                Me.SqlDataSource2.SelectCommand = tmpsql
                Me.SqlDataSource2.DataBind()
                CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedIndex = 0
            Catch ex As Exception
                Dim sss As String = ""
            End Try
        End If

        Try
            CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = True
            CType(FormView1.FindControl("DropDownList2"), DropDownList).Enabled = True
        Catch ex As Exception

        End Try

        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee"
        tmpsql += " WHERE Effective = 'Y' and AreaCode='" + CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue + "'"
        tmpsql += "   And Salesman = 'Y'"
        Me.SqlDataSource3.SelectCommand = tmpsql
        Me.SqlDataSource3.DataBind()

        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee"
        tmpsql += " WHERE Effective = 'Y' and AreaCode='" + CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue + "'"
        tmpsql += "   And Cashier = 'Y'"
        Me.SqlDataSource4.SelectCommand = tmpsql
        Me.SqlDataSource4.DataBind()
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



    Protected Sub DropDownList3_SelectedIndexChanged1(sender As Object, e As System.EventArgs)
        Me.dispEmployee()
        CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Text = CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue
        CType(FormView1.FindControl("SalesmanTextBox"), TextBox).Text = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
        CType(FormView1.FindControl("CashierTextBox"), TextBox).Text = CType(FormView1.FindControl("DropDownList2"), DropDownList).SelectedValue
    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        CType(FormView1.FindControl("SalesmanTextBox"), TextBox).Text = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
    End Sub

    Protected Sub DropDownList2_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        CType(FormView1.FindControl("CashierTextBox"), TextBox).Text = CType(FormView1.FindControl("DropDownList2"), DropDownList).SelectedValue

    End Sub

    Protected Sub FormView1_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.FormViewPageEventArgs) Handles FormView1.PageIndexChanging

    End Sub

    Protected Sub CustomerNoTextBox_TextChanged(sender As Object, e As System.EventArgs)
        Dim getCustomerNo As String = CType(FormView1.FindControl("CustomerNoTextBox"), TextBox).Text
        Dim sSql As String
        Dim sRow As Collection = New Collection
        sSql = "select * from [MMSCustomers] where [CustomerNo]='" + getCustomerNo + "'"
        sRow = modDB.GetRowData(sSql)
        If sRow.Count > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "客戶代號已經被使用! \n")
        End If
    End Sub


    Protected Sub GUINumberTextBox_TextChanged(sender As Object, e As System.EventArgs)
        If modUtil.BANCheck(CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Text) = False Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "統一編號檢查錯誤! \n")
        End If
    End Sub

    Protected Sub Button6_Click(sender As Object, e As System.EventArgs)
        enable_edit()
        CType(FormView1.FindControl("Button6"), Button).Visible = False
        CType(FormView1.FindControl("Button5"), Button).Visible = False
        Dim tmpsql As String
        CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue = CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Text
        CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue = CType(FormView1.FindControl("SalesmanTextBox"), TextBox).Text
        CType(FormView1.FindControl("DropDownList2"), DropDownList).SelectedValue = CType(FormView1.FindControl("CashierTextBox"), TextBox).Text
        If CType(FormView1.FindControl("EffectiveTextBox"), TextBox).Text = "N" Then
            CType(FormView1.FindControl("Button5"), Button).Text = "取消停用"
            CType(FormView1.FindControl("Button3"), Button).Visible = False
            CType(FormView1.FindControl("DropDownList3"), DropDownList).Enabled = False
            CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = False
            CType(FormView1.FindControl("DropDownList2"), DropDownList).Enabled = False
            CType(FormView1.FindControl("CustomerNameTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("AddressTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("EMAILTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("TelNoTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("ContactTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("GUINumberTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("Date1TextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("Date2TextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("DaysAllowedTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("TitleOfInvoiceTextBox"), TextBox).Enabled = False
            CType(FormView1.FindControl("MemoTextBox"), TextBox).Enabled = False
        Else
            tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea] WHERE Effective = 'Y' "
            Me.SqlDataSource2.SelectCommand = tmpsql
            Me.SqlDataSource2.DataBind()

            Me.SqlDataSource4.DataBind()
            CType(FormView1.FindControl("DropDownList3"), DropDownList).SelectedValue = CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Text
            dispEmployee()
        End If
        'dispEmployee()
        CType(FormView1.FindControl("DropDownList3"), DropDownList).Focus()
    End Sub
End Class
