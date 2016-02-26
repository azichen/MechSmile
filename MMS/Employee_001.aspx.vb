Imports System.Data

Partial Class MMS_Employee_001
    Inherits System.Web.UI.Page

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
                    Me.TitleLabel.Text = "員工資料 - 新增"
                    CType(FormView1.FindControl("EmployeeNoTextBox"), TextBox).Focus()
                Case "edit"
                    Me.FormView1.ChangeMode(FormViewMode.Edit)
                    Me.TitleLabel.Text = "員工資料 - 修改"
                    Dim tmpsql As String
                    tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea] where Effective='Y'"
                    Me.SqlDataSource2.SelectCommand = tmpsql
                    Me.SqlDataSource2.DataBind()
                    tmpsql = "select * from  MMSEmployee where EmployeeNo='" + Request("STNID").Trim + "'"
                    Me.SqlDataSource1.SelectCommand = tmpsql
                    Me.SqlDataSource1.DataBind()
                    Dim dv As System.Data.DataView = SqlDataSource1.Select(DataSourceSelectArguments.Empty)
                    Dim dt As System.Data.DataTable = dv.Table
                    If dt.Rows(0).Item("Salesman") = "Y" Then
                        CType(FormView1.FindControl("CheckBox1"), CheckBox).Checked = True
                    End If
                    If dt.Rows(0).Item("Cashier") = "Y" Then
                        CType(FormView1.FindControl("CheckBox2"), CheckBox).Checked = True
                    End If
                    If dt.Rows(0).Item("accountant") = "Y" Then
                        CType(FormView1.FindControl("CheckBox3"), CheckBox).Checked = True
                    End If
                    CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue = dt.Rows(0).Item("AreaCode")
                    If CType(FormView1.FindControl("EffectiveTextBox"), TextBox).Text = "N" Then
                        CType(FormView1.FindControl("Button5"), Button).Text = "取消停用"
                        CType(FormView1.FindControl("Button3"), Button).Visible = False
                        
                        CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("CheckBox1"), CheckBox).Enabled = False
                        CType(FormView1.FindControl("CheckBox2"), CheckBox).Enabled = False
                        CType(FormView1.FindControl("CheckBox3"), CheckBox).Enabled = False
                        CType(FormView1.FindControl("ExtNoTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("AccountsTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = False
                        tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
                        Me.SqlDataSource2.SelectCommand = tmpsql
                        Me.SqlDataSource2.DataBind()
                    End If
                    CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).Focus()
                    Me.disable_edit()
            End Select
            Try

                If ViewState("ROL").Substring(2, 1) = "N" Then CType(FormView1.FindControl("Button6"), Button).Visible = False
                If ViewState("ROL").Substring(3, 1) = "N" Then CType(FormView1.FindControl("Button5"), Button).Visible = False
            Catch ex As Exception

            End Try
        End If
    End Sub


    '修改檢查輸入資料
    Private Function CheckData_detail1() As String
        CheckData_detail1 = ""
        Dim getEmployeeNo As String = CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text
        Dim getEmployeeName As String = CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).Text
        Dim getAccounts As String = CType(FormView1.FindControl("AccountsTextBox"), TextBox).Text
        Dim getAreaCode As String = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
        Dim getExtNo As String = CType(FormView1.FindControl("ExtNoTextBox"), TextBox).Text
        Dim getEmailAdr As String = CType(FormView1.FindControl("EmailAdrTextBox"), TextBox).Text

        Dim ch1 As Boolean = CType(FormView1.FindControl("CheckBox1"), CheckBox).Checked
        Dim ch2 As Boolean = CType(FormView1.FindControl("CheckBox2"), CheckBox).Checked
        Dim ch3 As Boolean = CType(FormView1.FindControl("CheckBox3"), CheckBox).Checked

        'If getAreaCode = "" Then
        '    CheckData_detail1 += "區域不可空白! \n"
        'End If

        If ch1 = False And ch2 = False And ch3 = False Then
            CheckData_detail1 += "最少選擇一種員工類別! \n"
        End If
    End Function


    '新增檢查輸入資料
    Private Function CheckData_detail() As String
        CheckData_detail = ""
        Dim getEmployeeNo As String = CType(FormView1.FindControl("EmployeeNoTextBox"), TextBox).Text
        Dim getEmployeeName As String = CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).Text
        Dim getAccounts As String = CType(FormView1.FindControl("AccountsTextBox"), TextBox).Text
        Dim getAreaCode As String = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
        Dim getExtNo As String = CType(FormView1.FindControl("ExtNoTextBox"), TextBox).Text

        Dim ch1 As Boolean = CType(FormView1.FindControl("CheckBox1"), CheckBox).Checked
        Dim ch2 As Boolean = CType(FormView1.FindControl("CheckBox2"), CheckBox).Checked
        Dim ch3 As Boolean = CType(FormView1.FindControl("CheckBox3"), CheckBox).Checked
        If getEmployeeNo = "" Then
            CheckData_detail += "員工代號不可空白! \n"
        End If

        If getAreaCode = "" Then
            CheckData_detail += "區域不可空白! \n"
        End If

        If ch1 = False And ch2 = False And ch3 = False Then
            CheckData_detail += "最少選擇一種員工類別! \n"
        End If

        Dim sSql As String
        Dim sRow As Collection = New Collection
        sSql = "select * from MMSEmployee where EmployeeNo='" + getEmployeeNo + "'"
        sRow = modDB.GetRowData(sSql)
        If sRow.Count > 0 Then
            CheckData_detail += "員工代號已經被使用! \n"
        End If
    End Function

    '新增
    Protected Sub Button1_Click(sender As Object, e As System.EventArgs)
        Dim sMsg As String = ""
        Dim getMonSE As String = ""

        sMsg = CheckData_detail()

        If sMsg.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & sMsg)
            Exit Sub
        End If
        Dim getEmployeeNo As String = CType(FormView1.FindControl("EmployeeNoTextBox"), TextBox).Text
        Dim getEmployeeName As String = CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).Text
        Dim getAccounts As String = CType(FormView1.FindControl("AccountsTextBox"), TextBox).Text
        Dim getAreaCode As String = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
        Dim getExtNo As String = CType(FormView1.FindControl("ExtNoTextBox"), TextBox).Text
        Dim ch1 As Integer = CType(FormView1.FindControl("CheckBox1"), CheckBox).Checked
        Dim ch2 As Integer = CType(FormView1.FindControl("CheckBox2"), CheckBox).Checked
        Dim ch3 As Integer = CType(FormView1.FindControl("CheckBox3"), CheckBox).Checked
        Dim tmpsql As String

        tmpsql = "insert into MMSEmployee(EmployeeNo,EmployeeName,Accounts,AreaCode,ExtNo,EmailAdr,Salesman,Cashier,accountant,Effective)"
        tmpsql += " values("
        tmpsql += "'" + getEmployeeNo + "',"
        tmpsql += "'" + getEmployeeName + "',"
        tmpsql += "'" + getAccounts + "',"
        tmpsql += "'" + getAreaCode + "',"
        tmpsql += "'" + getExtNo + "',"
        If ch1 = -1 Then
            tmpsql += "'Y',"
        Else
            tmpsql += "'N',"
        End If
        If ch2 = -1 Then
            tmpsql += "'Y',"
        Else
            tmpsql += "'N',"
        End If
        If ch3 = -1 Then
            tmpsql += "'Y',"
        Else
            tmpsql += "'N',"
        End If
        tmpsql += "'Y')"
        Me.SqlDataSource1.InsertCommand = tmpsql
        Me.SqlDataSource1.Insert()

        SetAreaUsed(getAreaCode)
        Response.Redirect("Employee.aspx?Returnflag=1&msg=資料新增成功!!")

    End Sub

    '修改刪除
    Protected Sub Button5_Click(sender As Object, e As System.EventArgs)
        Dim getAreaCode As String = CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text
        Dim getTextBox1 As String = CType(FormView1.FindControl("TextBox1"), TextBox).Text
        Dim tmpsql, sMsg As String

        If CType(FormView1.FindControl("Button5"), Button).Text = "取消停用" Then
            tmpsql = "select * from MMSArea where AreaCode='" + CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue + "'"
            If get_DataTable(tmpsql).Rows(0).Item("Effective") = "N" Then
                modUtil.showMsg(Me.Page, "訊息", "此員工原所屬區域已經停用! \n 請重新選擇區域後點選更新, \n 以完成取消停用!")

                CType(FormView1.FindControl("Button5"), Button).Visible = False
                CType(FormView1.FindControl("Button3"), Button).Visible = True
                CType(FormView1.FindControl("EffectiveTextBox"), TextBox).Text = "Y"

                CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).Enabled = True 
                CType(FormView1.FindControl("CheckBox1"), CheckBox).Enabled = True
                CType(FormView1.FindControl("CheckBox2"), CheckBox).Enabled = True
                CType(FormView1.FindControl("CheckBox3"), CheckBox).Enabled = True
                CType(FormView1.FindControl("ExtNoTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("AccountsTextBox"), TextBox).Enabled = True
                CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = True

                tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea] WHERE [Effective] = 'Y'"
                SqlDataSource2.SelectCommand = tmpsql
                SqlDataSource2.DataBind()
            Else
                tmpsql = "update MMSEmployee set "
                tmpsql += " Effective='Y'"
                tmpsql += " where EmployeeNo='" + getAreaCode + "'"
                Me.SqlDataSource1.UpdateCommand = tmpsql
                Me.SqlDataSource1.Update()
                Response.Redirect("Employee.aspx?Returnflag=1&msg=該員工資料已取消停用!!")
            End If
        Else
            sMsg = CheckAreaChange()
            If sMsg.Length > 0 Then
                modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "此員工在客戶資料仍是業務員或收款員,修正後才可刪除!! \n")
                Exit Sub
            End If


            If getTextBox1 = "Y" Then
                tmpsql = "update MMSEmployee set "
                tmpsql += " Effective='N'"
                tmpsql += " where EmployeeNo='" + getAreaCode + "'"
                Me.SqlDataSource1.UpdateCommand = tmpsql
                Me.SqlDataSource1.Update()
                Response.Redirect("Employee.aspx?Returnflag=1&msg=該員工資料已使用過不可刪除,已將其停用!!")
            End If

            If getTextBox1 = "N" Then
                tmpsql = "delete MMSEmployee "
                tmpsql += " where EmployeeNo='" + getAreaCode + "'"
                Me.SqlDataSource1.DeleteCommand = tmpsql
                Me.SqlDataSource1.Delete()
                Response.Redirect("Employee.aspx?Returnflag=1&msg=該員工資料已刪除!!")
            End If
        End If
    End Sub

    '修改更新
    Protected Sub Button3_Click(sender As Object, e As System.EventArgs)
        Dim sMsg As String = ""
        Dim getMonSE As String = ""

        sMsg = CheckData_detail1()

        If sMsg.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & sMsg)
            Exit Sub
        End If

        sMsg = CheckTypeChange()
        If sMsg.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & sMsg)
            Exit Sub
        End If

        Dim getEmployeeNo As String = CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text
        Dim getEmployeeName As String = CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).Text
        Dim getAccounts As String = CType(FormView1.FindControl("AccountsTextBox"), TextBox).Text
        Dim getAreaCode As String = CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue
        Dim getExtNo As String = CType(FormView1.FindControl("ExtNoTextBox"), TextBox).Text
        Dim getEMailAdr As String = CType(FormView1.FindControl("EMailAdrTextBox"), TextBox).Text
        Dim getEffective As String = CType(FormView1.FindControl("EffectiveTextBox"), TextBox).Text
        Dim ch1 As Integer = CType(FormView1.FindControl("CheckBox1"), CheckBox).Checked
        Dim ch2 As Integer = CType(FormView1.FindControl("CheckBox2"), CheckBox).Checked
        Dim ch3 As Integer = CType(FormView1.FindControl("CheckBox3"), CheckBox).Checked
        Dim tmpsql As String
        tmpsql = "update MMSEmployee set "
        tmpsql += " EmployeeNo='" + getEmployeeNo + "',"
        tmpsql += " EmployeeName='" + getEmployeeName + "',"
        tmpsql += " Accounts='" + getAccounts + "',"
        tmpsql += " AreaCode='" + getAreaCode + "',"
        tmpsql += " ExtNo='" + getExtNo + "',"
        tmpsql += " EMailAdr='" + getEMailAdr + "',"
        tmpsql += " Effective='" + getEffective + "',"
        If ch1 = -1 Then
            tmpsql += " Salesman='Y',"
        Else
            tmpsql += " Salesman='N',"
        End If
        If ch2 = -1 Then
            tmpsql += " Cashier='Y',"
        Else
            tmpsql += " Cashier='N',"
        End If
        If ch3 = -1 Then
            tmpsql += "accountant='Y'"
        Else
            tmpsql += "accountant='N'"
        End If
        tmpsql += " where EmployeeNo='" + getEmployeeNo + "'"
        Me.SqlDataSource1.UpdateCommand = tmpsql
        Me.SqlDataSource1.Update()
        SetAreaUsed(getAreaCode)
        Response.Redirect("Employee.aspx?Returnflag=1&msg=資料修改成功!!")
    End Sub

    '修改取消
    Protected Sub Button4_Click(sender As Object, e As System.EventArgs)
        Response.Redirect("Employee.aspx?Returnflag=1&msg=")
    End Sub

    '新增取消
    Protected Sub Button2_Click(sender As Object, e As System.EventArgs)
        Response.Redirect("Employee.aspx?Returnflag=1&msg=")
    End Sub

    Private Function CheckAreaChange() As String
        CheckAreaChange = ""
        Dim tmpsql As String
        Dim dv As System.Data.DataView = SqlDataSource1.Select(DataSourceSelectArguments.Empty)
        Dim dt As System.Data.DataTable = dv.Table
        tmpsql = "select * from MMSCustomers where (Salesman='" + CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text + "'"
        tmpsql += " or  Cashier='" + CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text + "')"
        tmpsql += " and Effective='Y'"
        If get_DataTable(tmpsql).Rows.Count > 0 Then
            CheckAreaChange = "此員工在客戶資料仍是業務員或收款員,修正後才可變更區域!!\n"
            Exit Function
        End If
    End Function

    '檢查身份變更
    Private Function CheckTypeChange() As String
        CheckTypeChange = ""
        Dim tmpsql As String
        Dim dv As System.Data.DataView = SqlDataSource1.Select(DataSourceSelectArguments.Empty)
        Dim dt As System.Data.DataTable = dv.Table
        If CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Text <> CType(FormView1.FindControl("DropDownList1"), DropDownList).SelectedValue Then
            tmpsql = "select * from MMSCustomers where Salesman='" + CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text + "'"
            tmpsql += " or  Cashier='" + CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text + "'"
            tmpsql += " and Effective='Y'"
            If get_DataTable(tmpsql).Rows.Count > 0 Then
                CheckTypeChange = "此員工在客戶資料仍是業務員或收款員,修正後才可變更區域!!\n"
                Exit Function
            End If
        End If
        If CType(FormView1.FindControl("TextBox2"), TextBox).Text = "Y" And CType(FormView1.FindControl("CheckBox1"), CheckBox).Checked = False Then
            tmpsql = "select * from MMSCustomers where Salesman='" + CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text + "'"
            tmpsql += " and Effective='Y'"
            If get_DataTable(tmpsql).Rows.Count > 0 Then
                CheckTypeChange = "此員工在客戶資料仍是業務員,修正後才可變更身分!!\n"
                Exit Function
            End If
        End If
        If CType(FormView1.FindControl("TextBox3"), TextBox).Text = "Y" And CType(FormView1.FindControl("CheckBox2"), CheckBox).Checked = False Then
            tmpsql = "select * from MMSCustomers where Cashier='" + CType(FormView1.FindControl("EmployeeNoLabel1"), Label).Text + "'"
            tmpsql += " and Effective='Y'"
            If get_DataTable(tmpsql).Rows.Count > 0 Then
                CheckTypeChange = "此員工在客戶資料仍是收款員,修正後才可變更身分!!\n"
                Exit Function
            End If
        End If
    End Function

#Region "DB function"
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

    Protected Sub Button6_Click(sender As Object, e As System.EventArgs)
        enable_edit()
    End Sub

    Private Sub disable_edit()
        CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("ExtNoTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("AccountsTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("CheckBox1"), CheckBox).Enabled = False
        CType(FormView1.FindControl("CheckBox2"), CheckBox).Enabled = False
        CType(FormView1.FindControl("CheckBox3"), CheckBox).Enabled = False
        CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = False
        CType(FormView1.FindControl("Button3"), Button).Visible = False
    End Sub


    Private Sub enable_edit()
        CType(FormView1.FindControl("EmployeeNameTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("ExtNoTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("AccountsTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("CheckBox1"), CheckBox).Enabled = True
        CType(FormView1.FindControl("CheckBox2"), CheckBox).Enabled = True
        CType(FormView1.FindControl("CheckBox3"), CheckBox).Enabled = True
        CType(FormView1.FindControl("DropDownList1"), DropDownList).Enabled = True
        CType(FormView1.FindControl("Button3"), Button).Visible = True
        CType(FormView1.FindControl("Button6"), Button).Visible = False
    End Sub
End Class
