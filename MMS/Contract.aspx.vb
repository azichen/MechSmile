
Partial Class MMS_Contract
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String

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
                If ViewState("ROL").Substring(1, 1) = "N" Then Me.btnNew.Visible = False
            End If

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            modDB.SetGridViewStyle(Me.grdList)          '* 套用gridview樣式
            modDB.SetFields("ContractNo", "合約編號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("BuildingName", "案場名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("CustomerNo", "客戶代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("CustomerName", "客戶名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("TelNo", "聯絡電話", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("Contact", "聯絡人", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("Salesman", "業務員", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("Cashier", "收款員", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("AreaCode", "區域", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("Quantity", "數量", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("StartDateC", "合約期間起", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("EndDateC", "合約期間訖", grdList, HorizontalAlign.Left, "", False)
            '--------------------------------------------------* 加入連結欄位
            Dim sField As New CommandField
            sField.ShowSelectButton = True
            sField.HeaderText = "功能"
            sField.SelectText = "瀏覽"
            sField.HeaderStyle.ForeColor = Drawing.Color.White
            sField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            Me.grdList.Columns.Add(sField)

            Me.grdList.RowStyle.Height = 18

            Me.txtCmpID.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtCmpID0.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtContractID.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtContractID0.Attributes.Add("onkeypress", "return ToUpper()")

            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 9, Me.ViewState("EmptyTable"))
            Me.lblMsg.Text = Request("msg")

            If Session("QryField") Is Nothing Then
                'Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 13, Me.ViewState("EmptyTable"))
            Else
                Call ReflashQryData()
            End If
            modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)

            Dim tmpsql As String
            tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
            Me.DropDownList5.DataSourceID = Nothing
            Me.DropDownList5.DataSource = get_DataTable(tmpsql)
            Me.DropDownList5.DataBind()
            tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee" ' WHERE Effective = 'Y' "
            Me.SqlDataSource3.SelectCommand = tmpsql
            Me.SqlDataSource3.DataBind()
            tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee" ' WHERE Effective = 'Y' "
            Me.SqlDataSource4.SelectCommand = tmpsql
            Me.SqlDataSource4.DataBind()
            dispEmployee()

            'ViewState("ROL") = modUtil.GetRolData(Request)
            'If (Mid(ViewState("ROL"), 5, 2) <> "03") And (Mid(ViewState("ROL"), 5, 2) <> "05") Then FormsAuthentication.RedirectToLoginPage()

            txtCmpID.Focus()

            Dim ss As String = GetAreaCode()
            If ss = "" Then
                Me.CheckBox2.Visible = True
            Else
                Me.CheckBox2.Visible = False
            End If

            'Me.DropDownList5.Enabled = True

            If (GetAreaCode() = "") Or (GetAreaCode() = "F") Then
                Me.CheckBox2.Visible = True
            Else
                Me.CheckBox2.Visible = False
                Dim i As Integer
                For i = 0 To Me.DropDownList5.Items.Count - 1
                    If Me.DropDownList5.Items(i).Value = GetAreaCode() Then
                        Me.DropDownList5.SelectedIndex = i
                    End If
                Next
                Me.dispEmployee()
            End If


        End If

        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = Me.txtSql.Text
        '--------------------------------------------------* 設定GridView資料來源ID

        Me.ScriptManager1.RegisterPostBackControl(Me.btnExcel)
    End Sub

    '******************************************************************************************************
    '* 重新顯示查詢結果 
    '******************************************************************************************************
    Private Sub ReflashQryData()
        Dim sCol As Hashtable = Session("QryField")
        Dim sKey(sCol.Count), sVal(sCol.Count) As String
        Dim I As Integer
        sCol.Keys.CopyTo(sKey, 0)
        sCol.Values.CopyTo(sVal, 0)
        For I = 0 To sCol.Count - 1
            If sKey(I) = "CheckBox1" Then
                CType(Me.form1.FindControl(sKey(I)), CheckBox).Checked = sVal(I)
            Else
                CType(Me.form1.FindControl(sKey(I)), TextBox).Text = sVal(I)
            End If

        Next
        Session("QryField") = Nothing
        Call ShowGrid()
    End Sub

    Protected Sub btnNew_Click(sender As Object, e As System.EventArgs) Handles btnNew.Click
        Session("QryField") = Me.ViewState("QryField")
        Response.Redirect("Contract_001.aspx?FormMode=add")
    End Sub


    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub btnQry_Click(sender As Object, e As System.EventArgs) Handles btnQry.Click
        Call ShowGrid()
        Me.lblMsg.Text = ""
    End Sub

    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        If txtContractID2.Text.Trim + txtCmpID.Text.Trim + txtCmpID0.Text.Trim + txtContractID1.Text.Trim + txtContractID.Text.Trim + txtContractID0.Text.Trim + txtDtFrom.Text.Trim + txtDtTo.Text.Trim + TextBox1.Text.Trim + TextBox2.Text.Trim = "" Then
            If Me.CheckBox2.Checked = False And Me.CheckBox3.Checked = False And Me.CheckBox4.Checked = False Then
                modUtil.showMsg(Me.Page, "訊息", "尚未輸入任何查詢條件！")
                Me.btnExcel.Enabled = False
                Exit Sub
            End If
        End If
        Dim tmpstr As String = ""
        If txtCmpID.Text.Trim + txtCmpID0.Text.Trim <> "" Then
            If txtCmpID.Text.Trim = "" Then
                tmpstr += "請輸入客戶代號起! \n"
            End If
            If txtCmpID0.Text.Trim = "" Then
                tmpstr += "請輸入客戶代號訖! \n"
            End If
        End If
        If txtContractID.Text.Trim + txtContractID0.Text.Trim <> "" Then
            If txtContractID.Text.Trim = "" Then
                tmpstr += "請輸入合約編號起! \n"
            End If
            If txtContractID0.Text.Trim = "" Then
                tmpstr += "請輸入合約編號訖! \n"
            End If
        End If
        If txtDtFrom.Text.Trim + txtDtTo.Text.Trim <> "" Then
            If txtDtFrom.Text.Trim = "" Then
                tmpstr += "請輸入合約日期訖之起始日期! \n"
            End If
            If txtDtTo.Text.Trim = "" Then
                tmpstr += "請輸入合約日期訖之結束日期! \n"
            End If
        End If
        If txtDtFrom.Text.Trim <> "" And txtDtTo.Text.Trim <> "" And txtDtFrom.Text.Trim > txtDtTo.Text.Trim Then
            tmpstr += "合約日期訖之結束日期不可大於起始日期! \n"
        End If
        If tmpstr <> "" Then
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
            Me.btnExcel.Enabled = False
            Exit Sub
        End If
        Dim sSql As String, sFields As New Hashtable
        sSql = "SELECT a.[ContractNo],a.[CustomerNo],a.[BuildingName],D.[CustomerName],a.[TelNo],a.[Contact],a.Quantity,b.EmployeeName as  Salesman,c.EmployeeName as Cashier,convert(varchar, convert(datetime,StartDateC), 111) as StartDateC,convert(varchar, convert(datetime,EndDateC), 111) as EndDateC "
        sSql += " FROM [MMSContract] a"
        sSql += " Inner Join MMSCustomers D on a.CustomerNo=D.CustomerNo " 'modify in 20140729
        sSql += " left outer join MMSEmployee b on a.Salesman=b.EmployeeNo "
        sSql += " left outer join MMSEmployee c on a.Cashier=c.EmployeeNo "
        sSql += " where a.ContractNo <>'x' "
        If (GetAreaCode() <> "") And (GetAreaCode() <> "F") Then
            sSql += " and a.AreaCode ='" + GetAreaCode() + "'"
        End If
        If Me.txtCmpID.Text.Trim <> "" Then
            sSql += " and  a.[CustomerNo] >= '" + Me.txtCmpID.Text.Trim + "'"
            sFields.Add("txtCmpID", txtCmpID.Text)
        End If
        If Me.txtCmpID0.Text.Trim <> "" Then
            sSql += " and  a.[CustomerNo] <= '" + Me.txtCmpID0.Text.Trim + "'"
            sFields.Add("txtCmpID0", txtCmpID0.Text)
        End If
        If Me.txtContractID1.Text.Trim <> "" Then
            sSql += " and  D.[CustomerName] like '%" + Me.txtContractID1.Text.Trim + "%'"
            sFields.Add("txtContractID1", txtContractID1.Text)
        End If
        If Me.txtContractID2.Text.Trim <> "" Then
            'sSql += " and  D.[CustomerName] like '%" + Me.txtContractID2.Text.Trim + "%'" 'debug in 20140729
            sSql += " and  D.[BuildingName] like '%" + Me.txtContractID2.Text.Trim + "%'"
            sFields.Add("txtContractID2", txtContractID2.Text)
        End If
        If Me.txtContractID.Text.Trim <> "" Then
            sSql += " and  a.[ContractNo] >= '" + Me.txtContractID.Text.Trim + "'"
            sFields.Add("txtContractID", txtContractID.Text)
        End If
        If Me.txtContractID0.Text.Trim <> "" Then
            sSql += " and  a.[ContractNo] <= '" + Me.txtContractID0.Text.Trim + "'"
            sFields.Add("txtContractID0", txtContractID0.Text)
        End If
        If Me.txtDtFrom.Text.Trim <> "" Then
            sSql += " and  a.[EndDateC] >= '" + Me.txtDtFrom.Text.Trim + "'"
            sFields.Add("txtDtFrom", txtDtFrom.Text)
        End If
        If Me.txtDtTo.Text.Trim <> "" Then
            sSql += " and  a.[EndDateC] <= '" + Me.txtDtTo.Text.Trim + "'"
            sFields.Add("txtDtTo", txtDtTo.Text)
        End If
        If Me.CheckBox1.Checked = True Then
            sSql += " and ArchiveDateA=''"
            sFields.Add("CheckBox1", "1")
        End If
        If Me.CheckBox2.Checked Then
            sSql += " and a.AreaCode='" + Me.DropDownList5.SelectedValue + "'"
        End If
        If Me.CheckBox3.Checked Then
            sSql += " and b.EmployeeNo='" + Me.DropDownList6.SelectedValue + "'"
        End If
        If Me.CheckBox4.Checked Then
            sSql += " and a.Cashier='" + Me.DropDownList7.SelectedValue + "'"
        End If
        If Me.TextBox1.Text <> "" Then
            sSql += " and a.Contact='" + Me.TextBox1.Text + "'"
        End If
        If Me.TextBox2.Text <> "" Then
            sSql += " and a.TelNo='" + Me.TextBox2.Text + "'"
        End If

        sSql += " order by a.[CustomerNo],a.[ContractNo] desc"

        'modDB.InsertSignRecord("AziTest", "emp=" & Me.DropDownList6.SelectedValue, My.User.Name) 'AZITEST
        'modDB.InsertSignRecord("AziTest", "sSql" & sSql.Substring(450, 90), My.User.Name) 'AZITEST

        sFields.Add("txtSql", sSql)
        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = sSql
        Me.txtSql.Text = sSql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        Me.grdList.DataSource = Nothing
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 9, Me.ViewState("EmptyTable"))

            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
            Me.btnExcel.Enabled = False

        Else
            Me.btnExcel.Enabled = True
        End If
        Me.ViewState("QryField") = sFields
    End Sub


#Region "GridView Event"
    '******************************************************************************************************
    '* GridView_DataBound => 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.DataBound
        modDB.SetGridPageNum(Me.grdList, PagerButtons.NumericFirstLast, HorizontalAlign.Left)

    End Sub

    '******************************************************************************************************
    '* GridView_RowDataBound => 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub grdList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdList.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                If (CType(ViewState("LineNo"), Int16) = 0) Then
                    e.Row.Attributes.Add("onmouseout", "onmouseoutColor1(this);")
                    e.Row.Attributes.Add("onmouseover", "onmouseoverColor1(this);")
                    ViewState("LineNo") = 1
                Else
                    e.Row.Attributes.Add("onmouseout", "onmouseoutColor2(this);")
                    e.Row.Attributes.Add("onmouseover", "onmouseoverColor2(this);")
                    ViewState("LineNo") = 0
                End If
        End Select
    End Sub
    Protected Sub grdList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.SelectedIndexChanged
        Session("QryField") = Me.ViewState("QryField")
        With Me.grdList.SelectedRow
            Response.Redirect("Contract_001.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
        End With
    End Sub
#End Region

    '清除
    Protected Sub btnClear_Click(sender As Object, e As System.EventArgs) Handles btnClear.Click
        Me.txtCmpID.Text = ""
        Me.txtCmpID0.Text = ""
        Me.txtContractID1.Text = ""
        Me.txtContractID.Text = ""
        Me.txtContractID0.Text = ""
        Me.txtDtFrom.Text = ""
        Me.txtDtTo.Text = ""
        Me.lblMsg.Text = ""
        Me.btnExcel.Enabled = False
        Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 9, Me.ViewState("EmptyTable"))
    End Sub


    Protected Sub txtCmpID_TextChanged(sender As Object, e As System.EventArgs) Handles txtCmpID.TextChanged
        If Me.txtCmpID0.Text.Trim <> "" Then Exit Sub
        Me.txtCmpID0.Text = txtCmpID.Text
        Me.txtCmpID0.Focus()
    End Sub

    Protected Sub txtContractID_TextChanged(sender As Object, e As System.EventArgs) Handles txtContractID.TextChanged
        If Me.txtContractID0.Text.Trim <> "" Then Exit Sub
        Me.txtContractID0.Text = Me.txtContractID.Text
    End Sub

    Protected Sub btnExcel_Click(sender As Object, e As System.EventArgs) Handles btnExcel.Click
        If Me.grdList.PageCount > 1 Then
            grdList.AllowPaging = False
            grdList.AllowSorting = False
            grdList.DataBind()
            modUtil.GridView2Excel("MMSContract.xls", Me.grdList)
            grdList.AllowPaging = True
            grdList.AllowSorting = True
            grdList.DataBind()
        Else
            modUtil.GridView2Excel("MMSContract.xls", Me.grdList)
        End If
    End Sub


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
        tmpsql += " WHERE Effective = 'Y'  and AreaCode='" + Me.DropDownList5.SelectedValue + "'"
        tmpsql += " and Salesman = 'Y'"
        Me.SqlDataSource3.SelectCommand = tmpsql
        Me.SqlDataSource3.DataBind()

        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee"
        tmpsql += " WHERE Effective = 'Y'  and AreaCode='" + Me.DropDownList5.SelectedValue + "'"
        tmpsql += " and Cashier = 'Y'"
        Me.SqlDataSource4.SelectCommand = tmpsql
        Me.SqlDataSource4.DataBind()
        'Me.DropDownList6.Enabled = True
        'Me.DropDownList7.Enabled = True
    End Sub

    Protected Sub CheckBox2_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Me.DropDownList5.Enabled = CheckBox2.Checked
    End Sub

    Protected Sub CheckBox3_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Me.DropDownList6.Enabled = CheckBox3.Checked
    End Sub

    Protected Sub CheckBox4_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox4.CheckedChanged
        Me.DropDownList7.Enabled = CheckBox4.Checked
    End Sub


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


    '取得操作者區域
    'Private Function get_area(ByVal id As String) As String
    '    Dim tmpsql As String
    '    tmpsql = "select * from MMSEmployee where EmployeeNo='" + id + "'"
    '    Try
    '        get_area = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
    '    Catch ex As Exception
    '        get_area = ""
    '    End Try
    'End Function



    Private Function GetAreaCode() As String
        'Dim tmpsql As String
        'Dim STNID As String = Request.Cookies("STNID").Value
        'tmpsql = "select * from MECHSTNM where STNID='" + STNID + "'"
        'If get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "    " Or get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "4601" Then
        '    tmpsql = "select * from MMSArea where UNITCODE='" + get_DataTable(tmpsql).Rows(0).Item("STNUID") + "'"
        '    Dim tmparea As String
        '    Try
        '        GetAreaCode = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
        '    Catch ex As Exception
        '        GetAreaCode = ""
        '    End Try
        'Else
        '    GetAreaCode = ""
        'End If
        ''GetAreaCode = "C"
        Dim Areacode As String = ""
        GetAreaCode = modUnset.GetAreaCode(My.User.Name).Trim
    End Function

    Protected Sub DropDownList5_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList5.SelectedIndexChanged
        dispEmployee()
    End Sub
End Class
