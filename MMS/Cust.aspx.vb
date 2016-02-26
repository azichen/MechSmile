
Partial Class MMS_Cust
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
            modDB.SetFields("CustomerNo", "客戶代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("CustomerName", "客戶名稱", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("AreaCode", "區域", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("GUINumber", "統一編號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("Salesman", "業務員", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("Cashier", "收款員", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("TELNO", "電話", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("ADDRESS", "住址", grdList, HorizontalAlign.Left, "", False)
            '--------------------------------------------------* 加入連結欄位
            Dim sField As New CommandField
            sField.ShowSelectButton = True
            sField.HeaderText = "功能"
            sField.SelectText = "瀏覽"
            sField.HeaderStyle.ForeColor = Drawing.Color.White
            sField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            Me.grdList.Columns.Add(sField)
            Me.grdList.RowStyle.Height = 18

            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))
            Me.MsgLabel.Text = Request("msg")



            Me.TextBox4.Attributes.Add("onkeypress", "return CheckKeyNumber()")

            Me.TextBox1.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox2.Attributes.Add("onkeypress", "return ToUpper()")
            If Session("QryField") Is Nothing Then
                'Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 13, Me.ViewState("EmptyTable"))
            Else
                Call ReflashQryData()
            End If
            Me.TextBox1.Focus()
        End If

        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.item_defList.SelectCommand = Me.txtSql.Text
        '--------------------------------------------------* 設定GridView資料來源ID

        Me.ScriptManager1.RegisterPostBackControl(Me.ExcelButton)
    End Sub

    Protected Sub btnNew_Click(sender As Object, e As System.EventArgs) Handles btnNew.Click
        '檢查是否有設定員工資料
        '檢查是否有業務員
        Dim sSql As String
        Dim CheckData_detail As String = ""
        Dim sRow As Collection = New Collection
        sSql = "select * from [MMSEmployee] where [Salesman]='Y' and Effective='Y'"
        sRow = modDB.GetRowData(sSql)
        If sRow.Count = 0 Then
            CheckData_detail += "請先新增業務員資料! \n"
        End If
        sSql = "select * from [MMSEmployee] where [Cashier]='Y' and Effective='Y'"
        sRow = modDB.GetRowData(sSql)
        If sRow.Count = 0 Then
            CheckData_detail += "請先新增收款員資料! \n"
        End If

        If CheckData_detail.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & CheckData_detail)
        Else
            Session("QryField") = Me.ViewState("QryField")
            Response.Redirect("Cust_001.aspx?FormMode=add")
        End If


    End Sub
    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub QueryButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QueryButton.Click
        Call ShowGrid()
        Me.MsgLabel.Text = ""
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
            Try
                CType(Me.form1.FindControl(sKey(I)), TextBox).Text = sVal(I)
            Catch ex As Exception
            End Try
            Try
                CType(Me.form1.FindControl(sKey(I)), CheckBox).Checked = sVal(I)
            Catch ex As Exception
            End Try
        Next
        Session("QryField") = Nothing
        Call ShowGrid()
    End Sub



    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        If Me.TextBox1.Text.Trim <> "" Or Me.TextBox2.Text.Trim <> "" Then
            If Me.TextBox1.Text = "" Then
                modUtil.showMsg(Me.Page, "訊息", "客戶代號起必須有資料！")
                Me.ExcelButton.Enabled = False
                Exit Sub
            End If
            If Me.TextBox2.Text = "" Then
                modUtil.showMsg(Me.Page, "訊息", "客戶代號迄必須有資料！")
                Me.ExcelButton.Enabled = False
                Exit Sub
            End If
        End If
        Dim sSql As String, sFields As New Hashtable
        sSql = "select a.CustomerNo,a.CustomerName,b.AreaName as AreaCode,a.GUINumber,c.EmployeeName as Salesman,d.EmployeeName as Cashier "
        sSql += " ,A.TELNO,A.ADDRESS"
        sSql += " from [MMSCustomers] a"
        sSql += " left outer join MMSArea b on a.AreaCode = b.AreaCode"
        sSql += " left outer join MMSEmployee c on a.Salesman=c.EmployeeNo"
        sSql += " left outer join MMSEmployee d on a.Cashier=d.EmployeeNo "
        sSql += " where a.CustomerNo<>'xxxxxx'"
        If get_area(Request.Cookies("STNID").Value) <> "" Then
            sSql += " and a.AreaCode ='" + get_area(My.User.Name) + "'"
        End If
        If Me.TextBox1.Text.Trim <> "" Or Me.TextBox2.Text.Trim <> "" Then
            sSql += " and ([CustomerNo]>='" + Me.TextBox1.Text + "' and [CustomerNo]<='" + Me.TextBox2.Text + "')"
        End If
        If Me.TextBox3.Text.Trim <> "" Then
            sSql += " and [CustomerName] like '%" + Me.TextBox3.Text.Trim + "%'"
        End If
        If Me.TextBox4.Text.Trim <> "" Then
            sSql += " and [GUINumber] like '%" + Me.TextBox4.Text.Trim + "%'"
        End If
        If Me.CheckBox1.Checked = False Then
            sSql += " and a.Effective='Y'"
        End If

        If TextBox1.Text <> "" Then
            sFields.Add("TextBox1", TextBox1.Text)
        End If

        If TextBox2.Text <> "" Then
            sFields.Add("TextBox2", TextBox2.Text)
        End If

        If TextBox3.Text <> "" Then
            sFields.Add("TextBox3", TextBox3.Text)
        End If

        If TextBox4.Text <> "" Then
            sFields.Add("TextBox4", TextBox4.Text)
        End If

        sFields.Add("txtSql", sSql)
        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.item_defList.SelectCommand = sSql
        Me.txtSql.Text = sSql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        Me.grdList.DataSource = Nothing
        Me.grdList.DataSourceID = Me.item_defList.ID
        'Me.grdList.DataSource = Me.item_defList
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 8, Me.ViewState("EmptyTable"))

            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
            Me.ExcelButton.Enabled = False

        Else
            Me.ExcelButton.Enabled = True
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
            Response.Redirect("Cust_001.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
        End With
    End Sub
#End Region

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub ClearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearButton.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.CheckBox1.Checked = False
        modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))
        Session("employeeQ") = Nothing
        ExcelButton.Enabled = False
    End Sub

    Protected Sub TextBox1_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox1.TextChanged
        Me.TextBox2.Text = Me.TextBox1.Text
    End Sub

    Protected Sub ExcelButton_Click(sender As Object, e As System.EventArgs) Handles ExcelButton.Click
        grdList.AllowPaging = False
        grdList.AllowSorting = False
        grdList.DataBind()
        modUtil.GridView2Excel("MMSCust.xls", Me.grdList)
        grdList.AllowPaging = True
        grdList.AllowSorting = True
        grdList.DataBind()
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
    Private Function get_area(ByVal id As String) As String
        Dim tmpsql As String
        tmpsql = "select * from MMSEmployee where EmployeeNo='" + id + "'"
        Try
            get_area = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
        Catch ex As Exception
            get_area = ""
        End Try
    End Function
End Class
