Imports System.Data
Partial Class MMS_Employee
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
            modDB.SetFields("EmployeeNo", "員工代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("EmployeeName", "員工姓名", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("Accounts", "會計代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("AreaCode", "區域", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("ExtNo", "分機號碼", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("EMailAdr", "電子信箱", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("EmployeeType", "類別", grdList, HorizontalAlign.Left, "", False)
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

            If Session("QryField") Is Nothing Then
                'Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 13, Me.ViewState("EmptyTable"))
            Else
                Call ReflashQryData()
            End If
            Me.TextBox2.Focus()
        End If

        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.item_defList.SelectCommand = Me.txtSql.Text
        '--------------------------------------------------* 設定GridView資料來源ID
        Me.TextBox2.Attributes.Add("onkeypress", "return ToUpper()")
        Me.ScriptManager1.RegisterPostBackControl(Me.ExcelButton)
    End Sub


    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub ExcelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExcelButton.Click
        grdList.AllowPaging = False
        grdList.AllowSorting = False
        grdList.DataBind()
        modUtil.GridView2Excel("MMSEmployee.xls", Me.grdList)
        grdList.AllowPaging = True
        grdList.AllowSorting = True
        grdList.DataBind()
    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub QueryButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QueryButton.Click
        Call ShowGrid()
        Me.MsgLabel.Text = ""
        Session("employeeQ") = Me.item_defList.SelectCommand
    End Sub

    '******************************************************************************************************
    '* 新增一筆資料
    '******************************************************************************************************
    Protected Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Session("QryField") = Me.ViewState("QryField")
        Session("employeeQ") = Me.item_defList.SelectCommand
        Session("iCurPage") = grdList.PageIndex
        Response.Redirect("Employee_001.aspx?FormMode=add")
    End Sub

    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        Dim sSql As String, sFields As New Hashtable

        sSql = "select a.EmployeeNo,a.EmployeeName,a.Accounts,b.AreaName as AreaCode,a.ExtNo,a.EMailAdr"
        sSql += " from MMSEmployee a"
        sSql += " left outer join MMSArea b on a.AreaCode=b.AreaCode "
        sSql += " where a.EmployeeNo<>'x'"
        If Me.TextBox2.Text.Trim <> "" Then
            sSql += " and a.EmployeeNo >= '" + Me.TextBox2.Text.Trim + "'"
        End If
        If Me.TextBox3.Text.Trim <> "" Then
            sSql += " and a.EmployeeNo <= '" + Me.TextBox3.Text.Trim + "'"
        End If
        If Me.TextBox1.Text.Trim <> "" Then
            sSql += " and a.EmployeeName like '%" + Me.TextBox1.Text.Trim + "%'"
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

        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.item_defList.SelectCommand = sSql
        Me.txtSql.Text = sSql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        Me.grdList.DataSource = Nothing
        Me.grdList.DataSourceID = Me.item_defList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))

            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
            Me.ExcelButton.Enabled = False

        Else
            Me.ExcelButton.Enabled = True
        End If
        sFields.Add("txtSql", sSql)
        Me.ViewState("QryField") = sFields
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
            CType(Me.form1.FindControl(sKey(I)), TextBox).Text = sVal(I)
        Next
        Session("QryField") = Nothing
        Call ShowGrid()
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
            Response.Redirect("Employee_001.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
        End With

        Session("iCurPage") = grdList.PageIndex
    End Sub
#End Region

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub ClearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearButton.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.CheckBox1.Checked = False
        modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))
        Session("employeeQ") = Nothing
        Me.ExcelButton.Enabled = False
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

    Protected Sub TextBox2_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox2.TextChanged
        Me.TextBox2.Text = Me.TextBox2.Text.ToUpper
        If Me.TextBox2.Text.Trim <> "" And Me.TextBox3.Text.Trim <> "" Then
            Me.TextBox3.Text = Me.TextBox2.Text
        End If
    End Sub
End Class
