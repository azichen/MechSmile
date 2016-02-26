Imports System.Data
Partial Class MMS_InvoiceExport1
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String
    Dim tt1, tt2 As DataTable

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

            '檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else
                Dim sFileName As String = Request.Path.Substring(Request.Path.LastIndexOf("/") + 1, Request.Path.Length - (Request.Path.LastIndexOf("/") + 1))
                Dim sRol As String = HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
                ViewState("ROL") = modUtil.GetRolData(sFileName, sRol)
            End If

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            modDB.SetGridViewStyle(Me.grdList)          '* 套用gridview樣式
            modDB.SetFields("InvoiceNo", "發票號碼", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("發票日期", "發票日期", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("傳票日期", "傳票日期", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("發票區分", "發票區分", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("對象別", "對象別", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("客戶代號", "客戶代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("客戶名稱", "客戶名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("統一編號", "統一編號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("課稅區分", "課稅區分", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("未稅總金額", "未稅總金額", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("總稅額", "總稅額", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("稅額單位", "稅額單位", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("中心代號", "中心代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("業務員", "業務員", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("單位代號", "單位代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("備註", "備註", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("有效", "有效", grdList, HorizontalAlign.Left, "", False)
            '--------------------------------------------------* 加入連結欄位
            Me.grdList.RowStyle.Height = 18

            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 16, Me.ViewState("EmptyTable"))

            dispArea()
            If Me.DropDownList1.Items.Count > 0 Then
                dispEmployee()
            End If
            modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
            modUtil.SetDateObj(Me.txtDtFrom0, False, Me.txtDtTo0, False)
            modUtil.SetDateObj(Me.txtDtTo0, False, Nothing, False)

            Me.txtDtFrom1.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtDtFrom2.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtDtFrom3.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtDtFrom4.Attributes.Add("onkeypress", "return ToUpper()")
            txtDtFrom.Focus()

            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 16, Me.ViewState("EmptyTable"))
            Else
                Call ReflashQryData()
            End If
        End If
        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = Me.txtSql.Text
        '--------------------------------------------------* 設定GridView資料來源ID

        Me.ScriptManager1.RegisterPostBackControl(Me.ExcelButton)
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

    Private Sub dispArea()
        Dim tmpsql As String
        tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
        tt1 = get_DataTable(tmpsql)
        Me.DropDownList1.DataTextField = "AreaName"
        Me.DropDownList1.DataValueField = "AreaCode"
        DropDownList1.DataSource = tt1
        DropDownList1.DataBind()

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



    '顯示區域所屬業務員收款員
    Private Sub dispEmployee()
        Dim tmpsql As String

        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee"
        tmpsql += " WHERE Effective = 'Y' and AreaCode='" + Me.DropDownList1.SelectedValue + "'"
        tmpsql += " and Salesman = 'Y'"
        tt2 = get_DataTable(tmpsql)
        Me.DropDownList2.DataTextField = "EmployeeName"
        Me.DropDownList2.DataValueField = "EmployeeNo"
        DropDownList2.DataSource = tt2
        DropDownList2.DataBind()

    End Sub

    Protected Sub btnQry_Click(sender As Object, e As System.EventArgs) Handles ExcelButton.Click
        If Me.grdList.PageCount > 1 Then
            Me.grdList.DataSource = Nothing
            Me.grdList.DataSourceID = Me.dscList.ID
            grdList.AllowPaging = False
            grdList.AllowSorting = False
            grdList.DataBind()
            modUtil.GridView2Excel("MMSInvoiceExport.xls", Me.grdList)
            grdList.AllowPaging = True
            grdList.AllowSorting = True
            grdList.DataBind()
        Else

            'grdList.AllowPaging = False
            'Me.grdList.DataSource = Nothing
            Me.grdList.DataSource = Me.dscList
            grdList.DataBind()
            modUtil.GridView2Excel("MMSInvoiceExport.xls", Me.grdList)
        End If

    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        dispEmployee()
    End Sub

    Protected Sub CheckBox2_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Me.DropDownList1.Enabled = Me.CheckBox2.Checked
        Me.DropDownList2.Enabled = Me.CheckBox2.Checked
    End Sub

    Protected Sub btnQry1_Click(sender As Object, e As System.EventArgs) Handles btnQry1.Click
        Me.txtDtFrom.Text = ""
        Me.txtDtFrom0.Text = ""
        Me.txtDtFrom1.Text = ""
        Me.txtDtFrom2.Text = ""
        Me.txtDtFrom3.Text = ""
        Me.txtDtFrom4.Text = ""
        Me.txtDtTo.Text = ""
        Me.txtDtTo0.Text = ""
        Me.DropDownList1.SelectedIndex = 0
        Me.dispEmployee()
        Me.CheckBox2.Checked = False
        Me.CheckBox1.Checked = False
        Me.DropDownList1.Enabled = False
        Me.DropDownList2.Enabled = False
        Me.ExcelButton.Enabled = False
        Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 16, Me.ViewState("EmptyTable"))

    End Sub

    Protected Sub txtDtFrom1_TextChanged(sender As Object, e As System.EventArgs) Handles txtDtFrom1.TextChanged
        If Me.txtDtFrom2.Text.Trim = "" Then Me.txtDtFrom2.Text = Me.txtDtFrom1.Text
    End Sub

    Protected Sub txtDtFrom3_TextChanged(sender As Object, e As System.EventArgs) Handles txtDtFrom3.TextChanged
        If Me.txtDtFrom4.Text.Trim = "" Then Me.txtDtFrom4.Text = Me.txtDtFrom3.Text
    End Sub

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        Call ShowGrid()

    End Sub

    Private Sub ShowGrid()
        Dim tmpstr As String, sFields As New Hashtable
        tmpstr = ""
        tmpstr += Me.txtDtFrom.Text
        tmpstr += Me.txtDtFrom0.Text
        tmpstr += Me.txtDtFrom1.Text
        tmpstr += Me.txtDtFrom2.Text
        tmpstr += Me.txtDtFrom3.Text
        tmpstr += Me.txtDtFrom4.Text
        tmpstr += Me.txtDtTo.Text
        tmpstr += Me.txtDtTo0.Text
        If Me.CheckBox2.Checked = True Then
            tmpstr += " "
        End If
        If tmpstr = "" Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "您未輸入任何查詢條件!!")
            Exit Sub
        End If
        Dim tmpsql As String
        'EXEC MMS_E01 發票日期(起),發票日期(迄),申請日期(起),申請日期(迄),客戶代號(起),客戶代號(迄),發票號碼(起),發票號碼(訖),業務員,過濾作廢發票
        tmpsql = "EXEC MMS_E01 "
        If Me.txtDtFrom0.Text <> "" Then '發票日期(起)
            tmpsql += "'" + Me.txtDtFrom0.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtTo0.Text <> "" Then '發票日期(迄)
            tmpsql += "'" + Me.txtDtTo0.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtFrom.Text <> "" Then '申請日期(起)
            tmpsql += "'" + Me.txtDtFrom.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtTo.Text <> "" Then '申請日期(迄)
            tmpsql += "'" + Me.txtDtTo.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtFrom1.Text <> "" Then '客戶代號(起)
            tmpsql += "'" + Me.txtDtFrom1.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtFrom2.Text <> "" Then '客戶代號(迄)
            tmpsql += "'" + Me.txtDtFrom2.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtFrom3.Text <> "" Then '發票號碼(起)
            tmpsql += "'" + Me.txtDtFrom3.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtFrom4.Text <> "" Then '發票號碼(迄)
            tmpsql += "'" + Me.txtDtFrom4.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.CheckBox2.Checked = True Then '業務員
            tmpsql += "'" + Me.DropDownList2.SelectedValue + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.CheckBox1.Checked = True Then '過濾作廢發票
            tmpsql += "'N'"
        Else
            tmpsql += "'Y'"
        End If

        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = tmpsql
        Me.txtSql.Text = tmpsql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        'Me.grdList.DataSource = Nothing
        Me.grdList.DataSource = Me.dscList
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Me.ExcelButton.Enabled = False
        Else
            Me.ExcelButton.Enabled = True
        End If
        If Me.grdList.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "查無資料!!")
        End If
        sFields.Add("txtSql", tmpsql)
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
            Response.Redirect("Employee_001.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
        End With

        Session("iCurPage") = grdList.PageIndex
    End Sub
#End Region
End Class