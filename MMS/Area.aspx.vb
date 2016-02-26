
Partial Class MMS_Areas
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
            modDB.SetFields("AreaCode", "區域代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("AreaName", "區域名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("TelNo", "電話", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("FaxNo", "傳真", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("Address", "地址", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("TaxUnit", "稅額單位", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("ProfitNo", "中心代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("UnitCode", "單位代號", grdList, HorizontalAlign.Left, "{0:N0}", False)
            '--------------------------------------------------* 加入連結欄位
            Dim sField As New CommandField
            sField.ShowSelectButton = True
            sField.HeaderText = "功能"
            sField.SelectText = "瀏覽"
            sField.HeaderStyle.ForeColor = Drawing.Color.White
            sField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            Me.grdList.Columns.Add(sField)
            Me.grdList.RowStyle.Height = 18

            Me.TextBox1.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox2.Attributes.Add("onkeypress", "return ToUpper()")

            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 8, Me.ViewState("EmptyTable"))
            Me.MsgLabel.Text = Request("msg")

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
    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub ExcelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExcelButton.Click
        If Me.grdList.PageCount > 1 Then
            grdList.AllowPaging = False
            grdList.AllowSorting = False
            grdList.DataBind()
            modUtil.GridView2Excel("MMSAreas.xls", Me.grdList)
            grdList.AllowPaging = True
            grdList.AllowSorting = True
            grdList.DataBind()
        Else

            Me.grdList.DataSource = Me.item_defList
            Me.grdList.DataBind()
            modUtil.GridView2Excel("MMSAreas.xls", Me.grdList)
        End If

    End Sub

    '******************************************************************************************************
    '* 新增一筆資料
    '******************************************************************************************************
    Protected Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Session("QryField") = Me.ViewState("QryField")
        Response.Redirect("Area_001.aspx?FormMode=add")
    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub QueryButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QueryButton.Click
        Call ShowGrid()
        Me.MsgLabel.Text = ""
    End Sub

    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        Dim sSql As String, sFields As New Hashtable

        If TextBox1.Text <> "" Then
            sFields.Add("TextBox1", TextBox1.Text)
        End If

        If TextBox2.Text <> "" Then
            sFields.Add("TextBox2", TextBox2.Text)
        End If

        sSql = "select * from MMSArea where AreaCode<>'XXX'"
        If Me.TextBox1.Text.Trim <> "" Or Me.TextBox2.Text.Trim <> "" Then
            sSql += " and AreaCode between '" + Me.TextBox1.Text.Trim + "' and '" + Me.TextBox2.Text.Trim + "'"
        End If
        If Me.CheckBox1.Checked = False Then
            sSql += " and Effective='Y'"
        End If
        sFields.Add("txtSql", sSql)
        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.item_defList.SelectCommand = sSql
        Me.txtSql.Text = sSql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        'Me.grdList.DataSourceID = Me.item_defList.ID
        Me.grdList.DataSource = Me.item_defList
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
            Response.Redirect("Area_001.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
        End With
    End Sub
#End Region

    Protected Sub ClearButton_Click(sender As Object, e As System.EventArgs) Handles ClearButton.Click
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.CheckBox1.Checked = False
        Me.ExcelButton.Enabled = False
        Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 8, Me.ViewState("EmptyTable"))
    End Sub

    Protected Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Me.TextBox2.Text = Me.TextBox1.Text
    End Sub
End Class
