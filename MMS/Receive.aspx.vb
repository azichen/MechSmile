
Partial Class MMS_Receive
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
            modDB.SetFields("ReceivablesNo", "收款單號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("CustomerNo", "客戶代碼", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("CustomerName", "客戶名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("DateOfReceipt", "收款日期", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("TotalAmount1", "合計金額", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("Cashier", "收款員", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("Statusn", "狀態", grdList, HorizontalAlign.Left, "", False)
            '--------------------------------------------------* 加入連結欄位
            Dim sField As New CommandField
            sField.ShowSelectButton = True
            sField.HeaderText = "功能"
            sField.SelectText = "瀏覽"
            sField.HeaderStyle.ForeColor = Drawing.Color.White
            sField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            Me.grdList.Columns.Add(sField)

            Me.grdList.RowStyle.Height = 18

            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 7, Me.ViewState("EmptyTable"))
            Me.lblMsg.Text = Request("msg")


            Me.txtCmpID.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtSaleID1.Attributes.Add("onkeypress", "return ToUpper()")

            Dim tmpsql As String
            tmpsql = "select * from MMSEmployee where EmployeeNo='" + My.User.Name + "' and Salesman='Y'"
            Dim tt As System.Data.DataTable
            tt = get_DataTable(tmpsql)
            If tt.Rows.Count > 0 Then
                Me.txtSaleID1.Text = tt.Rows(0).Item("EmployeeNo")
                Me.txtSaleID1.Text = tt.Rows(0).Item("EmployeeNo")
                Me.Label1.Text = tt.Rows(0).Item("EmployeeName")
                Me.Label2.Text = tt.Rows(0).Item("EmployeeName")
                Me.txtSaleID1.Enabled = False
            End If

            modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)

            If Session("QryField") Is Nothing Then
                'Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 13, Me.ViewState("EmptyTable"))
            Else
                Call ReflashQryData()
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
            CType(Me.form1.FindControl(sKey(I)), TextBox).Text = sVal(I)
        Next
        Session("QryField") = Nothing
        Call ShowGrid()
    End Sub

    Protected Sub btnNew_Click(sender As Object, e As System.EventArgs) Handles btnNew.Click
        Response.Redirect("ReceiveApply.aspx?FormMode=add")
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
        Dim sSql As String, sFields As New Hashtable
        Dim Areacode As String = ""
        Areacode = modUnset.GetAreaCode(My.User.Name).Trim

        sSql = "SELECT *,REPLACE(CONVERT(varchar(20), (CAST(TotalAmount AS money)), 1), '.00', '')  as TotalAmount1,"
        sSql += " case Status"
        sSql += " when 'Y' then '審核通過'"
        sSql += " when 'N' then '未審核'"
        sSql += " end as Statusn"
        sSql += " from MMSReceivablesM A "
        sSql += " where [ReceivablesNo] <>'x' "
        If (Areacode <> "") And (Areacode <> "F") Then
            sSql += " and (select AreaCode from MMSCustomers where CustomerNo=A.CustomerNo)='" + Areacode + "'"
            'modDB.InsertSignRecord("AziTest", "sSql=" + " and (select AreaCode from MMSCustomers where CustomerNo=A.CustomerNo)='" + Areacode + "'", My.User.Name)
        End If

        If Me.txtSaleID1.Text.Trim <> "" Then
            sSql += " and  [Cashier] = '" + Me.txtSaleID1.Text.Trim + "'"
            sFields.Add("txtSaleID1", txtSaleID1.Text)
        End If
        If Me.txtCmpID.Text.Trim <> "" Then
            sSql += " and  [CustomerNo] = '" + Me.txtCmpID.Text.Trim + "'"
            sFields.Add("txtCmpID", txtCmpID.Text)
        End If
        If Me.txtDtFrom.Text.Trim <> "" Then
            sSql += " and  [DocDate] >= '" + Me.txtDtFrom.Text.Trim + "'"
            sFields.Add("txtDtFrom", txtDtFrom.Text)
        End If
        If Me.txtDtTo.Text.Trim <> "" Then
            sSql += " and  [DocDate] <= '" + Me.txtDtTo.Text.Trim + "'"
            sFields.Add("txtDtTo", txtDtTo.Text)
        End If

        sSql += " ORDER BY CustomerNo,DateOfReceipt DESC "

        sFields.Add("txtSql", sSql)
        'If Me.txtSaleID1.Text.Trim <> "" Then
        '    sSql += " and [CustomerNo] in (select [CustomerNo] from [MMSContract] where [Salesman]='" + Me.txtSaleID1.Text + "' )"
        'End If
        ''--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = sSql
        Me.txtSql.Text = sSql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        Me.grdList.DataSource = Nothing
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))

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
            If .Cells(6).Text = "審核通過" Then
                Response.Redirect("ReceiveApply.aspx?FormMode=read&STNID=" & .Cells(0).Text)
            Else
                Response.Redirect("ReceiveApply.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
            End If
            'Response.Redirect("ReceiveApply.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
        End With
    End Sub
#End Region

    '清除
    Protected Sub btnClear_Click(sender As Object, e As System.EventArgs) Handles btnClear.Click
        Me.txtCmpID.Text = ""
        Me.txtDtFrom.Text = ""
        Me.txtDtTo.Text = ""
        Me.lblMsg.Text = ""
        If Me.txtSaleID1.Enabled = True Then
            Me.txtSaleID1.Text = ""
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If

        Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))
    End Sub


    Protected Sub txtCmpID_TextChanged(sender As Object, e As System.EventArgs) Handles txtCmpID.TextChanged
        Dim tmpsql As String
        tmpsql = "select * from [MMSCustomers] where [CustomerNo]='" + Me.txtCmpID.Text + "' "
        Dim tt As System.Data.DataTable
        tt = get_DataTable(tmpsql)
        If tt.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "無此客戶！")
            Me.Label2.Text = ""
            Exit Sub
        End If
        Me.Label2.Text = tt.Rows(0).Item("CustomerName")
    End Sub


    Protected Sub txtSaleID1_TextChanged(sender As Object, e As System.EventArgs) Handles txtSaleID1.TextChanged
        Dim tmpsql As String
        tmpsql = "select * from MMSEmployee where EmployeeNo='" + Me.txtSaleID1.Text + "' and Salesman='Y'"
        Dim tt As System.Data.DataTable
        tt = get_DataTable(tmpsql)
        If tt.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "無此業務人員！")
            Me.Label1.Text = ""
            Exit Sub
        End If
        Me.Label1.Text = tt.Rows(0).Item("EmployeeName")
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
End Class
