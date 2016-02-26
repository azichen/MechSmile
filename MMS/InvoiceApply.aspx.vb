
Partial Class MMS_InvoiceApply
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

            '檢查是否有登錄
            'If (Not User.Identity.IsAuthenticated) Then
            '    FormsAuthentication.RedirectToLoginPage()
            'Else
            '    Dim sFileName As String = Request.Path.Substring(Request.Path.LastIndexOf("/") + 1, Request.Path.Length - (Request.Path.LastIndexOf("/") + 1))
            '    Dim sRol As String = HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
            '    ViewState("ROL") = modUtil.GetRolData(sFileName, sRol)
            '    If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
            '    If ViewState("ROL").Substring(1, 1) = "N" Then
            '        Me.btnNew.Visible = False
            '        Me.btnNew0.Visible = False
            '    End If
            'End If

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            modDB.SetGridViewStyle(Me.grdList)          '* 套用gridview樣式
            modDB.SetFields("DocNo", "申請單號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("CustomerNo", "客戶代碼", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("CustomerName", "客戶名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("InvoiceDate", "發票日期", grdList, HorizontalAlign.Center, "", False)
            modDB.SetFields("INVOICENO", "發票號碼", grdList, HorizontalAlign.Center, "", False)
            modDB.SetFields("CONTRACTNO", "合約代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("INVOICEPERIOD", "期別", grdList, HorizontalAlign.Left, "", False)
            'Add in 20141006
            modDB.SetFields("ITEMNAME", "品名", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("AMOUNT", "(含稅)金額", grdList, HorizontalAlign.Center, "", False)
            '
            modDB.SetFields("Applicant", "申請人", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("StatusN", "狀態", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("INVN", "發票", grdList, HorizontalAlign.Left, "", False)
            '--------------------------------------------------* 加入連結欄位
            Dim sField As New CommandField
            sField.ShowSelectButton = True
            sField.HeaderText = "功能"
            sField.SelectText = "瀏覽"
            sField.HeaderStyle.ForeColor = Drawing.Color.White
            sField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            Me.grdList.Columns.Add(sField)

            Me.grdList.RowStyle.Height = 18

            'Me.grdList.Columns([5].DefaultCellStyle.Format = "N2") 'X是你那一列所在的索引 

            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 12, Me.ViewState("EmptyTable"))
            Me.lblMsg.Text = Request("msg")

            Me.txtCmpID.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtContractID.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtContractID0.Attributes.Add("onkeypress", "return ToUpper()")
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
            txtSaleID1.Focus()
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
        Try
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
        Catch ex As Exception
            Session("QryField") = Nothing
        End Try

    End Sub

    Protected Sub btnNew_Click(sender As Object, e As System.EventArgs) Handles btnNew.Click
        Session("QryField") = Me.ViewState("QryField")
        Response.Redirect("InvoiceApply002.aspx?FormMode=add")
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
        '20140905
        Dim Areacode As String = ""
        Areacode = modUnset.GetAreaCode(My.User.Name).Trim
        '
        Dim sSql As String, sFields As New Hashtable
        'sSql = "SELECT A.DOCNO,CUSTOMERNO,CUSTOMERNAME,Applicant,CONTRACTNO,INVOICEPERIOD"

        'sSql = "SELECT A.DOCNO, CUSTOMERNO, CUSTOMERNAME, Applicant,C.INVOICENO, B.CONTRACTNO, B.INVOICEPERIOD "
        'sSql += " ,Case Status when 'Y' then '審核通過' when 'N' then '未審核' end as StatusN"
        'sSql += " ,B.itemname,Sum(C.Amount) As Amount " ' ADD IN 20141006
        'sSql += " From MMSInvoiceApplyM A INNER JOIN MMSINVOICEAPPLYD B ON A.DOCNO=B.DOCNO AND SERIAL=1"
        'sSql += "                         INNER JOIN MMSINVOICEAPPLYD C ON A.DOCNO=C.DOCNO "

        sSql = "SELECT A.DOCNO, A.CUSTOMERNO, D.CUSTOMERNAME, Applicant,B.InvoiceDate,B.INVOICENO,CONTRACTNO,B.INVOICEPERIOD,B.Itemname,B.Amount"
        sSql += " ,Case A.Status when 'Y' then '審核通過' when 'N' then '未審核' end as StatusN"
        sSql += " ,CASE C.Status WHEN 'Y' THEN '正常' WHEN 'V' THEN '作廢' ELSE '' END AS INVN"
        sSql += " From MMSInvoiceApplyM A "
        sSql += " INNER JOIN MMSINVOICEAPPLYD B ON A.DOCNO=B.DOCNO "
        sSql += " INNER JOIN MMSCUSTOMERS D ON A.CUSTOMERNO=D.CUSTOMERNO "
        sSql += " LEFT JOIN MMSINVOICEM C ON B.DOCNO=C.APPLYDOCNO"
        ' B.INVOICENO=C.INVOICENO"
        sSql += " where A.DocNo<>'x' "

        'If GetAreaCode() <> "" Then
        If Areacode <> "" Then
            sSql += " and (select AreaCode from MMSCustomers where CustomerNo=a.CustomerNo)='" + Areacode + "'"
        End If

        If Me.txtCmpID.Text.Trim <> "" Then
            sSql += " and  A.CustomerNo = '" + Me.txtCmpID.Text.Trim + "'"
            sFields.Add("txtCmpID", txtCmpID.Text)
        End If

        If Me.txtContractID.Text.Trim <> "" Then
            sSql += " and  B.ContractNo >= '" + Me.txtContractID.Text.Trim + "'"
            sFields.Add("txtContractID", txtContractID.Text)
        End If

        If Me.txtContractID0.Text.Trim <> "" Then
            sSql += " and B.ContractNo <= '" + Me.txtContractID0.Text.Trim + "'"
            sFields.Add("txtContractID0", txtContractID0.Text)
        End If
        If Me.txtDtFrom.Text.Trim <> "" Then
            sSql += " and ApplyDate >= '" + Me.txtDtFrom.Text.Trim + "'"
            sFields.Add("txtDtFrom", txtDtFrom.Text)
        End If
        If Me.txtDtTo.Text.Trim <> "" Then
            sSql += " and ApplyDate <= '" + Me.txtDtTo.Text.Trim + "'"
            sFields.Add("txtDtTo", txtDtTo.Text)
        End If

        If Me.txtSaleID1.Text.Trim <> "" Then
            sFields.Add("txtSaleID1", txtSaleID1.Text)
            sSql += " and A.CustomerNo in (select CustomerNo from MMSContract where Salesman='" + Me.txtSaleID1.Text + "' )"
        End If
        'sSql += " GROUP BY A.DOCNO,CUSTOMERNO,CUSTOMERNAME,Applicant,C.INVOICENO,B.CONTRACTNO,B.INVOICEPERIOD,Status,B.Itemname "
        '20141126
        'sSql += " Order BY A.DOCNO DESC "
        sSql += " Order BY A.CUSTOMERNO,B.InvoiceDate DESC "


        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = sSql
        Me.txtSql.Text = sSql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        Me.grdList.DataSource = Nothing
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 12, Me.ViewState("EmptyTable"))

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
            Response.Redirect("InvoiceApply002.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
        End With
    End Sub
#End Region

    '清除
    Protected Sub btnClear_Click(sender As Object, e As System.EventArgs) Handles btnClear.Click
        Me.txtCmpID.Text = ""
        Me.txtContractID.Text = ""
        Me.txtContractID0.Text = ""
        Me.txtDtFrom.Text = ""
        Me.txtDtTo.Text = ""
        Me.lblMsg.Text = ""
        If Me.txtSaleID1.Enabled = True Then
            Me.txtSaleID1.Text = ""
            Me.Label1.Text = ""
            Me.Label2.Text = ""
        End If

        Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 11, Me.ViewState("EmptyTable"))
    End Sub

    Protected Sub btnNew0_Click(sender As Object, e As System.EventArgs) Handles btnNew0.Click
        Response.Redirect("InvoiceApply001.aspx")
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

    Protected Sub txtContractID_TextChanged(sender As Object, e As System.EventArgs) Handles txtContractID.TextChanged
        txtContractID0.Text = txtContractID.Text
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

    Protected Sub btnExcel_Click(sender As Object, e As System.EventArgs) Handles btnExcel.Click
        If Me.grdList.PageCount > 1 Then
            grdList.AllowPaging = False
            grdList.AllowSorting = False
            grdList.DataBind()
            modUtil.GridView2Excel("MMSInvoiceApply.xls", Me.grdList)
            grdList.AllowPaging = True
            grdList.AllowSorting = True
            grdList.DataBind()
        Else
            Me.grdList.DataSource = Nothing
            Me.grdList.DataSourceID = Me.dscList.ID
            'Me.grdList.DataSource = Me.item_defList
            Me.grdList.DataBind()
            modUtil.GridView2Excel("MMSInvoiceApply.xls", Me.grdList)
        End If
    End Sub



    '取得操作者區域
    'Private Function get_area(ByVal id As String) As String
    '    Dim tmpsql As String
    '    tmpsql = "select * from MMSEmployee where EmployeeNo='" + id + "'"
    '    Try
    '        get_area = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
    '    Catch ex As Exception
    '        get_area = "A"
    '    End Try
    'End Function


    'Private Function GetAreaCode() As String
    '    Dim tmpsql As String
    '    Dim STNID As String = Request.Cookies("STNID").Value
    '    tmpsql = "select * from MECHSTNM where STNID='" + STNID + "'"
    '    If get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "    " Or get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "4601" Then
    '        tmpsql = "Select * From MMSArea Where UNITCODE='" + get_DataTable(tmpsql).Rows(0).Item("STNUID") + "'"
    '        'Dim tmparea As String
    '        Try
    '            GetAreaCode = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
    '        Catch ex As Exception
    '            GetAreaCode = ""
    '        End Try
    '    Else
    '        GetAreaCode = ""
    '    End If
    'End Function
End Class
