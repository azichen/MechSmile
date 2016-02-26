Imports System.Data

Partial Class MMS_ReceiveQry
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            'modDB.SetGridViewStyle(Me.GridView4)          '* 套用gridview樣式

            '----------------------------------------* 設定GridView樣式
            modDB.SetGridViewStyle(Me.grdList, 7)
            Me.grdList.RowStyle.Height = 20
            modDB.SetFields("客戶代碼", "客戶代碼", grdList, False)
            modDB.SetFields("客戶名稱", "客戶名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("發票日期", "發票日期", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("發票號碼", "發票號碼", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("統一編號", "統一編號", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("傳票號碼", "傳票號碼", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("發票金額", "發票金額", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("已收款金額", "已收款金額", grdList, HorizontalAlign.Right, "", False)

            '--------------------------------------------------* 加入連結欄位
            Dim addColumn1 As New ButtonField()
            addColumn1.CommandName = "cmdSelect"
            addColumn1.Text = "選取"
            addColumn1.HeaderStyle.ForeColor = Drawing.Color.White
            addColumn1.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            addColumn1.ButtonType = ButtonType.Link
            addColumn1.HeaderText = "功能"
            grdList.Columns.Add(addColumn1)

            Me.lblMsg.Text = Request("msg")

            Me.txtCmpID.Attributes.Add("onkeypress", "return ToUpper()")

            Dim tmpsql As String
            tmpsql = "select * from MMSEmployee where EmployeeNo='" + My.User.Name + "' and Salesman='Y'"
            Dim tt As System.Data.DataTable
            tt = get_DataTable(tmpsql)
            If tt.Rows.Count > 0 Then

            End If

            Me.lblMsg.Text = Request("msg")
            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 7, Me.ViewState("EmptyTable"))
            Else
                'Call ReflashQryData() '* 移至責任區的DataBound
            End If

            modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)

            txtCmpID.Focus()
        End If


        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = Me.txtSql.Text
        '--------------------------------------------------* 設定GridView資料來源ID

        Me.ScriptManager1.RegisterPostBackControl(Me.btnExcel)
    End Sub



    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        Call ShowGrid()
        Session("iCurPage") = Nothing
        Me.lblMsg.Text = ""
        Me.GridView5.DataSource = Nothing
        Me.GridView3.DataSource = Nothing
        Me.GridView5.DataBind()
        Me.GridView3.DataBind()
    End Sub

    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        Dim sSql As String
        Dim Areacode As String = ""
        Areacode = modUnset.GetAreaCode(My.User.Name).Trim
        '
        sSql = "select "
        sSql += " a.CustomerNo as 客戶代碼,"
        sSql += " b.CustomerName as 客戶名稱,"
        sSql += " a.InvoiceDate as 發票日期,"
        sSql += " a.InvoiceNo as 發票號碼,"
        sSql += " a.GUINumber as 統一編號,"
        sSql += " D.DOCNO as 傳票號碼,"
        sSql += " REPLACE(convert(varchar,cast(a.Amount as money),-1),'.00','') as 發票金額,"
        sSql += " REPLACE(convert(varchar,cast(a.ReceiptAmount as money),-1),'.00','') as 已收款金額"
        sSql += " from MMSInvoiceM a"
        sSql += " left outer join MMSCustomers b on a.CustomerNo=b.CustomerNo "
        'sSql += " LEFT OUTER JOIN MMSVoucherNumber D ON A.InvoiceNo = D.InvoiceNo And substring(D.Accounts,1,4)<>'5128'" 'add by Azi 20140609
        sSql = sSql + " LEFT OUTER JOIN (Select DOCNO,INVOICENO,SUM(AMOUNT) AS AMOUNT FROM MMSVoucherNumber " _
                    + "      WHERE substring(Accounts,1,4)<>'5128' GROUP BY DOCNO,INVOICENO) D ON A.InvoiceNo = D.InvoiceNo " 'MODIFY by Azi 20140826
        sSql += " Where a.CustomerNo <> 'xxx' And a.status <>'V'"
        sSql += "   And a.status<>'R' "
        'add in 20140909
        If (Areacode <> "") And (Areacode <> "F") Then
            sSql += "and B.AREACODE = '" + Areacode + "'"
        End If
        '
        If txtCmpID.Text <> "" Then
            sSql += "and a.CustomerNo >= '" + Me.txtCmpID.Text + "'"
        End If
        If txtCmpID0.Text <> "" Then
            sSql += "and a.CustomerNo <= '" + Me.txtCmpID0.Text + "'"
        End If
        If txtSaleID2.Text <> "" Then
            sSql += "and b.CustomerName like '%" + Me.txtSaleID2.Text + "%'"
        End If
        If txtCmpID1.Text <> "" Then
            sSql += "and a.InvoiceNo like '%" + Me.txtCmpID1.Text + "%'"
        End If
        If Me.txtDtFrom.Text.Trim <> "" Then
            sSql += " and  a.InvoiceDate >= '" + Me.txtDtFrom.Text.Trim + "'"
        End If
        If Me.txtDtTo.Text.Trim <> "" Then
            sSql += " and  a.InvoiceDate <= '" + Me.txtDtTo.Text.Trim + "'"
        End If
        '
        If RadioButtonList1.SelectedIndex <> 0 Then
            If Me.RadioButtonList1.SelectedIndex = 1 Then
                sSql += " and a.Amount <> a.ReceiptAmount"
            Else
                sSql += " and a.Amount = a.ReceiptAmount"
            End If
        End If
        ''--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = sSql
        Me.txtSql.Text = sSql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        Me.dscList.SelectCommand = sSql
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        'Session("MainGrid") = get_DataTable(sSql)
        'Me.GridView4.DataSource = Session("MainGrid")
        'Me.GridView4.DataBind()
        'Me.GridView4.Columns(5).ItemStyle.HorizontalAlign = HorizontalAlign.Right
        'Me.GridView4.Columns(5).HeaderStyle.ForeColor = Drawing.Color.White
        '
        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 8, Me.ViewState("EmptyTable"))
            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            ViewState("Sql") = sSql '* 用於 Excel/換頁
        End If
    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        'modUtil.GridView2Excel("station.xls", Me.grdList)
        modUtil.GridView2Excel("MMSReceiveQry.xls", Me.grdList)
    End Sub


#Region "GridView Event"
    '******************************************************************************************************
    '* GridView_DataBound => 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.DataBound
        modDB.SetGridPageNum(Me.grdList, PagerButtons.NumericFirstLast, HorizontalAlign.Left)
    End Sub

    '******************************************************************************************************
    '* 保存瀏覽頁碼
    '******************************************************************************************************
    Protected Sub grdList_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.PageIndexChanged
        Me.ViewState("iCurPage") = grdList.PageIndex
    End Sub

    Protected Sub grdList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdList.RowDataBound
        modDB.SetGridLightPen(e)
    End Sub

    Protected Sub grdList_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles grdList.RowCommand
        Dim index As Integer = Convert.ToInt32(e.CommandArgument)
        Dim row As GridViewRow = grdList.Rows(index)
        'modDB.InsertSignRecord("AziTest", "dispReceive(" + row.Cells(3).Text.Trim + ")", My.User.Name) 'AZITEST
        If e.CommandName = "cmdSelect" Then
            dispReceive(row.Cells(3).Text.Trim)
        End If

    End Sub


    Private Sub dispReceive(ByVal docno As String)
        Dim tmpsql As String
        tmpsql = "select * from MMSInvoiceD where InvoiceNo='" + docno + "'"
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)

        Me.GridView5.DataSource = tt
        Me.GridView5.DataBind()
        Me.GridView5.Columns(1).HeaderStyle.ForeColor = Drawing.Color.White
        Me.GridView5.Columns(4).HeaderStyle.ForeColor = Drawing.Color.White

        tmpsql = "select "
        tmpsql += " b.DateOfReceipt as 收款日期,"
        tmpsql += " a.InvoiceNo as 發票號碼,"
        tmpsql += " REPLACE(convert(varchar,cast(a.AmountOfReceive  as money),-1),'.00','') as 沖銷金額"
        tmpsql += " from MMSReceive a"
        tmpsql += " left outer join MMSReceivablesM b on a.DocNo=b.ReceivablesNo "
        tmpsql += "  where a.InvoiceNo='" + docno + "'"
        tt = get_DataTable(tmpsql)

        Me.GridView3.DataSource = tt
        Me.GridView3.DataBind()
    End Sub

    Private Sub DispGridDown1(ByVal CustNo As String, ByVal docno As String)
        Dim tmpsql As String
        'Dim tr As DataRow
        tmpsql = "SELECT *,0 as revicea"
        tmpsql += " FROM SMILE_HQ.dbo.MMSInvoiceM where CustomerNo='" + CustNo + "'"
        tmpsql += " and (UnReceiptAmount<>0 or InvoiceNo in(select InvoiceNo from dbo.MMSReceive where DocNo='" + docno + "'))"
        tmpsql += " and Status='Y'"
        tmpsql += " and InvoiceNo<>''"
        Dim tt As DataTable = get_DataTable(tmpsql)
        Dim tt1 As DataTable
        tmpsql = "select * from MMSReceive where DocNo='" + docno + "'"
        tt1 = get_DataTable(tmpsql)
        Dim i, j As Integer
        For i = 0 To tt1.Rows.Count - 1
            For j = 0 To tt.Rows.Count - 1
                If tt.Rows(j).Item("InvoiceNo") = tt1.Rows(i).Item("InvoiceNo") Then
                    tt.Rows(j).Item("revicea") = tt1.Rows(i).Item("AmountOfReceive")
                    'Else
                    '    tt.Rows(j).Item("revicea") = 0
                End If
            Next
        Next
        For i = tt.Rows.Count - 1 To 0 Step -1
            If tt.Rows(i).Item("revicea") = 0 Then
                tt.Rows(i).Delete()
            End If
        Next
        Session("BatchGrid3") = tt
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
    End Sub

    Private Function SumDown() As Integer
        SumDown = 0
        Dim tt As DataTable
        tt = Session("BatchGrid3")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("revicea").ToString <> "" Then
                SumDown += tt.Rows(i).Item("revicea")
            End If
        Next

    End Function

#End Region

    '清除
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.txtCmpID.Text = ""
        Me.txtCmpID0.Text = ""
        Me.txtSaleID2.Text = ""
        Me.txtCmpID1.Text = ""
        Me.txtDtFrom.Text = ""
        Me.txtDtTo.Text = ""
        Me.lblMsg.Text = ""

        txtCmpID.Focus()
        'Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))
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

   

    'Protected Sub GridView4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.SelectedIndexChanged
    '    With Me.grdList.SelectedRow
    '        dispReceive(Me.grdList.SelectedRow.Cells(3).Text)
    '    End With
    'End Sub

    'Private Sub RePopulateValues()
    '    Dim index As String
    '    Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
    '    Try
    '        If categoryIDList.Count > 0 Then
    '            For Each row As GridViewRow In GridView4.Rows
    '                index = GridView4.DataKeys(row.RowIndex).Value
    '                If (categoryIDList.Contains(index)) Then
    '                    CType(row.FindControl("CheckBox2"), CheckBox).Checked = True
    '                Else
    '                    CType(row.FindControl("CheckBox2"), CheckBox).Checked = False
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        For Each row As GridViewRow In GridView4.Rows
    '            CType(row.FindControl("CheckBox2"), CheckBox).Checked = False
    '        Next
    '    End Try

    'End Sub


    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles grdList.PageIndexChanging
        'RememberOldValues()
        grdList.PageIndex = e.NewPageIndex
        Me.grdList.DataSource = Session("MainGrid")
        Me.grdList.DataBind()
        'RePopulateValues()
    End Sub

    Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles grdList.RowCancelingEdit
        Me.grdList.EditIndex = -1
        Me.grdList.DataSource = Session("MainGrid")
        Me.grdList.DataBind()
    End Sub

    '取得西元年月日
    Public Function GET_FW_DATE() As String
        GET_FW_DATE = LTrim(RTrim(Today.Year.ToString)) + "/"
        If Today.Month < 10 Then GET_FW_DATE += "0"
        GET_FW_DATE += LTrim(RTrim(Today.Month.ToString)) + "/"
        If Today.Day < 10 Then GET_FW_DATE += "0"
        GET_FW_DATE += LTrim(RTrim(Today.Day.ToString))
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

    Protected Sub txtCmpID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If txtCmpID0.Text.Trim = "" Then
            txtCmpID0.Text = txtCmpID.Text
        End If
        txtCmpID0.Focus()
    End Sub
End Class
