Imports System.Data
Partial Class MMS_ReceiveReview
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
                If ViewState("ROL").Substring(1, 1) = "N" Then Me.Button3.Visible = False
            End If

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            'modDB.SetGridViewStyle(Me.GridView4)          '* 套用gridview樣式

            '--------------------------------------------------* 加入連結欄位
            Dim sField As New CommandField
            sField.ShowSelectButton = True
            sField.HeaderText = "功能"
            sField.SelectText = "瀏覽"
            sField.HeaderStyle.ForeColor = Drawing.Color.White
            sField.ItemStyle.HorizontalAlign = HorizontalAlign.Center

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
            modUtil.SetDateObj(Me.txtDtFr2, False, Me.txtDtTo2, False)
            modUtil.SetDateObj(Me.txtDtTo2, False, Nothing, False)
            disp_emp1()
            Me.Panel1.Visible = False
        End If


        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = Me.txtSql.Text
        '--------------------------------------------------* 設定GridView資料來源ID

        Me.ScriptManager1.RegisterPostBackControl(Me.btnExcel)
    End Sub

    Private Sub disp_emp1()
        Dim tmpsql As String
        tmpsql = "SELECT [EmployeeNo], [EmployeeNo]+'_'+[EmployeeName] as EmployeeName FROM [MMSEmployee] "
        tmpsql += " where Cashier='Y'"
        tmpsql += " and Effective='Y'"
        Session("dr1") = get_DataTable(tmpsql)
        Me.DropDownList1.DataSource = Session("dr1")
        Me.DropDownList1.DataTextField = "EmployeeName"
        Me.DropDownList1.DataValueField = "EmployeeNo"
        Me.DropDownList1.DataBind()
        If Me.DropDownList1.Items.Count > 0 Then
            Me.DropDownList1.SelectedIndex = 0
        End If
    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        Session("CHECKED_ITEMS") = Nothing 'Add in 20150209
        Call ShowGrid()
        Me.lblMsg.Text = ""
    End Sub

    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        'modDB.InsertSignRecord("AziTest", "ok-3", My.User.Name) 'AZITEST
        Panel1.Visible = False
        Dim sSql As String
        sSql = "SELECT "
        sSql += "0 as checkf,CASE STATUS WHEN 'Y' THEN REVIEWDATE ELSE '' END AS REVIEWDATE,"
        sSql += " *,Case Status when 'Y' then '審核通過' when 'N' then '未審核' end as Statusn"
        sSql += " ,CASE STATUS WHEN 'Y' THEN NAME ELSE '' END AS ReviewerName,(SalesDiscounts+Other) as OtherAmt"
        sSql += " From MMSReceivablesM A"
        sSql += " LEFT JOIN USER_PRO B ON A.Reviewers=B.ACCOUNT"
        sSql += " where [ReceivablesNo] <>'x' "
        If Me.txtSaleID1.Text.Trim <> "" Then
            sSql += " and  [Cashier] = '" + Me.txtSaleID1.Text.Trim + "'"
        End If
        If Me.txtCmpID.Text.Trim <> "" Then
            sSql += " and  [CustomerNo] = '" + Me.txtCmpID.Text.Trim + "'"
        End If
        If Me.txtDtFrom.Text.Trim <> "" Then
            sSql += " and  [DocDate] >= '" + Me.txtDtFrom.Text.Trim + "'"
        End If
        If Me.txtDtTo.Text.Trim <> "" Then
            sSql += " and  [DocDate] <= '" + Me.txtDtTo.Text.Trim + "'"
        End If
        'add in 20150205
        If Me.txtDtFr2.Text.Trim <> "" Then
            sSql += " and  [Reviewdate] >= '" + Me.txtDtFr2.Text.Trim + "'"
        End If
        If Me.txtDtTo2.Text.Trim <> "" Then
            sSql += " and  [Reviewdate] <= '" + Me.txtDtTo2.Text.Trim + "'"
        End If
        'ADD IN 20140606 BY AZI
        Button1.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Select Case Me.RadioButtonList1.SelectedIndex
            Case 0 '全部
                'sSql += " and InvoiceNo = ''"
            Case 1 '已審核
                sSql += " and Status ='Y' "
                Button1.Visible = True
                Button2.Visible = True
                Button4.Visible = True
                'sSql += " and (REVIEWDATE>'2000/01/01') "
                'sSql += " and (select count(*) from MMSInvoicePrint b where b.InvoiceNo=a.InvoiceNo) > 0"
            Case 2 '未審核
                sSql += " and ((Status <>'Y') or (Status is null)) " ' MODIFY BY AZI
                Button1.Visible = True
                Button2.Visible = True
                Button3.Visible = True
                'sSql += " and (select count(*) from MMSInvoicePrint b where b.InvoiceNo=a.InvoiceNo) = 0"
        End Select
        'add in 20141006 付款別
        If (CheckBox1.Checked) Or (CheckBox3.Checked) Or (CheckBox4.Checked) Or (CheckBox5.Checked) Then
            sSql += " AND ((1=2) " '
            If (CheckBox1.Checked) Then
                sSql += " OR (CASH>0)"
            End If

            If (CheckBox3.Checked) Then
                sSql += " OR (AMOUNTOFREMITTANCES>0) "
            End If
            '
            If (CheckBox4.Checked) Then
                sSql += " OR (AMOUNTOFCHECK>0) "
            End If
            '
            If (CheckBox5.Checked) Then
                sSql += " OR (SalesDiscounts>0) OR (Other>0)"
            End If
            'Else
            '    If (CheckBox3.Checked) Then
            '        sSql += " (AMOUNTOFREMITTANCES>0) "
            '        If (CheckBox4.Checked) Then
            '            sSql += " OR (AMOUNTOFCHECK>0) "
            '        End If
            '    Else
            '        sSql += " (AMOUNTOFCHECK>0) "
            '    End If
            'End If
            sSql += " ) " '
        Else
            sSql += " AND (1=2) " '
        End If
        '
        'If Me.txtSaleID1.Text.Trim <> "" Then
        '    sSql += " and [CustomerNo] in (select [CustomerNo] from [MMSContract] where [Salesman]='" + Me.txtSaleID1.Text + "' )"
        'End If
        ''--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.dscList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.dscList.SelectCommand = sSql
        Me.txtSql.Text = sSql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID

        Session("MainGrid") = get_DataTable(sSql)
        Me.GridView4.DataSource = Session("MainGrid")
        Me.GridView4.DataBind()
        'dispReceive("") 'azitest
        'dispReceive(Me.GridView4.SelectedRow.Cells(1).Text)

        If Me.GridView4.Rows.Count = 0 Then
            'Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))
            Me.Button1.Visible = False
            Me.Button2.Visible = False
            Me.Button3.Visible = False
            Me.btnExcel.Enabled = False
            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            'Me.Button1.Visible = True
            'Me.Button2.Visible = True
            'Me.Button3.Visible = True
        End If

    End Sub


#Region "GridView Event"
    '******************************************************************************************************
    '* GridView_DataBound => 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView4.DataBound
        modDB.SetGridPageNum(Me.GridView4, PagerButtons.NumericFirstLast, HorizontalAlign.Left)

    End Sub

    '******************************************************************************************************
    '* GridView_RowDataBound => 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub grdList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView4.RowDataBound
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


    Private Sub dispReceive(ByVal docno As String)
        Dim tmpsql As String
        tmpsql = "select * from MMSReceivablesM where ReceivablesNo='" + docno + "'"
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)
        Me.txtDtFrom2.Text = tt.Rows(0).Item("DateOfReceipt")
        Me.Label3.Text = tt.Rows(0).Item("DocDate")
        If tt.Rows(0).Item("Status") = "N" Then
            Me.Label5.Text = "未審核"
        Else
            Me.Label5.Text = "審核通過"
        End If
        Me.txtDtFrom0.Text = tt.Rows(0).Item("CustomerNo")
        Me.Label6.Text = tt.Rows(0).Item("CustomerName")
        DispGridUp(docno)
        Me.TextBox2.Text = tt.Rows(0).Item("Cash")
        Me.TextBox3.Text = tt.Rows(0).Item("AmountOfRemittances")
        Me.TextBox4.Text = tt.Rows(0).Item("RemittanceFee")
        Me.TextBox5.Text = tt.Rows(0).Item("SalesDiscounts")
        Me.TextBox6.Text = tt.Rows(0).Item("SalesTax")
        Me.TextBox7.Text = tt.Rows(0).Item("Other")
        Me.txtDtFrom1.Text = tt.Rows(0).Item("RemittanceDate")
        Me.Label7.Text = SumTop()
        Me.Label8.Text = tt.Rows(0).Item("RemittanceDate")
        Me.DropDownList1.SelectedIndex = -1
        DropDownList1.Items.FindByValue(tt.Rows(0).Item("Cashier")).Selected = True
        Me.TextBox8.Text = tt.Rows(0).Item("Memo")
        Me.Label4.Visible = True
        Me.Label5.Visible = True

        DispGridDown1(Me.txtDtFrom0.Text, docno)
    End Sub

    Private Function SumTop() As Integer
        SumTop = 0
        Dim tt As DataTable
        tt = Session("BatchGrid1")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If IsDBNull(tt.Rows(i).Item("AmountOfCheck")) = False Then
                If tt.Rows(i).Item("AmountOfCheck").ToString <> "" Then
                    SumTop += tt.Rows(i).Item("AmountOfCheck")
                End If
            End If
        Next
        If Me.TextBox2.Text <> "" Then
            SumTop += Int(Me.TextBox2.Text)
        Else
            Me.TextBox2.Text = 0
        End If
        If Me.TextBox3.Text <> "" Then
            SumTop += Int(Me.TextBox3.Text)
        Else
            Me.TextBox3.Text = 0
        End If
        If Me.TextBox4.Text <> "" Then
            SumTop += Int(Me.TextBox4.Text)
        Else
            Me.TextBox4.Text = 0
        End If
        If Me.TextBox5.Text <> "" Then
            SumTop += Int(Me.TextBox5.Text)
        Else
            Me.TextBox5.Text = 0
        End If
        If Me.TextBox6.Text <> "" Then
            SumTop += Int(Me.TextBox6.Text)
        Else
            Me.TextBox6.Text = 0
        End If
        If Me.TextBox7.Text <> "" Then
            SumTop += Int(Me.TextBox7.Text)
        Else
            Me.TextBox7.Text = 0
        End If
    End Function


    Private Sub DispGridDown1(ByVal CustNo As String, ByVal docno As String)
        Dim tmpsql As String
        Dim tr As DataRow
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

    Private Sub DispGridUp(ByVal ChequeNo As String)
        Dim tmpsql As String
        Dim tr As DataRow
        tmpsql = "SELECT Serial,ChequeNo,AmountOfCheck,PayingBank,BankAccount,ChequeExpiryDate"
        tmpsql += " FROM SMILE_HQ.dbo.MMSCheque where DocNo='" + ChequeNo + "'"
        Dim tt As DataTable = get_DataTable(tmpsql)
        If ChequeNo = "" Then
            tt.Rows.Clear()
            tr = tt.NewRow
            tr.Item("Serial") = tt.Rows.Count + 1
            tt.Rows.Add(tr)
        End If

        Session("BatchGrid1") = tt
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
    End Sub

    Private Sub disp_emp()
        Dim tmpsql As String
        tmpsql = "SELECT [EmployeeNo], [EmployeeNo]+'_'+[EmployeeName] as EmployeeName FROM [MMSEmployee] "
        tmpsql += " where Cashier='Y'"
        tmpsql += " and Effective='Y'"
        Session("dr1") = get_DataTable(tmpsql)
        Me.DropDownList1.DataSource = Session("dr1")
        Me.DropDownList1.DataTextField = "EmployeeName"
        Me.DropDownList1.DataValueField = "EmployeeNo"
        Me.DropDownList1.DataBind()
        If Me.DropDownList1.Items.Count > 0 Then
            Me.DropDownList1.SelectedIndex = 0
        End If
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

        'Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))
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

    Protected Sub GridView4_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles GridView4.SelectedIndexChanged
        With Me.GridView4.SelectedRow
            Me.Panel1.Visible = True
            dispReceive(Me.GridView4.SelectedRow.Cells(1).Text)
            Me.Label10.Text = Request("STNID").Trim
        End With
    End Sub

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        Dim categoryIDList As ArrayList = New ArrayList
        Dim tt As DataTable
        Dim i As Integer
        Session("CHECKED_ITEMS") = Nothing
        tt = Session("MainGrid")
        For i = 0 To tt.Rows.Count - 1
            categoryIDList.Add(tt.Rows(i).Item("ReceivablesNo").ToString)
        Next
        Try
            If (categoryIDList.Count > 0) Then
                Session("CHECKED_ITEMS") = categoryIDList
            End If
        Catch ex As Exception

        End Try
        Me.GridView4.DataSource = Session("MainGrid")
        Me.GridView4.DataBind()
        RePopulateValues()
    End Sub

    Private Sub RePopulateValues()
        Dim index As String
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Try
            If categoryIDList.Count > 0 Then
                For Each row As GridViewRow In GridView4.Rows
                    index = GridView4.DataKeys(row.RowIndex).Value
                    If (categoryIDList.Contains(index)) Then
                        CType(row.FindControl("CheckBox2"), CheckBox).Checked = True
                    Else
                        CType(row.FindControl("CheckBox2"), CheckBox).Checked = False
                    End If
                Next
            End If
        Catch ex As Exception
            For Each row As GridViewRow In GridView4.Rows
                CType(row.FindControl("CheckBox2"), CheckBox).Checked = False
            Next
        End Try

    End Sub

    Protected Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Session("CHECKED_ITEMS") = Nothing
        RePopulateValues()
    End Sub

    Protected Sub GridView1_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView4.PageIndexChanging
        RememberOldValues()
        GridView4.PageIndex = e.NewPageIndex
        Me.GridView4.DataSource = Session("MainGrid")
        Me.GridView4.DataBind()
        RePopulateValues()
    End Sub

    Protected Sub GridView1_RowCancelingEdit(sender As Object, e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView4.RowCancelingEdit
        Me.GridView4.EditIndex = -1
        Me.GridView4.DataSource = Session("MainGrid")
        Me.GridView4.DataBind()

    End Sub

    Private Sub RememberOldValues()
        Dim categoryIDList As ArrayList = New ArrayList
        Dim index As String
        For Each row As GridViewRow In GridView4.Rows
            index = GridView4.DataKeys(row.RowIndex).Value
            Dim result As Boolean = CType(row.FindControl("CheckBox2"), CheckBox).Checked
            If Session("CHECKED_ITEMS") Is Nothing = False Then
                categoryIDList = Session("CHECKED_ITEMS")
            End If
            If result Then
                If (categoryIDList.Contains(index)) = False Then
                    categoryIDList.Add(index)
                End If
            Else
                categoryIDList.Remove(index)
            End If
        Next
        Try
            If (categoryIDList.Count > 0) Then
                Session("CHECKED_ITEMS") = categoryIDList
            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub Button3_Click(sender As Object, e As System.EventArgs) Handles Button3.Click
        RememberOldValues()
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何資料!!")
            Exit Sub
        End If
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i As Integer
        Dim tt As DataTable
        Dim tmpstr As String = ""

        Dim docnoF As String = ""
        Dim docno As String
        For i = 0 To categoryIDList.Count - 1
            tmpsql = "update MMSReceivablesM set ReviewDate='" + GET_FW_DATE() + "',Reviewers='" + My.User.Name + "'"
            tmpsql += ",Status='Y'"
            tmpsql += " where ReceivablesNo='" + categoryIDList.Item(i) + "'"
            EXE_SQL(tmpsql)
        Next
        Response.Redirect("ReceiveReview.aspx?Returnflag=1&msg=申請審核成功!!")
    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        RememberOldValues()
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何資料!!")
            Exit Sub
        End If
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i As Integer
        Dim tt As DataTable
        Dim tmpstr As String = ""
        Dim CHKFLAG As String = ""
        'Dim docnoF As String = ""
        'Dim tt As System.Data.DataTable

        'Dim docno As String
        '取消審核的動作
        '檢查 是否已有傳票資料 
        For i = 0 To categoryIDList.Count - 1
            tmpsql = "select InvoiceNo from MMSVoucherNumber " _
                   + " where InvoiceNo in " _
                   + " (Select InvoiceNo from MMSReceive where DocNo='" + categoryIDList.Item(i) + "')"
            tt = get_DataTable(tmpsql)
            '
            If tt.Rows.Count = 0 Then
                tmpsql = "update MMSReceivablesM set ReviewDate='" + GET_FW_DATE() + "',Reviewers='" + My.User.Name + "'"
                tmpsql += ",Status='N'" '取消審核
                tmpsql += " where ReceivablesNo='" + categoryIDList.Item(i) + "'"
                EXE_SQL(tmpsql)
            Else
                CHKFLAG = "N"
            End If
        Next
        '
        If CHKFLAG = "N" Then
            Response.Redirect("ReceiveReview.aspx?Returnflag=1&msg=取消審核完成，但有部分資料已有傳票，不可取消，請確認!!")
        Else
            Response.Redirect("ReceiveReview.aspx?Returnflag=1&msg=取消審核成功!!")
        End If

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
End Class
