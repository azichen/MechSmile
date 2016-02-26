Imports System.Data

Partial Class MMS_InvoiceApply002
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
            dispArea()
            Me.DropDownList3.SelectedIndex = 0
            disp_emp()
            Me.Label1.Text = Format(Now, "yyyy/MM/dd")
            'If Request("src") = "menu" Then
            Session("CHECKED_ITEMS") = Nothing
            RePopulateValues()
            'End If
            Select Case Request("FormMode")
                Case "add"
                    DispGridUp()
                    Me.Button6.Visible = False
                    If GetAreaCode() <> "" Then
                        Me.DropDownList3.Enabled = False
                        Me.DropDownList3.SelectedValue = GetAreaCode()
                        disp_emp()
                    End If
                    Me.Button8.Visible = False
                    Me.Button9.Visible = False

                    Me.lblTitle.Text = "發票開立一般申請 - 新增"
                    Me.Button4.Text = "回主畫面"
                Case "edit"
                    dispDetail()
            End Select

            modDB.SetGridViewStyle(Me.GridView2)
            'Me.TextBox6.Attributes.Add("onkeypress", "return CheckKeyNumeric()")
            Me.TextBox6.Attributes.Add("onkeypress", "return CheckKeyNumericNeg()")
            Me.TextBox7.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TextBox8.Attributes.Add("onkeypress", "return CheckKeyNumber()")

            Me.TextBox1.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox3.Attributes.Add("onkeypress", "return CheckKeyNumber()")

            Me.DropDownList3.Focus()
            modUtil.SetObjReadOnly(Me, "TextBox8")
        End If
        modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)

        If ViewState("ROL").Substring(2, 1) = "N" Then Me.Button8.Visible = False '修改
        If ViewState("ROL").Substring(3, 1) = "N" Then Me.Button9.Visible = False '刪除
    End Sub

    Private Sub dispDetail()

        Me.lblTitle.Text = "發票開立一般申請 - 瀏覽"
        modUtil.SetObjReadOnly(Me, "txtDtTo")
        modUtil.SetObjReadOnly(Me, "TextBox1")
        modUtil.SetObjReadOnly(Me, "TextBox2")
        Me.Button2.Visible = False
        Me.Button3.Visible = False
        Me.DropDownList1.Visible = False
        Me.DropDownList3.Visible = False
        Me.Lable1.Visible = False
        Dim tt As DataTable
        Me.Panel1.Visible = False
        DispGridUp()
        Dim tmpsql As String
        'tmpsql = "Select A.CustomerNo,B.CustomerName,B.GUINumber,B.TitleOfInvoice,A.ReviewDate,A.Status,InvoiceType From MMSInvoiceApplyM A " _
        tmpsql = "Select A.CustomerNo,B.CustomerName,A.GUINumber,A.TitleOfInvoice,A.ReviewDate,A.Status,InvoiceType From MMSInvoiceApplyM A " _
               + " INNER JOIN MMSCUSTOMERS B ON A.CUSTOMERNO=B.CUSTOMERNO " _
               + " WHERE [DocNo]='" + Request("STNID").Trim + "'"
        tt = get_DataTable(tmpsql)
        'Try
        '    Me.RadioButtonList1.SelectedIndex = Int(tt.Rows(0).Item("InvoiceType")) - 1
        'Catch ex As Exception
        '    Me.RadioButtonList1.SelectedIndex = 0
        'End Try

        If IsDBNull(tt.Rows(0).Item("ReviewDate")) = False And tt.Rows(0).Item("ReviewDate") <> "" Then
            Me.Button1.Visible = False
            Me.Button8.Visible = False
        End If

        If tt.Rows(0).Item("Status") = "Y" Then
            Me.Button6.Visible = False
            Me.GridView2.Columns(6).Visible = False
            Me.Button9.Visible = False
        Else
            Me.Button8.Visible = True
            Me.Button9.Visible = True
        End If
        '
        If tt.Rows(0).Item("InvoiceType") = "2" Then
            RadioButtonList1.SelectedIndex = 1
        Else
            RadioButtonList1.SelectedIndex = 0
        End If

        Try
            Me.DropDownList3.SelectedValue = tt.Rows(0).Item("ApplyDepartment")
        Catch ex As Exception

        End Try

        Me.TextBox1.Text = tt.Rows(0).Item("CustomerNo")
        Me.TextBox2.Text = tt.Rows(0).Item("CustomerName")
        Me.TextBox3.Text = tt.Rows(0).Item("GUINumber")
        Me.TextBox4.Text = tt.Rows(0).Item("TitleOfInvoice")

        'tmpsql = "select 0 as checkf,Serial as Sn,ItemName,convert(money,UnitPrice) as UnitPrice,Quantity as Qty,Amount,Memo  as memo from MMSInvoiceApplyD"
        tmpsql = "select 0 as checkf,Serial as Sn,ItemName,convert(money,UnitPrice) as UnitPrice,Quantity as Qty,Amount,Memo  as memo,InvoicePeriod from MMSInvoiceApplyD"
        tmpsql += " where DocNo='" + Request("STNID").Trim + "'"
        tt = get_DataTable(tmpsql)
        Session("BatchGrid1") = tt
        Me.GridView2.DataSource = Session("BatchGrid1")
        Me.GridView2.DataBind()
        Me.Button1.Visible = False
        tmpsql = "select * from  MMSInvoiceApplyD where [DocNo]='" + Request("STNID").Trim + "'"
        tt = get_DataTable(tmpsql)
        Me.txtDtTo.Text = tt.Rows(0).Item("InvoiceDate")
        Me.TextBox11.Text = tt.Rows(0).Item("Memo")
        Me.Button4.Text = "回主畫面"
        Me.Button6.Visible = False
        Me.Panel3.Visible = False
        Me.RadioButtonList1.Enabled = False

        modUtil.SetObjReadOnly(Me, "txtDtTo")
        modUtil.SetObjReadOnly(Me, "TextBox3")
        modUtil.SetObjReadOnly(Me, "TextBox4")
        modUtil.SetObjReadOnly(Me, "TextBox11")
        Me.GridView2.Columns(5).Visible = False
        Me.GridView2.Columns(6).Visible = False
        Me.Image2.Visible = False
    End Sub



    Private Sub dispArea()
        Dim tmpsql As String
        tmpsql = "SELECT AreaCode, AreaCode+'-'+AreaName as  AreaName  FROM MMSArea where Effective ='Y' "
        'tmpsql += " and Effective='Y'"
        Me.SqlDataSource3.SelectCommand = tmpsql
        Me.SqlDataSource3.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.DropDownList3.DataSourceID = Me.SqlDataSource3.ID
        Me.DropDownList3.DataTextField = "AreaName"
        Me.DropDownList3.DataValueField = "AreaCode"
        Me.DropDownList3.DataBind()
    End Sub

    Private Sub disp_emp()
        Dim tmpsql As String
        tmpsql = "SELECT [EmployeeNo], [EmployeeNo]+'_'+[EmployeeName] as EmployeeName FROM [MMSEmployee] "
        tmpsql += " where Salesman='Y'"
        tmpsql += " and Effective='Y'"
        tmpsql += " and AreaCode='" + Me.DropDownList3.SelectedValue + "'"
        Me.SqlDataSource1.SelectCommand = tmpsql
        Me.SqlDataSource1.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.DropDownList1.DataSourceID = Me.SqlDataSource1.ID
        Me.DropDownList1.DataTextField = "EmployeeName"
        Me.DropDownList1.DataValueField = "EmployeeNo"
        Me.DropDownList1.DataBind()
        If Me.DropDownList1.Items.Count > 0 Then
            Me.DropDownList1.SelectedIndex = 0
            DispGrid()
        End If
    End Sub

    Private Sub DispGridUp()
        Dim tmpsql As String
        Dim tr As DataRow
        'tmpsql = "select 0 as checkf,0 as Sn,'' as ItemName,0.888 as UnitPrice,0 as Qty,0 as Amount,''  as memo"
        tmpsql = "select 0 as checkf,0 as Sn,'' as ItemName,0.888 as UnitPrice,0 as Qty,0 as Amount,''  as memo,'' as InvoicePeriod"
        Dim tt As DataTable = get_DataTable(tmpsql)
        tt.Rows.Clear()
        'tr = tt.NewRow
        'tr.Item("Sn") = tt.Rows.Count + 1
        'tt.Rows.Add(tr)
        Session("BatchGrid1") = tt
        Me.GridView2.DataSource = Session("BatchGrid1")
        Me.GridView2.DataBind()
    End Sub



    Private Sub update_grid1()
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowEditing(sender As Object, e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Me.GridView1.EditIndex = e.NewEditIndex
        update_grid1()
    End Sub

    Protected Sub GridView1_RowCancelingEdit(sender As Object, e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit
        Me.GridView1.EditIndex = -1
        update_grid1()
    End Sub

    Protected Sub GridView1_RowUpdating(sender As Object, e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        'Retrieve the table from the session object.
        Dim dt As DataTable = CType(Session("BatchGrid"), DataTable)

        'Update the values.
        Dim row = GridView1.Rows(Me.GridView1.EditIndex)
        dt.Rows(row.DataItemIndex)("ItemName") = (CType((row.Cells(7).Controls(0)), TextBox)).Text
        Me.GridView1.EditIndex = -1
        update_grid1()
    End Sub

    Private Sub DispGrid()
        Dim tmpsql As String
        tmpsql = "select 0 as checkf,a.CustomerNo,a.ContractNo,C.CustomerName,"
        tmpsql += " case a.InvoiceCycle"
        tmpsql += " when '1' then '年'"
        tmpsql += " when '2' then '季'"
        tmpsql += " when '3' then '半年'"
        tmpsql += " when '4' then '二月'"
        tmpsql += " when '5' then '月'"
        tmpsql += " End AS InvoiceCycle,"
        tmpsql += " B.InvoicePeriod, " 'add in 20140606 by Azi
        tmpsql += " REPLACE(CONVERT(varchar(20), (CAST(a.PeriodMaintenanceAmount AS money)), 1), '.00', '') as PeriodMaintenanceAmount"
        tmpsql += " ,case a.DatePlus "
        'tmpsql += " when '1' then b.[InvoicePeriod]+ItemName"
        'tmpsql += " when '2' then ItemName+b.[InvoicePeriod]"
        tmpsql += " when '1' then ItemName + CAST(CONVERT(INT,(SUBSTRING(B.INVOICEPERIOD,1,4)))-1911 AS VARCHAR(3)) + '年' + SUBSTRING(B.INVOICEPERIOD,5,2) + '月' "
        tmpsql += " when '2' then CAST(CONVERT(INT,(SUBSTRING(B.INVOICEPERIOD,1,4)))-1911 AS VARCHAR(3)) + '年' + SUBSTRING(B.INVOICEPERIOD,5,2) + '月' + ItemName "
        tmpsql += " when '3' then ItemName"
        tmpsql += " end as ItemName ,"
        tmpsql += " a.ContractNo+B.InvoicePeriod as tmp"
        tmpsql += " from MMSContract a"
        tmpsql += " INNER JOIN MMSCUSTOMERS C ON A.CUSTOMERNO = C.CUSTOMERNO"
        tmpsql += " left outer join MMSInvoiceA B on A.ContractNo = B.ContractNo"
        tmpsql += " where a.Salesman<>'XXXXX'" '+ Me.DropDownList1.SelectedValue + "'"
        'modify in 2014/11/14
        'tmpsql += " and a.CancelDate=''"
        tmpsql += " and ((A.CancelDate='') OR (B.InvoicePeriod < (SUBSTRING(A.CancelDate,1,4)+SUBSTRING(A.CancelDate,6,2)))) "
        tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "') and B.YN = 'N' "
        tmpsql += " and a.CustomerNo='" + Me.TextBox1.Text + "'"
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)
        Session("BatchGrid") = tt
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()
    End Sub

    '取得西元年月
    Public Function GET_W_DATE() As String
        Dim Dday As Date = DateAdd(DateInterval.Month, 2, Today)
        GET_W_DATE = LTrim(RTrim(Dday.Year.ToString))
        If Dday.Month < 10 Then GET_W_DATE += "0"
        GET_W_DATE += LTrim(RTrim(Dday.Month.ToString))
    End Function

    '取得西元年月
    Public Function GET_T_DATE() As String
        Dim Dday As Date = Today
        GET_T_DATE = LTrim(RTrim(Dday.Year.ToString))
        If Dday.Month < 10 Then GET_T_DATE += "0"
        GET_T_DATE += LTrim(RTrim(Dday.Month.ToString))
    End Function

    'Protected Sub DropDownList3_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList3.SelectedIndexChanged
    '    disp_emp()
    'End Sub

    '******************************************************************************************************
    '* GridView_RowDataBound => 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
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
            Case DataControlRowType.Header
                e.Row.Cells(6).Text = "每期保養<br>金額(含稅)" 'modify in 20140606
        End Select
    End Sub

    Protected Sub Button4_Click(sender As Object, e As System.EventArgs) Handles Button4.Click
        If Me.Button4.Text = "取消設定" Then
            dispDetail()
            Me.Button8.Visible = True
            Exit Sub
        End If
        If Me.Button4.Text = "回主畫面" Then
            Response.Redirect("InvoiceApply.aspx?Returnflag=1&msg=")
        End If

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

    '全選
    Protected Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Dim categoryIDList As ArrayList = New ArrayList
        Dim tt As DataTable
        Dim i As Integer
        Session("CHECKED_ITEMS") = Nothing
        tt = Session("BatchGrid")
        For i = 0 To tt.Rows.Count - 1
            categoryIDList.Add(tt.Rows(i).Item("CustomerNo"))
        Next
        Try
            If (categoryIDList.Count > 0) Then
                Session("CHECKED_ITEMS") = categoryIDList
            End If
        Catch ex As Exception

        End Try
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()
        RePopulateValues()
    End Sub

    Private Sub SelectAll()
        Dim categoryIDList As ArrayList = New ArrayList
        Dim tt As DataTable
        Dim i As Integer
        Session("CHECKED_ITEMS") = Nothing
        tt = Session("BatchGrid")
        For i = 0 To tt.Rows.Count - 1
            categoryIDList.Add(tt.Rows(i).Item("CustomerNo"))
        Next
        Try
            If (categoryIDList.Count > 0) Then
                Session("CHECKED_ITEMS") = categoryIDList
            End If
        Catch ex As Exception

        End Try
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()
        RePopulateValues()
    End Sub


    Protected Sub Button3_Click(sender As Object, e As System.EventArgs) Handles Button3.Click
        Session("CHECKED_ITEMS") = Nothing
        RePopulateValues()
    End Sub

    Private Sub RePopulateValues()
        Dim index As String
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Try
            If categoryIDList.Count > 0 Then
                For Each row As GridViewRow In GridView1.Rows
                    index = GridView1.DataKeys(row.RowIndex).Value
                    If (categoryIDList.Contains(index)) Then
                        CType(row.FindControl("CheckBox2"), CheckBox).Checked = True
                    Else
                        CType(row.FindControl("CheckBox2"), CheckBox).Checked = False
                    End If
                Next
            End If
        Catch ex As Exception
            For Each row As GridViewRow In GridView1.Rows
                CType(row.FindControl("CheckBox2"), CheckBox).Checked = False
            Next
        End Try

    End Sub

    Private Sub RememberOldValues()
        Dim categoryIDList As ArrayList = New ArrayList
        Dim index As String
        'If Session("CHECKED_ITEMS") Is Nothing = False Then
        '    categoryIDList = Session("CHECKED_ITEMS")
        'End If
        For Each row As GridViewRow In GridView1.Rows
            index = GridView1.DataKeys(row.RowIndex).Value
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
            'If CType(row.FindControl("CheckBox2"), CheckBox).Checked Then
            '    If (categoryIDList.Contains(index)) = False Then
            '        categoryIDList.Add(index)
            '    End If
            'Else
            '    categoryIDList.Remove(index)
            'End If
        Next

        Try
            If (categoryIDList.Count > 0) Then
                Session("CHECKED_ITEMS") = categoryIDList
            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub GridView1_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        RememberOldValues()
        GridView1.PageIndex = e.NewPageIndex
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()
        RePopulateValues()
    End Sub

    Private Function getDocNo(ByVal vInvoicePeriod As String) As String
        Dim tmpsql As String
        Dim tt As DataTable
        If vInvoicePeriod <> "" Then
            tmpsql = "select max(DocNo) from MMSInvoiceApplyM where DocNo like '" + vInvoicePeriod.Substring(0, 6) + "%'"
        Else
            tmpsql = "select max(DocNo) from MMSInvoiceApplyM where DocNo like '" + GET_T_DATE() + "%'"
        End If
        tt = get_DataTable(tmpsql)
        If IsDBNull(tt.Rows(0).Item(0)) Then
            If vInvoicePeriod <> "" Then
                getDocNo = vInvoicePeriod.Substring(0, 6) + "001"
            Else
                getDocNo = GET_T_DATE() + "001"
            End If
        Else
            getDocNo = Int(tt.Rows(0).Item(0)) + 1
        End If
    End Function

    Private Function CheckLineCount() As Boolean
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tt1 As DataTable
        Try
            If categoryIDList.Count + GridView2.Rows.Count > 5 Then
                CheckLineCount = False
            Else
                CheckLineCount = True
            End If
        Catch ex As Exception
            CheckLineCount = True
        End Try

    End Function


    '新增
    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        RememberOldValues()
        Dim tmpsql, vitemname As String
        Dim tt As DataTable
        Dim i As Integer
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        tt = Session("BatchGrid1")
        If categoryIDList Is Nothing And tt.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "尚未輸入任何發票內容項目!!")
            Exit Sub
        End If
        If RadioButtonList1.SelectedIndex = 0 And Me.TextBox11.Text.Trim.Length > 100 Then
            modUtil.showMsg(Me.Page, "訊息", "發票備註不可超過50中文字!!")
            Exit Sub
        End If
        If Me.CheckLineCount = False Then
            modUtil.showMsg(Me.Page, "訊息", "發票內容項目不可超過五筆!!")
            Exit Sub
        End If
        'SelectAll()

        Dim VGUINumber As String = Me.TextBox3.Text.Trim '20140806

        If VGUINumber <> "" Then
            If Not modUtil.BANCheck(VGUINumber) Then
                modUtil.showMsg(Me.Page, "訊息", "公司統編不合乎規則，麻煩再核對一次!!")
                Exit Sub
            End If
        End If

        Me.GridView2.EditIndex = -1
        tt = Session("BatchGrid1")
        '
        'modDB.InsertSignRecord("AziTest", "A BatchGrid1 counter=" & tt.Rows.Count.ToString, My.User.Name) 'AZITEST

        For i = tt.Rows.Count - 1 To 0 Step -1
            If tt.Rows(i).Item("ItemName").ToString.Trim = "" Then
                tt.Rows(i).Delete()
            End If
        Next

        'Me.UpdateGrid2()

        If Me.txtDtTo.Text.Trim = "" Then
            modUtil.showMsg(Me.Page, "訊息", "請輸入發票日期!!")
            Exit Sub
        End If
        'tmpsql = "select * from MMSCustomers where Salesman='" + Me.DropDownList1.SelectedValue + "' and CustomerNo='" + Me.TextBox1.Text.Trim + "'"
        tmpsql = "select * from MMSCustomers where  CustomerNo='" + Me.TextBox1.Text.Trim + "' "
        If Me.GetAreaCode <> "" Then
            tmpsql += "and AreaCode='" + Me.GetAreaCode + "'"
        End If
        tt = get_DataTable(tmpsql)
        If tt.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "您所輸入的客戶代號不存在或不屬於所屬區域!!")
            Me.TextBox1.Focus()
            Exit Sub
        End If

        Dim tdocno As String = getDocNo("") '取得單號

        tmpsql = "INSERT INTO MMSInvoiceApplyM([DocNo],[CustomerNo],[CustomerName],[ApplyDepartment],[GUINumber],[TitleOfInvoice]"
        tmpsql += " ,[Applicant],[ApplyDate],[ReviewDate],[Reviewers],[Status],[InvoiceType]) VALUES ("
        tmpsql += "'" + tdocno + "'"
        tmpsql += ",'" + Me.TextBox1.Text.Trim + "'" 'CustomerNo
        tmpsql += ",'" + Me.TextBox2.Text.Trim + "'" 'CustomerName
        tmpsql += ",'" + Request.Cookies("STNID").Value + "'"
        'tmpsql += ",'" + Me.TextBox3.Text.Trim + "'" 'GUINumber
        tmpsql += ",'" + VGUINumber + "'" 'GUINumber
        tmpsql += ",'" + Me.TextBox4.Text + "'" 'TitleOfInvoice
        tmpsql += ",'" + My.User.Name + "'"
        tmpsql += ",'" + Me.GET_FW_DATE + "'"
        tmpsql += ",''"
        tmpsql += ",''"
        tmpsql += ",'N'"
        tmpsql += ",'" + Str(Me.RadioButtonList1.SelectedIndex + 1).Trim + "'"
        tmpsql += " )"

        EXE_SQL(tmpsql)

        tt = Session("BatchGrid1")
        Dim VMEMO As String = Me.TextBox11.Text.Trim

        For i = 0 To tt.Rows.Count - 1
            'ADD IN 20141003 備註只存在 SERIAL=1 的資料中，避免存在每一筆品項中
            If i > 0 Then
                VMEMO = ""
            End If
            '
            tmpsql = "INSERT INTO MMSInvoiceApplyD([DocNo],[Serial],[DocType],[ContractNo],[InvoicePeriod],[PeriodMaintenanceAmount]"
            tmpsql += ",[ItemName],[UnitPrice],[Quantity],[Amount],[Memo],[InvoiceDate],[InvoiceNo],[Effective])"
            tmpsql += "VALUES("
            tmpsql += "'" + tdocno + "'"
            tmpsql += "," + (i + 1).ToString.Trim '項次
            tmpsql += ",'2'" 'DocType
            tmpsql += ",''" 'ContractNo
            tmpsql += ",'" + GET_T_DATE() + "'" 'InvoicePeriod
            tmpsql += ",0"  'PeriodMaintenanceAmount
            tmpsql += ",'" + tt.Rows(i).Item("ItemName") + "'" 'ItemName
            tmpsql += "," + tt.Rows(i).Item("UnitPrice").ToString 'UnitPrice
            tmpsql += "," + tt.Rows(i).Item("Qty").ToString 'Quantity
            tmpsql += "," + tt.Rows(i).Item("Amount").ToString 'Amount
            tmpsql += ",'" + VMEMO + "'"
            tmpsql += ",'" + Me.txtDtTo.Text + "'"
            tmpsql += ",''"
            tmpsql += ",'N'"
            tmpsql += ")"
            EXE_SQL(tmpsql)

        Next

        'GridView1

        ''未選擇任何合約
        'If Session("CHECKED_ITEMS") Is Nothing Then
        '    Response.Redirect("InvoiceApply.aspx?Returnflag=1&msg=一般申請成功!!")
        '    Exit Sub
        'End If

        Dim tt1 As DataTable
        'modDB.InsertSignRecord("AziTest", "norma_qty=" + categoryIDList.Count.ToString, My.User.Name) 'AZITEST 
        '
        'If categoryIDList.Count > 0 Then
        If Not (Session("CHECKED_ITEMS") Is Nothing) Then
            For i = 0 To categoryIDList.Count - 1
                'modDB.InsertSignRecord("AziTest", "categoryIDList.Item(" + i.ToString + ")= " & categoryIDList.Item(i), My.User.Name) 'AZITEST 
                'modDB.InsertSignRecord("AziTest", "categoryIDList.Item(" + i.ToString + "),1,14)= " & Mid(categoryIDList.Item(i), 1, 14), My.User.Name) 'AZITEST 

                'ADD IN 20141003 備註只存在 SERIAL=1 的資料中，避免存在每一筆品項中
                If i > 0 Then
                    VMEMO = ""
                End If

                tmpsql = "select * from MMSContract where ContractNo="
                tmpsql += "'" + Mid(categoryIDList.Item(i), 1, categoryIDList.Item(i).ToString.Length - 6) + "'"
                'tmpsql += "'" + categoryIDList.Item(i) + "'"

                tt1 = get_DataTable(tmpsql)

                vitemname = getDocItemName(Mid(categoryIDList.Item(i), 1, 14), Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6))
                '
                modDB.InsertSignRecord("(InvoiceApply002)", "發票申請 CONTRACTNO =" & tt1.Rows(0).Item("ContractNo") & " INVYM =" & Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6), My.User.Name)

                tmpsql = "INSERT INTO MMSInvoiceApplyD([DocNo],[Serial],[DocType],[ContractNo],[InvoicePeriod],[PeriodMaintenanceAmount]"
                tmpsql += ",[ItemName],[UnitPrice],[Quantity],[Amount],[Memo],[InvoiceDate],[InvoiceNo],[Effective])"
                tmpsql += "VALUES("
                tmpsql += "'" + tdocno + "'"
                tmpsql += "," + (tt.Rows.Count + i + 1).ToString  'Serial
                tmpsql += ",'1'"
                tmpsql += ",'" + tt1.Rows(0).Item("ContractNo") + "'"
                'tmpsql += ",'" + GET_W_DATE() + "'" '發票期別
                tmpsql += ",'" + Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6) + "'" '發票期別
                tmpsql += "," + tt1.Rows(0).Item("PeriodMaintenanceAmount").ToString
                'tmpsql += ",'" + getDocItemName(Mid(categoryIDList.Item(i), 1, 14), Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6)) + "'"
                tmpsql += ",'" + vitemname + "'"
                tmpsql += "," + tt1.Rows(0).Item("PeriodMaintenanceAmount").ToString
                tmpsql += ",1"
                tmpsql += "," + tt1.Rows(0).Item("PeriodMaintenanceAmount").ToString
                tmpsql += ",'" + VMEMO + "'"
                tmpsql += ",'" + Me.txtDtTo.Text + "'"
                tmpsql += ",''"
                tmpsql += ",'N'"
                tmpsql += ")"
                'modDB.InsertSignRecord("(InvoiceApply002)-InvoiceApplyD", "ContractNo=" + tt1.Rows(0).Item("ContractNo") + " AND InvoicePeriod=" + Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6), My.User.Name)

                EXE_SQL(tmpsql)

                tmpsql = "update MMSInvoiceA set YN='Y',updatetime=getdate()"
                tmpsql += " where ContractNo='" + tt1.Rows(0).Item("ContractNo") + "' AND InvoicePeriod='" + Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6) + "'"
                'modify in 20150123
                modDB.InsertSignRecord("(InvoiceApply002)-發票期別註記", "ContractNo=" + tt1.Rows(0).Item("ContractNo") + " AND InvoicePeriod=" + Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6), My.User.Name)

                'modDB.InsertSignRecord("A", "C serial=" + (tt.Rows.Count + i + 1).ToString, My.User.Name) 'AZITEST
                EXE_SQL(tmpsql)

                UpdateLastInvoiceDate(GET_W_DATE, tt1.Rows(0).Item("ContractNo"))
            Next
        End If
        Response.Redirect("InvoiceApply.aspx?Returnflag=1&msg=一般申請成功!!")
    End Sub

    Private Function getDocItemName(ByVal sn As String, ByVal IP As String) As String
        Dim tt As DataTable
        tt = Session("BatchGrid")
        
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            'modDB.InsertSignRecord("AziTest", "Rows(i)_ContractNo=" & tt.Rows(i).Item("ContractNo") & " Rows(i)_itemname=" & tt.Rows(i).Item("itemname"), My.User.Name) 'AZITEST 

            If (tt.Rows(i).Item("ContractNo") = sn) And (tt.Rows(i).Item("InvoicePeriod") = IP) Then
                If Not IsDBNull(tt.Rows(i).Item("ItemName")) Then
                    getDocItemName = tt.Rows(i).Item("ItemName")
                Else
                    getDocItemName = ""
                End If

                Exit For
            End If
        Next
    End Function

    '更新合約最後發票開立期別
    Private Sub UpdateLastInvoiceDate(ByVal dd As String, cno As String)
        Dim tmpsql As String
        tmpsql = "update MMSContract set LastInvoiceDate='" + dd + "' where ContractNo='" + cno + "'"
        EXE_SQL(tmpsql)
    End Sub



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



    '取得西元年月日
    Public Function GET_FW_DATE() As String
        GET_FW_DATE = LTrim(RTrim(Today.Year.ToString)) + "/"
        If Today.Month < 10 Then GET_FW_DATE += "0"
        GET_FW_DATE += LTrim(RTrim(Today.Month.ToString)) + "/"
        If Today.Day < 10 Then GET_FW_DATE += "0"
        GET_FW_DATE += LTrim(RTrim(Today.Day.ToString))
    End Function

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        DispGrid()
    End Sub

    Protected Sub TextBox1_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox1.TextChanged, DropDownList3.SelectedIndexChanged
        If Me.TextBox1.Text.Trim = "" Then Exit Sub
        Dim tmpsql As String
        Dim tt As DataTable
        'tmpsql = "select * from MMSCustomers where Salesman='" + Me.DropDownList1.SelectedValue + "' and CustomerNo='" + Me.TextBox1.Text.Trim + "'"
        tmpsql = "select * from MMSCustomers where  CustomerNo='" + Me.TextBox1.Text.Trim + "' "
        If Me.GetAreaCode <> "" Then
            tmpsql += "and AreaCode='" + Me.GetAreaCode + "'"
        Else
            tmpsql += "and AreaCode='" + Me.DropDownList3.SelectedValue + "'"
        End If
        tt = get_DataTable(tmpsql)
        If tt.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "您所輸入的客戶代號不存在或不屬於所屬區域!!")
            Me.TextBox1.Focus()
            Exit Sub
        End If
        Me.TextBox1.Text = tt.Rows(0).Item("CustomerNo")
        Me.TextBox2.Text = tt.Rows(0).Item("CustomerName")
        Me.TextBox3.Text = tt.Rows(0).Item("GUINumber")
        Me.TextBox4.Text = tt.Rows(0).Item("TitleOfInvoice")

        DispGrid()
    End Sub

    Private Sub countmoney()
        If Me.TextBox6.Text.Trim = "" Or Me.TextBox7.Text.Trim = "" Then
            Exit Sub
        End If
        Dim num1 As Double
        num1 = Me.TextBox6.Text
        'Me.TextBox8.Text = Math.Round(num1 * Int(Me.TextBox7.Text), 0)
        Me.TextBox8.Text = Math.Round(num1 * Int(Me.TextBox7.Text), 0, MidpointRounding.AwayFromZero)
        'modDB.InsertSignRecord("AziTest", "num1=" + num1.ToString + " qty=" + Int(Me.TextBox7.Text).ToString, My.User.Name) 'AZITEST
        Me.TextBox7.Focus()
    End Sub

    Private Sub clean_edit()
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        'Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""

    End Sub

    Protected Sub Button5_Click(sender As Object, e As System.EventArgs) Handles Button5.Click
        '檢查輸入資料
        countmoney()
        Dim tmpstr As String = ""
        If Me.TextBox5.Text.Trim = "" Then
            tmpstr += "請輸入品名!  \n"
        End If
        If Me.TextBox6.Text.Trim = "" Then
            tmpstr += "請輸入單價!  \n"
        End If
        If Me.TextBox7.Text.Trim = "" Then
            tmpstr += "請輸入數量!  \n"
        End If
        'add in 20151016
        If Me.TextBox5.Text.Trim.Length > 100 Then
            tmpstr += "品名長度不可超過100字元(50中文字)! \n"
        End If
        '
        If tmpstr <> "" Then
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
            Exit Sub
        End If

        Dim tt As DataTable
        Dim tr As DataRow
        tt = Session("BatchGrid1")
        If Me.TextBox10.Text.Trim = "" Then
            tr = tt.NewRow
            tr.Item(2) = Me.TextBox5.Text.Trim
            tr.Item(3) = Me.TextBox6.Text.Trim
            tr.Item(4) = Me.TextBox7.Text.Trim
            tr.Item(5) = Me.TextBox8.Text.Trim
            'tr.Item(6) = Me.TextBox9.Text.Trim
            tt.Rows.Add(tr)
            Dim i As Integer
            For i = 1 To tt.Rows.Count
                tt.Rows(i - 1).Item("Sn") = i
            Next
        Else
            Dim i As Integer
            For i = 0 To tt.Rows.Count - 1
                If tt.Rows(i).Item(1).ToString = Me.TextBox10.Text Then
                    tt.Rows(i).Item(2) = Me.TextBox5.Text.Trim
                    tt.Rows(i).Item(3) = Me.TextBox6.Text.Trim
                    tt.Rows(i).Item(4) = Me.TextBox7.Text.Trim
                    tt.Rows(i).Item(5) = Me.TextBox8.Text.Trim
                    'tt.Rows(i).Item(6) = Me.TextBox9.Text.Trim
                End If
            Next
        End If

        Session("BatchGrid1") = tt
        Me.GridView2.DataSource = Session("BatchGrid1")
        Me.GridView2.DataBind()
        clean_edit()
        RememberOldValues()
    End Sub

    Protected Sub GridView2_RowCancelingEdit(sender As Object, e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView2.RowCancelingEdit
        Me.GridView2.EditIndex = -1
        UpdateGrid2()
    End Sub

    Protected Sub GridView2_RowDeleted(sender As Object, e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView2.RowDeleted

    End Sub

    Protected Sub GridView2_RowDeleting(sender As Object, e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView2.RowDeleting
        Dim tt As DataTable
        Dim tr As DataRow
        tt = Session("BatchGrid1")
        Dim i As Integer
        For i = tt.Rows.Count - 1 To 0 Step -1
            If tt.Rows(i).Item(1).ToString = Me.GridView2.Rows(e.RowIndex).Cells(0).Text Then
                tt.Rows(i).Delete()
            End If
        Next
        If Me.Button6.Visible = False Then
            For i = 1 To tt.Rows.Count
                tt.Rows(i - 1).Item("Sn") = i
            Next
        End If
        tt.AcceptChanges()
        Session("BatchGrid1") = tt
        Me.UpdateGrid2()
    End Sub

    Protected Sub GridView2_RowEditing(sender As Object, e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView2.RowEditing

        'Me.GridView2.EditIndex = e.NewEditIndex
        UpdateGrid2()
    End Sub

    Protected Sub GridView2_RowUpdating(sender As Object, e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView2.RowUpdating
        'Retrieve the table from the session object.
        Dim dt As DataTable = CType(Session("BatchGrid1"), DataTable)

        'Update the values.
        Dim row = GridView2.Rows(Me.GridView2.EditIndex)
        If (CType((row.Cells(2).Controls(0)), TextBox)).Text.Trim = "" Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "請輸入品名!!")
            Exit Sub
        End If
        If IsNumeric((CType((row.Cells(2).Controls(0)), TextBox)).Text) = False Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "單價請輸入數字!!")
            Exit Sub
        End If
        If IsNumeric((CType((row.Cells(3).Controls(0)), TextBox)).Text) = False Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "數量請輸入數字!!")
            Exit Sub
        End If
        dt.Rows(row.DataItemIndex)("ItemName") = (CType((row.Cells(1).Controls(0)), TextBox)).Text
        dt.Rows(row.DataItemIndex)("UnitPrice") = (CType((row.Cells(2).Controls(0)), TextBox)).Text
        dt.Rows(row.DataItemIndex)("Qty") = (CType((row.Cells(3).Controls(0)), TextBox)).Text
        dt.Rows(row.DataItemIndex)("memo") = (CType((row.Cells(5).Controls(0)), TextBox)).Text

        dt.Rows(row.DataItemIndex)("Amount") = Int((CType((row.Cells(2).Controls(0)), TextBox)).Text) * Int((CType((row.Cells(3).Controls(0)), TextBox)).Text)

        'Reset the edit index.
        GridView2.EditIndex = -1

        'Bind data to the GridView control.
        UpdateGrid2()
    End Sub

    Private Sub UpdateGrid2()
        Me.GridView2.DataSource = Session("BatchGrid1")
        Me.GridView2.DataBind()
    End Sub

    '修改
    Protected Sub Button6_Click(sender As Object, e As System.EventArgs) Handles Button6.Click
        Dim tmpsql As String
        Dim tt As DataTable
        Dim i As Integer

        If RadioButtonList1.SelectedIndex = 0 And Me.TextBox11.Text.Trim.Length > 50 Then
            modUtil.showMsg(Me.Page, "訊息", "發票備註不可超過50字!!")
            Exit Sub
        End If

        Me.GridView2.EditIndex = -1
        tt = Session("BatchGrid1")
        For i = tt.Rows.Count - 1 To 0 Step -1
            If tt.Rows(i).Item("ItemName").ToString.Trim = "" Then
                tt.Rows(i).Delete()
            End If
        Next

        Dim tdocno As String = Request("STNID").Trim '取得單號

        tmpsql = "update MMSInvoiceApplyM set "
        tmpsql += " GUINumber='" + Me.TextBox3.Text + "'"
        tmpsql += ",TitleOfInvoice='" + Me.TextBox4.Text + "'"
        tmpsql += ",InvoiceType='" + Str(Me.RadioButtonList1.SelectedIndex + 1).Trim + "'"
        tmpsql += " where DocNo='" + tdocno + "'"
        EXE_SQL(tmpsql)

        tt = Session("BatchGrid1")

        Dim tts As String = ""

        'modDB.InsertSignRecord("AziTest", "button6 action!", My.User.Name) 'AZITEST

        For i = 0 To tt.Rows.Count - 1
            tmpsql = "update MMSInvoiceApplyD set "
            tmpsql += " ItemName='" + tt.Rows(i).Item("ItemName") + "'"
            tmpsql += ",UnitPrice=" + tt.Rows(i).Item("UnitPrice").ToString
            tmpsql += ",Quantity=" + tt.Rows(i).Item("Qty").ToString
            tmpsql += ",Amount=" + tt.Rows(i).Item("Amount").ToString
            tmpsql += ",Memo='" + Me.TextBox11.Text + "'"
            tmpsql += ",InvoiceDate='" + Me.txtDtTo.Text + "'" '20150903 財務雅芬要求加入
            tmpsql += " where DocNo='" + tdocno + "'"
            tmpsql += " and Serial=" + tt.Rows(i).Item("Sn").ToString

            'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "Serial=" + tt.Rows(i).Item("Sn").ToString, My.User.Name) 'AZITEST

            EXE_SQL(tmpsql)
            If tts <> "" Then
                tts += ","
            End If
            tts += tt.Rows(i).Item("Sn").ToString
        Next

        tmpsql = "delete from MMSInvoiceApplyD "
        tmpsql += " where DocNo='" + tdocno + "'"
        tmpsql += " and Serial not in(" + tts + ")"
        EXE_SQL(tmpsql)

        Me.dispDetail()
        'Response.Redirect("InvoiceApply.aspx?Returnflag=1&msg=更新成功!!")
    End Sub

    Protected Sub TextBox6_TextChanged(sender As Object, e As System.EventArgs) Handles TextBox6.TextChanged, TextBox7.TextChanged
        Me.TextBox8.Text = 0
        countmoney()
        Me.TextBox7.Focus()
    End Sub

    Protected Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        Me.clean_edit()
    End Sub

    Protected Sub GridView2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles GridView2.SelectedIndexChanged
        Me.TextBox5.Text = Me.GridView2.Rows(Me.GridView2.SelectedIndex).Cells(1).Text
        Me.TextBox6.Text = Me.GridView2.Rows(Me.GridView2.SelectedIndex).Cells(2).Text
        Me.TextBox7.Text = Me.GridView2.Rows(Me.GridView2.SelectedIndex).Cells(3).Text
        Me.TextBox8.Text = Me.GridView2.Rows(Me.GridView2.SelectedIndex).Cells(4).Text
        'Me.TextBox9.Text = Me.GridView2.Rows(Me.GridView2.SelectedIndex).Cells(5).Text.Trim
        'If Me.TextBox9.Text = "&nbsp;" Then
        '    Me.TextBox9.Text = ""
        'End If
        Me.TextBox10.Text = Me.GridView2.Rows(Me.GridView2.SelectedIndex).Cells(0).Text
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


    Private Function GetAreaCode() As String
        Dim tmpsql As String
        Dim STNID As String = Request.Cookies("STNID").Value
        tmpsql = "select * from MECHSTNM where STNID='" + STNID + "'"
        If get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "    " Or get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "4601" Then
            tmpsql = "select * from MMSArea where UnitCode='" + get_DataTable(tmpsql).Rows(0).Item("STNUID") + "'"
            Dim tmparea As String
            Try
                GetAreaCode = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
            Catch ex As Exception
                GetAreaCode = ""
            End Try
        Else
            GetAreaCode = ""
        End If
        'GetAreaCode = "C"
    End Function

    Protected Sub Button8_Click(sender As Object, e As System.EventArgs) Handles Button8.Click
        '20150303 把修改關閉，因為修改送出的程式不齊全。要修改的話先刪後新增
        modUtil.SetObjEnabled(txtDtTo)
        modUtil.SetObjEnabled(TextBox3)
        modUtil.SetObjEnabled(TextBox4)
        modUtil.SetObjEnabled(TextBox11)

        Me.RadioButtonList1.Enabled = True
        Me.GridView2.Columns(5).Visible = True
        Me.GridView2.Columns(6).Visible = True
        Me.Image2.Visible = True
        Me.Panel3.Visible = True
        Me.Button4.Text = "取消設定"
        Me.Button9.Visible = False
        Me.Button6.Visible = True
        Me.Button8.Visible = False
        Me.lblTitle.Text = "發票開立一般申請 - 修改"
    End Sub

    Protected Sub Button9_Click(sender As Object, e As System.EventArgs) Handles Button9.Click
        Dim tmpsql As String
        Dim tt As DataTable ', tt1
        Dim i As Integer
        Dim vContractNo, vInvoicePeriod As String
        vContractNo = ""
        vInvoicePeriod = ""

        Me.GridView2.EditIndex = -1
        tt = Session("BatchGrid1")
        For i = tt.Rows.Count - 1 To 0 Step -1
            If tt.Rows(i).Item("ItemName").ToString.Trim = "" Then
                tt.Rows(i).Delete()
            End If
        Next

        Dim tdocno As String = Request("STNID").Trim '取得單號

        '20140109
        tmpsql = "select ContractNo,InvoicePeriod from dbo.MMSInvoiceApplyD where DocNo='" + tdocno + "'"
        tt = get_DataTable(tmpsql)
        For i = 0 To tt.Rows.Count - 1
            vContractNo = tt.Rows(i).Item("ContractNo")
            vInvoicePeriod = tt.Rows(i).Item("InvoicePeriod")

            modDB.InsertSignRecord("(InvoiceApply002)", "發票申請-刪除: DOCNO=" + tdocno + " ContractNo=" + vContractNo + ",InvoicePeriod=" + vInvoicePeriod, My.User.Name)

            If (vContractNo <> "") And (vInvoicePeriod <> "") Then
                '20150618 增加如果有重覆資料月保養資料，則不作
                tmpsql = "Select COUNT(*) AS RECQTY from MMSInvoiceApplyM A"
                tmpsql += " INNER JOIN MMSInvoiceApplyD B ON A.DOCNO=B.DOCNO"
                tmpsql += " Where B.ContractNo='" + vContractNo + "' AND B.InvoicePeriod='" + vInvoicePeriod + "'"
                tmpsql += "   And A.DOCNO<>'" + tdocno + "'"
                '
                'tt1 = get_DataTable(tmpsql)
                If (get_DataTable(tmpsql).Rows(0).Item("RECQTY") <= 0) Then
                    tmpsql = ""
                    tmpsql = "update MMSInvoiceA"
                    tmpsql += " set [YN]='N',updatetime=getdate()"
                    tmpsql += " where ContractNo='" + vContractNo + "' AND InvoicePeriod='" + vInvoicePeriod + "'"
                    'tmpsql += " where [ContractNo]+[InvoicePeriod] "
                    'tmpsql += " in(select [ContractNo]+[InvoicePeriod] from dbo.MMSInvoiceApplyD where  DocNo='" + tdocno + "')"
                    EXE_SQL(tmpsql)
                Else
                    modDB.InsertSignRecord("(InvoiceApply002)", "發票申請-刪除:合約_年月=" + vContractNo + "_" + vInvoicePeriod + "資料重複，註記不作處理", My.User.Name)
                End If
            End If
        Next
        modDB.InsertSignRecord("(InvoiceApply002)", "發票申請-刪除ApplyM : " + tdocno, My.User.Name)

        tmpsql = "delete from  MMSInvoiceApplyM "
        tmpsql += " where DocNo='" + tdocno + "'"
        EXE_SQL(tmpsql)

        modDB.InsertSignRecord("(InvoiceApply002)", "發票申請-刪除ApplyD : " + tdocno, My.User.Name)
        tmpsql = "delete from MMSInvoiceApplyD "
        tmpsql += " where DocNo='" + tdocno + "'"
        EXE_SQL(tmpsql)

        Response.Redirect("InvoiceApply.aspx?Returnflag=1&msg=刪除成功!!")
    End Sub
End Class
