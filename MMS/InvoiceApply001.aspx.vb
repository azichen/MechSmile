Imports System.Data

Partial Class MMS_InvoiceApply001
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
            Me.DropDownList3.Focus()

            Session("CHECKED_ITEMS") = Nothing
            If GetAreaCode() <> "" Then
                Me.DropDownList3.Enabled = False
                Me.DropDownList3.SelectedValue = GetAreaCode()
                disp_emp()
            End If
        End If
        modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
    End Sub

    Private Sub dispArea()
        Dim tmpsql As String
        tmpsql = "SELECT AreaCode, AreaCode+'-'+AreaName as  AreaName  FROM MMSArea where Effective ='Y' "
        tmpsql += " and Effective='Y'"
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

    Private Sub DispGrid()
        Dim tmpsql As String
        tmpsql = "select 0 as checkf,a.CustomerNo,a.ContractNo,a.CustomerName,substring(b.[InvoicePeriod],1,4)+'/'+substring(b.[InvoicePeriod],5,2) as InvoicePeriod,"
        tmpsql += " case a.InvoiceCycle"
        tmpsql += " when '1' then '年'"
        tmpsql += " when '2' then '季'"
        tmpsql += " when '3' then '半年'"
        tmpsql += " when '4' then '二月'"
        tmpsql += " when '5' then '月'"
        tmpsql += " when '6' then '二年'"
        tmpsql += " End AS InvoiceCycle,"
        tmpsql += " REPLACE(CONVERT(varchar(20), (CAST(a.PeriodMaintenanceAmount AS money)), 1), '.00', '') as PeriodMaintenanceAmount  "
        tmpsql += " ,case a.DatePlus "
        'tmpsql += " when '1' then ItemName+b.[InvoicePeriod]"
        'tmpsql += " when '2' then b.[InvoicePeriod]+ItemName"
        tmpsql += " when '1' then ItemName + CAST(CONVERT(INT,(SUBSTRING(B.INVOICEPERIOD,1,4)))-1911 AS VARCHAR(3)) + '年' + SUBSTRING(B.INVOICEPERIOD,5,2) + '月' "
        tmpsql += " when '2' then CAST(CONVERT(INT,(SUBSTRING(B.INVOICEPERIOD,1,4)))-1911 AS VARCHAR(3)) + '年' + SUBSTRING(B.INVOICEPERIOD,5,2) + '月' + ItemName "
        tmpsql += " when '3' then ItemName"
        tmpsql += " end as ItemName, '' as Memo, "
        tmpsql += " a.ContractNo + b.[InvoicePeriod] as pk"
        tmpsql += " from MMSContract a"
        tmpsql += " left outer join MMSInvoiceA B on A.ContractNo = B.ContractNo"
        If Me.CheckBox3.Checked = True Then
            tmpsql += " where a.Salesman='" + Me.DropDownList1.SelectedValue + "'"
        Else
            tmpsql += " where a.Salesman <> 'XXXXXXX'"
        End If
        tmpsql += " and a.AreaCode='" + Me.DropDownList3.SelectedValue.Substring(0, 1) + "'"
        tmpsql += " and a.CancelDate=''"
        tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "') and B.YN = 'N' "
        Try
            If Me.TextBox1.Text.Trim <> "" Then
                tmpsql += " and B.InvoicePeriod = '" + Mid(Me.TextBox1.Text.Trim, 1, 4) + Mid(Me.TextBox1.Text.Trim, 6, 2) + "'"
            Else
                tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "')"
            End If
        Catch ex As Exception
            Me.TextBox1.Text = GET_W_DATE()
            tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "')"
        End Try

        Dim tt As DataTable
        tt = get_DataTable(tmpsql)
        Session("BatchGrid") = tt
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()

        If tt.Rows.Count > 0 Then
            Dim i As Integer
            '客戶代號
            Me.DropDownList5.Visible = True
            Me.Label3.Visible = True
            Me.DropDownList5.Items.Clear()
            tmpsql = "select distinct a.CustomerNo "
            tmpsql += " from MMSContract a"
            tmpsql += " left outer join MMSInvoiceA B on A.ContractNo = B.ContractNo"
            If Me.CheckBox3.Checked = True Then
                tmpsql += " where a.Salesman='" + Me.DropDownList1.SelectedValue + "'"
            Else
                tmpsql += " where a.Salesman <> 'XXXXXXX'"
            End If
            tmpsql += " and a.AreaCode='" + Me.DropDownList3.SelectedValue.Substring(0, 1) + "'"
            tmpsql += " and a.CancelDate=''"
            tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "') and B.YN = 'N' "
            Try
                If Me.TextBox1.Text.Trim <> "" Then
                    tmpsql += " and B.InvoicePeriod = '" + Mid(Me.TextBox1.Text.Trim, 1, 4) + Mid(Me.TextBox1.Text.Trim, 6, 2) + "'"
                Else
                    tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "')"
                End If
            Catch ex As Exception
                Me.TextBox1.Text = GET_W_DATE()
                tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "')"
            End Try
            tmpsql += " Order by A.CustomerNo" 'add in 20150112
            tt = get_DataTable(tmpsql)
            Me.DropDownList5.Items.Add("")
            For i = 0 To tt.Rows.Count - 1
                Me.DropDownList5.Items.Add(tt.Rows(i).Item(0))
            Next
        Else
            Me.DropDownList5.Visible = False
            Me.Label3.Visible = False
        End If

    End Sub

    '取得範圍西元年月
    Public Function GET_W_DATE() As String
        Dim Dday As Date = DateAdd(DateInterval.Month, 3, Today)
        GET_W_DATE = LTrim(RTrim(Dday.Year.ToString))
        If Dday.Month < 10 Then GET_W_DATE += "0"
        GET_W_DATE += LTrim(RTrim(Dday.Month.ToString))
    End Function

    '取得現在西元年月
    Public Function GET_T_DATE() As String
        Dim Dday As Date = Today
        GET_T_DATE = LTrim(RTrim(Dday.Year.ToString))
        If Dday.Month < 10 Then GET_T_DATE += "0"
        GET_T_DATE += LTrim(RTrim(Dday.Month.ToString))
    End Function

    Protected Sub DropDownList3_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList3.SelectedIndexChanged
        disp_emp()
    End Sub

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
                e.Row.Cells(6).Text = "每期保養<br>金額(含稅)"
        End Select
    End Sub

    Protected Sub Button4_Click(sender As Object, e As System.EventArgs) Handles Button4.Click
        Response.Redirect("InvoiceApply.aspx?Returnflag=1&msg=")
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
            categoryIDList.Add(tt.Rows(i).Item("pk"))
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

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
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

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        RememberOldValues()

        Dim ss As String 

        If Me.txtDtTo.Text.Trim = "" Then
            modUtil.showMsg(Me.Page, "訊息", "請輸入發票日期!!")
            Exit Sub
        End If
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何客戶!!")
            Exit Sub
        End If
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i As Integer
        Dim tt As DataTable
        Dim ssItemName As String

        For i = 0 To categoryIDList.Count - 1

            ss = getDocMomo(Mid(categoryIDList.Item(i), 1, 14), Mid(categoryIDList.Item(i), 15, 6))
            ssItemName = getDocItemName(Mid(categoryIDList.Item(i), 1, 14), Mid(categoryIDList.Item(i), 15, 6))

            'modDB.InsertSignRecord("azitest", "i=" & i.ToString & "categoryIDList.Item(i)=" & categoryIDList.Item(i) & " ssItemName =" & ssItemName, My.User.Name) 'azitest

            tmpsql = "select * from MMSContract where ContractNo="
            tmpsql += "'" + Mid(categoryIDList.Item(i), 1, 14) + "'"
            tt = get_DataTable(tmpsql)
            Dim tdocno As String = getDocNo(Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6))
            tmpsql = "INSERT INTO MMSInvoiceApplyM([DocNo],[CustomerNo],[CustomerName],[ApplyDepartment],[GUINumber],[TitleOfInvoice]"
            tmpsql += " ,[Applicant],[ApplyDate],[ReviewDate],[Reviewers],[Status],[InvoiceType]) VALUES("
            tmpsql += "'" + tdocno + "'"
            tmpsql += ",'" + tt.Rows(0).Item("CustomerNo") + "'"
            tmpsql += ",'" + tt.Rows(0).Item("CustomerName") + "'"
            tmpsql += ",'" + Request.Cookies("STNID").Value + "'"
            tmpsql += ",'" + tt.Rows(0).Item("GUINumber") + "'"
            tmpsql += ",'" + tt.Rows(0).Item("TitleOfInvoice") + "'"
            tmpsql += ",'" + My.User.Name + "'"
            tmpsql += ",'" + Label1.Text + "'"
            tmpsql += ",''"
            tmpsql += ",''"
            tmpsql += ",'N','1'"
            tmpsql += " )"
            EXE_SQL(tmpsql)
            tmpsql = "INSERT INTO MMSInvoiceApplyD([DocNo],[Serial],[DocType],[ContractNo],[InvoicePeriod],[PeriodMaintenanceAmount]"
            tmpsql += ",[ItemName],[UnitPrice],[Quantity],[Amount],[Memo],[InvoiceDate],[InvoiceNo],[Effective])"
            tmpsql += "VALUES("
            tmpsql += "'" + tdocno + "'"
            tmpsql += ",1"
            tmpsql += ",'3'"
            tmpsql += ",'" + tt.Rows(0).Item("ContractNo") + "'"
            tmpsql += ",'" + Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6) + "'"
            tmpsql += "," + tt.Rows(0).Item("PeriodMaintenanceAmount").ToString
            tmpsql += ",'" + ssItemName + "'"
            tmpsql += "," + tt.Rows(0).Item("PeriodMaintenanceAmount").ToString
            tmpsql += ",1"
            tmpsql += "," + tt.Rows(0).Item("PeriodMaintenanceAmount").ToString
            tmpsql += ",'" + ss + "'"
            tmpsql += ",'" + Me.txtDtTo.Text + "'"
            tmpsql += ",''"
            tmpsql += ",'Y'"
            tmpsql += ")"
            EXE_SQL(tmpsql)

            UpdateLastInvoiceDate(Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6), tt.Rows(0).Item("ContractNo"))

            tmpsql = "update MMSInvoiceA set YN='Y',updatetime=getdate() where ContractNo='" + tt.Rows(0).Item("ContractNo") + "' and InvoicePeriod='" + Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6) + "'"
            '20150123 add
            modDB.InsertSignRecord("(InvoiceApply001)", "發票申請 CONTRACTNO =" & tt.Rows(0).Item("ContractNo") & " INVYM =" & Mid(categoryIDList.Item(i), categoryIDList.Item(i).ToString.Length - 5, 6), My.User.Name)
            EXE_SQL(tmpsql)
        Next
        Response.Redirect("InvoiceApply.aspx?Returnflag=1&msg=常態申請成功!!")
    End Sub

    Private Function getDocMomo(ByVal sContractno As String, ByVal sInvoicePeriod As String) As String
        Dim tt As DataTable
        tt = Session("BatchGrid")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If (tt.Rows(i).Item("ContractNo") = sContractno) And (tt.Rows(i).Item("InvoicePeriod") = (sInvoicePeriod.Substring(0, 4) + "/" + sInvoicePeriod.Substring(4, 2))) Then
                getDocMomo = tt.Rows(i).Item("Memo")
                Exit For
            End If
        Next
    End Function

    Private Function getDocItemName(ByVal sContractno As String, ByVal sInvoicePeriod As String) As String
        Dim tt As DataTable
        tt = Session("BatchGrid")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            'modDB.InsertSignRecord("azitest", "getDocItem i=" & i.ToString & " NO=" & tt.Rows(i).Item("ContractNo") & " IP=" & tt.Rows(i).Item("InvoicePeriod"), My.User.Name)
            If (tt.Rows(i).Item("ContractNo") = sContractno) And (tt.Rows(i).Item("InvoicePeriod") = (sInvoicePeriod.Substring(0, 4) + "/" + sInvoicePeriod.Substring(4, 2))) Then
                getDocItemName = tt.Rows(i).Item("ItemName")
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

    'Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
    '    DispGrid()
    'End Sub

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
        dt.Rows(row.DataItemIndex)("memo") = (CType((row.Cells(9).Controls(0)), TextBox)).Text 'modify by azi 
        dt.Rows(row.DataItemIndex)("ItemName") = (CType((row.Cells(7).Controls(0)), TextBox)).Text 'modify by azi
        Me.GridView1.EditIndex = -1
        update_grid1()
    End Sub

    Protected Sub DropDownList4_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList5.SelectedIndexChanged
        Dim tmpsql As String
        tmpsql = "select 0 as checkf,a.CustomerNo,a.ContractNo,a.CustomerName,substring(b.[InvoicePeriod],1,4)+'/'+substring(b.[InvoicePeriod],5,2) as InvoicePeriod,"
        tmpsql += " case a.InvoiceCycle"
        tmpsql += " when '1' then '年'"
        tmpsql += " when '2' then '季'"
        tmpsql += " when '3' then '半年'"
        tmpsql += " when '4' then '二月'"
        tmpsql += " when '5' then '月'"
        tmpsql += " when '6' then '二年'"
        tmpsql += " End AS InvoiceCycle,"
        tmpsql += " a.PeriodMaintenanceAmount "
        tmpsql += " ,case a.DatePlus "
        'tmpsql += " when '1' then ItemName+b.[InvoicePeriod]"
        'tmpsql += " when '2' then b.[InvoicePeriod]+ItemName"
        tmpsql += " when '1' then ItemName + CAST(CONVERT(INT,(SUBSTRING(B.INVOICEPERIOD,1,4)))-1911 AS VARCHAR(3)) + '年' + SUBSTRING(B.INVOICEPERIOD,5,2) + '月' "
        tmpsql += " when '2' then CAST(CONVERT(INT,(SUBSTRING(B.INVOICEPERIOD,1,4)))-1911 AS VARCHAR(3)) + '年' + SUBSTRING(B.INVOICEPERIOD,5,2) + '月' + ItemName "
        tmpsql += " when '3' then ItemName"
        tmpsql += " end as ItemName, '' as Memo, "
        tmpsql += " a.ContractNo + b.[InvoicePeriod] as pk"
        tmpsql += " from MMSContract a"
        tmpsql += " left outer join MMSInvoiceA B on A.ContractNo = B.ContractNo"
        If Me.CheckBox3.Checked = True Then
            tmpsql += " where a.Salesman='" + Me.DropDownList1.SelectedValue + "'"
        Else
            tmpsql += " where a.Salesman <> 'XXXXXXX'"
        End If
        tmpsql += " and a.AreaCode='" + Me.DropDownList3.SelectedValue.Substring(0, 1) + "'"
        tmpsql += " and a.CancelDate=''"
        tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "') and B.YN = 'N' "
        Try
            If Me.TextBox1.Text.Trim <> "" Then
                tmpsql += " and B.InvoicePeriod = '" + Mid(Me.TextBox1.Text.Trim, 1, 4) + Mid(Me.TextBox1.Text.Trim, 6, 2) + "'"
            Else
                tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "')"
            End If
        Catch ex As Exception
            Me.TextBox1.Text = GET_W_DATE()
            tmpsql += " and B.InvoicePeriod <= ('" + GET_W_DATE() + "')"
        End Try
        If Me.DropDownList5.Text.Trim <> "" Then
            tmpsql += " and a.CustomerNo='" + Me.DropDownList5.Text + "'"
        End If
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)
        Session("BatchGrid") = tt
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()
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



    Protected Sub Button5_Click(sender As Object, e As System.EventArgs) Handles Button5.Click
        DispGrid()
    End Sub

    Protected Sub CheckBox3_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If Me.CheckBox3.Checked = True Then
            Me.DropDownList1.Enabled = True
        Else
            Me.DropDownList1.Enabled = False
        End If
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub
End Class
