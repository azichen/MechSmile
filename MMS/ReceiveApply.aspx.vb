Imports System.Data
Imports Microsoft.Reporting.WebForms


Partial Class MMS_Contract_001
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Session("BatchGrid1") = Nothing
            '檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else
                Dim sFileName As String = Request.Path.Substring(Request.Path.LastIndexOf("/") + 1, Request.Path.Length - (Request.Path.LastIndexOf("/") + 1))
                Dim sRol As String = HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
                ViewState("ROL") = modUtil.GetRolData(sFileName, sRol)
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                If ViewState("ROL").Substring(2, 1) = "N" Then Me.Button9.Visible = False
                If ViewState("ROL").Substring(2, 1) = "N" Then Me.Button4.Visible = False
            End If

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            DispGridUp("")
            DispGridDown("")

            Dim tt As DataTable = Session("BatchGrid1")
            tt.Rows.Clear()
            Me.GridView1.DataSource = tt
            Me.GridView1.DataBind()
            Session("BatchGrid1") = tt
            modDB.SetGridViewStyle(Me.GridView1)          '* 套用gridview樣式
            modDB.SetGridViewStyle(Me.GridView3)          '* 套用gridview樣式
            '--------------------------------------------------* 加入連結欄位

            Me.TextBox18.Text = GET_FW_DATE()
            Me.Label4.Visible = False
            Me.TextBox19.Visible = False
            Me.txtDtFrom0.Attributes.Add("onkeypress", "return ToUpper()")
            modUtil.SetDateObj(Me.txtDtFrom, False, Nothing, False)
            modUtil.SetDateObj(Me.txtDtFrom1, False, Nothing, False)
            modUtil.SetDateObj(Me.TextBox14, False, Nothing, False)

            Me.TextBox2.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TextBox3.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TextBox4.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TextBox5.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TextBox6.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TextBox11.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.TextBox13.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            disp_emp()
            Me.DropDownList1.DataSource = Session("dr1")
            Me.DropDownList1.DataBind()
            Me.Label1.Text = "尚未收款發票："
            Button11.Visible = False

            Select Case Request("FormMode")
                Case "add"
                    Me.Label9.Text = "收款資料 - 新增"
                    Me.Button1.Visible = True
                    Button4.Visible = False
                    Button8.Visible = False
                Case "edit"
                    Me.Label9.Text = "收款資料 - 修改"
                    Me.Button4.Visible = True
                    dispReceive(Request("STNID").Trim)
                    Me.Label10.Text = Request("STNID").Trim
                    Button5.Visible = False
                    Button4.Visible = False
                    Button8.Visible = True
                    Button2.Visible = False
                    Button1.Visible = False
                    Button9.Visible = True
                    Button3.Visible = True
                    SetReadOnly()
                Case "read"
                    Me.Label9.Text = "收款資料 - 已審核_唯讀_不可修改"
                    Me.Label1.Text = "對應發票："
                    dispReceive(Request("STNID").Trim)
                    Me.Label10.Text = Request("STNID").Trim
                    Button5.Visible = False
                    Button4.Visible = False
                    Button2.Visible = False
                    Button1.Visible = False
                    Button9.Visible = False
                    Button3.Visible = False
                    Button8.Visible = True
                    'azitemp
                    If My.User.Name = "08804407" Then
                        Button11.Visible = True
                    End If
                    SetReadOnly()
            End Select
            Page.Form.DefaultButton = Button7.UniqueID
            Me.txtDtFrom.Text = Me.GET_FW_DATE
        End If
        If Session("BatchGrid3") Is Nothing Then
            Me.GridView3.DataSource = Session("BatchGrid3")
            Me.GridView3.DataBind()
        End If
        Me.TextBox16.Text = SumTop.ToString
        modUtil.SetObjReadOnly(Me, "TextBox9")
        modUtil.SetObjReadOnly(Me, "TextBox16")
        modUtil.SetObjReadOnly(Me, "TextBox17")
        modUtil.SetObjReadOnly(Me, "TextBox18")
        modUtil.SetObjReadOnly(Me, "TextBox19")
        Me.TextBox17.Text = HttpUtility.UrlDecode(Request.Cookies("NAME").Value)
    End Sub

    Private Sub SetReadOnly()
        Me.TextBox10.ReadOnly = True
        Me.TextBox11.ReadOnly = True
        Me.TextBox12.ReadOnly = True
        Me.TextBox13.ReadOnly = True
        Me.TextBox14.ReadOnly = True
        Me.TextBox2.ReadOnly = True
        Me.TextBox3.ReadOnly = True
        Me.TextBox4.ReadOnly = True
        Me.TextBox5.ReadOnly = True
        Me.TextBox6.ReadOnly = True
        Me.TextBox7.ReadOnly = True
        Me.TextBox8.ReadOnly = True
        Me.TextBox16.ReadOnly = True
        Me.TextBox17.ReadOnly = True
        Me.txtDtFrom1.ReadOnly = True

        Me.txtDtFrom.ReadOnly = True
        Me.TextBox18.ReadOnly = True
        Me.TextBox19.ReadOnly = True
        Me.txtDtFrom0.ReadOnly = True

        DropDownList1.Enabled = False
        Button6.Enabled = False
        Button10.Enabled = False
        'GridView3.Rows(6).

    End Sub

    Private Sub SetUnReadOnly()
        Me.TextBox10.ReadOnly = False
        Me.TextBox11.ReadOnly = False
        Me.TextBox12.ReadOnly = False
        Me.TextBox13.ReadOnly = False
        Me.TextBox14.ReadOnly = False
        Me.TextBox2.ReadOnly = False
        Me.TextBox3.ReadOnly = False
        Me.TextBox4.ReadOnly = False
        Me.TextBox5.ReadOnly = False
        Me.TextBox6.ReadOnly = False
        Me.TextBox7.ReadOnly = False
        Me.TextBox8.ReadOnly = False
        Me.TextBox16.ReadOnly = False
        Me.TextBox17.ReadOnly = False
        Me.txtDtFrom1.ReadOnly = False

        Me.txtDtFrom.ReadOnly = False
        Me.TextBox18.ReadOnly = False
        Me.TextBox19.ReadOnly = False
        Me.txtDtFrom0.ReadOnly = False

        DropDownList1.Enabled = True
        Button6.Enabled = True
        Me.Button4.Visible = True
        Me.Button8.Visible = True
        Me.Button9.Visible = False
        Me.Button3.Visible = False
        Button10.Enabled = True
    End Sub

    Private Sub dispReceive(ByVal docno As String)
        Dim tmpsql As String
        tmpsql = "select * from MMSReceivablesM where ReceivablesNo='" + docno + "'"
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)
        Me.txtDtFrom.Text = tt.Rows(0).Item("DateOfReceipt")
        Me.TextBox18.Text = tt.Rows(0).Item("DocDate")
        'If tt.Rows(0).Item("Status") = "N" Then
        '    Me.TextBox19.Text = "未審核"
        'Else
        '    Me.TextBox19.Text = "審核通過"
        'End If
        Me.txtDtFrom0.Text = tt.Rows(0).Item("CustomerNo")
        Me.TextBox9.Text = tt.Rows(0).Item("CustomerName")
        DispGridUp(docno)
        Me.TextBox2.Text = tt.Rows(0).Item("Cash")
        Me.TextBox3.Text = tt.Rows(0).Item("AmountOfRemittances")
        Me.TextBox4.Text = tt.Rows(0).Item("RemittanceFee")
        Me.TextBox5.Text = tt.Rows(0).Item("SalesDiscounts")
        Me.TextBox6.Text = tt.Rows(0).Item("SalesTax")
        Me.TextBox7.Text = tt.Rows(0).Item("Other")
        Me.txtDtFrom1.Text = tt.Rows(0).Item("RemittanceDate")
        Me.TextBox16.Text = SumTop()
        Me.TextBox17.Text = tt.Rows(0).Item("RemittanceDate")
        Me.DropDownList1.SelectedIndex = -1
        DropDownList1.Items.FindByValue(tt.Rows(0).Item("Cashier")).Selected = True
        Me.TextBox8.Text = tt.Rows(0).Item("Memo")
        Me.Label4.Visible = True
        Me.TextBox19.Visible = True
        If tt.Rows(0).Item("Status") = "N" Then
            Me.TextBox19.Text = "未審核"
        Else
            Me.TextBox19.Text = "審核通過"
            Me.Button4.Visible = False
            Me.Button9.Visible = False 'add in 20140616
            Me.Button3.Visible = False 'add in 20140828
        End If
        DispGridDown1(Me.txtDtFrom0.Text, docno)
    End Sub


    Private Function SumTop() As Integer
        SumTop = 0
        Dim tt As DataTable
        If Session("BatchGrid1") Is Nothing Then Exit Function
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

    Private Function SumCheck() As Integer
        SumCheck = 0
        Dim tt As DataTable
        tt = Session("BatchGrid1")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If IsDBNull(tt.Rows(i).Item("AmountOfCheck")) = False Then
                If tt.Rows(i).Item("AmountOfCheck").ToString <> "" Then
                    SumCheck += tt.Rows(i).Item("AmountOfCheck")
                End If
            End If
        Next
    End Function

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

    Private Sub DispGridDown1(ByVal CustNo As String, ByVal docno As String)
        Dim tmpsql As String
        Dim tr As DataRow
        tmpsql = "SELECT *,0 as revicea"
        tmpsql += " FROM SMILE_HQ.dbo.MMSInvoiceM where CustomerNo='" + CustNo + "'"
        'AZITEMP
        'If Request("FormMode") = "read" Then
        If Me.TextBox19.Text = "審核通過" Then
            tmpsql += " and InvoiceNo in(Select InvoiceNo from dbo.MMSReceive where DocNo='" + docno + "')"
        Else
            tmpsql += " and (UnReceiptAmount<>0 or InvoiceNo in(select InvoiceNo from dbo.MMSReceive where DocNo='" + docno + "'))"
        End If

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
                End If
            Next
        Next
        Session("BatchGrid3") = tt
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
    End Sub

    Private Sub DispGridDown(ByVal CustNo As String)
        Dim tmpsql As String
        Dim tr As DataRow
        tmpsql = "SELECT *,0 as revicea"
        tmpsql += " FROM SMILE_HQ.dbo.MMSInvoiceM where CustomerNo='" + CustNo + "'"
        tmpsql += " and UnReceiptAmount<>0"
        tmpsql += " and Status='Y'"
        tmpsql += " and InvoiceNo<>''"
        tmpsql += " Order By InvoiceDate" 'add by Azi in 20140612
        Dim tt As DataTable = get_DataTable(tmpsql)

        'tr = tt.NewRow
        'tt.Rows.Add(tr)
        Session("BatchGrid3") = tt
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
    End Sub

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

    'Private Sub RememberOldValues2()
    '    Dim categoryIDList As ArrayList = New ArrayList
    '    Dim index As String
    '    For Each row As GridViewRow In GridView3.Rows
    '        index = GridView3.DataKeys(row.RowIndex).Value
    '        Dim result As Boolean = CType(row.FindControl("CheckBox2"), CheckBox).Checked
    '        If Session("CHECKED_ITEMS") Is Nothing = False Then
    '            categoryIDList = Session("CHECKED_ITEMS")
    '        End If
    '        If result Then
    '            If (categoryIDList.Contains(index)) = False Then
    '                categoryIDList.Add(index)
    '            End If
    '        Else
    '            categoryIDList.Remove(index)
    '        End If
    '    Next
    '    Try
    '        If (categoryIDList.Count > 0) Then
    '            Session("CHECKED_ITEMS") = categoryIDList
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub


    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************


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


    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        Dim sSql, tmpstr As String
        sSql = "select"
        sSql += " 0 as checkf,"
        sSql += " 0 as sn,"
        sSql += " *,sn as snM,"
        sSql += " case Status"
        sSql += " when 'Y' then '有效'"
        sSql += " when 'N' then '申請作廢'"
        sSql += " when 'V' then '作廢'"
        sSql += " end as StatusN,"
        sSql += " case InvoiceType"
        sSql += " when '1' then '列印發票'"
        sSql += " when '2' then '手開發票'"
        sSql += " end as InvoiceTypeM"
        sSql += " from MMSInvoiceM"
        sSql += " where InvoiceNo <> 'XXXXXXXXX'"
        If Me.txtDtFrom.Text <> "" Then
            sSql += " and InvoiceDate >= '" + Me.txtDtFrom.Text + "'"
        End If
        If Me.txtDtFrom.Text <> "" Then
            sSql += " and InvoiceNo >= '" + Me.txtDtFrom0.Text + "'"
        End If
        tmpstr = ""

        Dim tt As DataTable = get_DataTable(sSql)

        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            tt.Rows(i).Item("sn") = i
        Next

        Session("BatchGrid1") = tt
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
        Me.GridView3.DataSource = Nothing
        Me.GridView3.DataBind()

    End Sub


#Region "GridView Event"

    Protected Sub GridView1_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        RememberOldValues()
        GridView1.PageIndex = e.NewPageIndex
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
        RePopulateValues()
    End Sub

    Protected Sub GridView1_RowCancelingEdit(sender As Object, e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit
        Me.GridView1.EditIndex = -1
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()

    End Sub

    Protected Sub GridView3_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView3.PageIndexChanging
        'RememberOldValues()
        GridView3.PageIndex = e.NewPageIndex
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
        'RePopulateValues()
    End Sub

    Protected Sub GridView3_RowCancelingEdit(sender As Object, e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView3.RowCancelingEdit
        Me.GridView3.EditIndex = -1
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
    End Sub

    '******************************************************************************************************
    '* GridView_RowDataBound => 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound, GridView3.RowDataBound
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

    Protected Sub GridView1_RowDeleting(sender As Object, e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim i As Integer
        Dim dt As DataTable = CType(Session("BatchGrid1"), DataTable)
        For i = dt.Rows.Count - 1 To 0 Step -1
            If dt.Rows(i).Item(0) = Me.GridView1.Rows(e.RowIndex).Cells(0).Text Then
                dt.Rows(i).Delete()
            End If
        Next
        For i = 1 To dt.Rows.Count
            dt.Rows(i - 1).Item("Serial") = i
        Next
        Session("BatchGrid1") = dt
        Me.GridView1.DataSource = dt
        Me.GridView1.DataBind()
        Me.TextBox16.Text = SumTop.ToString
    End Sub

    Protected Sub GridView1_RowEditing(sender As Object, e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Me.GridView1.EditIndex = e.NewEditIndex
        'Dim row = GridView1.Rows(Me.GridView1.EditIndex)
        'modUtil.SetDateObj((CType((row.Cells(9).Controls(0)), TextBox)), False, Nothing, False)
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
    End Sub


    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Me.TextBox10.Text = GridView1.SelectedRow.Cells(1).Text
        Me.TextBox11.Text = GridView1.SelectedRow.Cells(2).Text
        Me.TextBox12.Text = GridView1.SelectedRow.Cells(3).Text
        Me.TextBox13.Text = GridView1.SelectedRow.Cells(4).Text
        Me.TextBox14.Text = GridView1.SelectedRow.Cells(5).Text
        Me.TextBox15.Text = GridView1.SelectedRow.Cells(0).Text

    End Sub
#End Region

    Private Function getDocNo(ByVal sn As String) As String
        Dim tt As DataTable
        tt = Session("BatchGrid1")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("sn") = sn Then
                If tt.Rows(i).Item("Status") = "N" Then
                    getDocNo = tt.Rows(i).Item("InvoiceNo")
                Else
                    getDocNo = ""
                End If

                Exit For
            End If
        Next
    End Function

    '取得西元年月日
    Public Function GET_FW_DATE() As String
        GET_FW_DATE = LTrim(RTrim(Today.Year.ToString)) + "/"
        If Today.Month < 10 Then GET_FW_DATE += "0"
        GET_FW_DATE += LTrim(RTrim(Today.Month.ToString)) + "/"
        If Today.Day < 10 Then GET_FW_DATE += "0"
        GET_FW_DATE += LTrim(RTrim(Today.Day.ToString))
    End Function

    Private Function checkDoc(ByVal sn As String) As Boolean
        Dim tt As DataTable
        Dim dd As String
        tt = Session("BatchGrid1")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("sn") = sn Then
                dd = tt.Rows(i).Item("DocNo")
                Exit For
            End If
        Next

        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("DocNo") = dd And categoryIDList.Contains(tt.Rows(i).Item("sn").ToString) = False Then
                checkDoc = False
                Exit Function
            End If
        Next
        checkDoc = True
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

    Private Function changeDate(ByVal dd As String) As String
        If dd = "" Then
            changeDate = ""
            Exit Function
        End If
        Dim tdate As Date
        tdate = dd
        changeDate = LTrim(RTrim(tdate.Year.ToString)) + "/"
        If tdate.Month < 10 Then changeDate += "0"
        changeDate += LTrim(RTrim(tdate.Month.ToString)) + "/"
        If tdate.Day < 10 Then changeDate += "0"
        changeDate += LTrim(RTrim(tdate.Day.ToString))
    End Function


    Protected Sub txtDtFrom0_TextChanged(sender As Object, e As System.EventArgs) Handles txtDtFrom0.TextChanged
        'azitemp
        Dim Areacode As String = ""
        Areacode = modUnset.GetAreaCode(My.User.Name).Trim
        '
        Dim tmpsql As String
        tmpsql = "select CustomerName,Cashier from MMSCustomers where Effective='Y' "
        tmpsql += " and CustomerNo='" + txtDtFrom0.Text + "'"
        If (Areacode <> "") And (Areacode <> "F") Then
            tmpsql += " AND AREACODE ='" + Areacode + "'"
        End If
        '
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)
        If tt.Rows.Count = 0 Then
            Me.TextBox9.Text = ""
            modUtil.showMsg(Me.Page, "訊息", "查無 : 客戶代號'" + txtDtFrom0.Text + "' ！")
            Me.DropDownList1.SelectedIndex = 0
            Exit Sub
        Else
            Me.TextBox9.Text = tt.Rows(0).Item(0)
            DispGridDown(Me.txtDtFrom0.Text.Trim)
            '
            DropDownList1.SelectedIndex = -1
            DropDownList1.Items.FindByValue(tt.Rows(0).Item(1)).Selected = True
            If Me.GridView3.Rows.Count = 0 Then
                modUtil.showMsg(Me.Page, "訊息", "查無 : 此客戶無尚未沖銷之發票 ！")
            End If
        End If
        '       
    End Sub

    Private Sub clean_check()
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox13.Text = ""
        Me.TextBox14.Text = ""
        Me.TextBox15.Text = ""
    End Sub

    Protected Sub GridView3_RowEditing(sender As Object, e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView3.RowEditing
        Me.GridView3.EditIndex = e.NewEditIndex
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
    End Sub

    Protected Sub GridView3_RowUpdating(sender As Object, e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView3.RowUpdating
        Dim row = GridView3.Rows(Me.GridView3.EditIndex)
        Dim i As Integer
        Try
            i = (CType((row.Cells(5).Controls(0)), TextBox)).Text
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "訊息", "支票金額不可為非數字!! \n")
            Exit Sub
        End Try
        Dim dt As DataTable = CType(Session("BatchGrid3"), DataTable)
        dt.Rows(row.DataItemIndex)("revicea") = (CType((row.Cells(5).Controls(0)), TextBox)).Text
        Me.GridView3.EditIndex = -1
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
    End Sub


    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        '檢查必要欄位
        Dim tmpstr As String = ""

        If Me.TextBox18.Text = "" Then
            tmpstr += "請選擇建檔日期!! \n"
        End If

        If Me.txtDtFrom.Text = "" Then
            tmpstr += "請選擇收款日期!! \n"
        End If
        If Me.txtDtFrom0.Text = "" Then
            tmpstr += "請輸入客戶代號!! \n"
        Else
            If Me.TextBox9.Text = "" Then
                tmpstr += "請輸入有效的客戶代號!! \n"
            End If
        End If
        If Me.SumTop = 0 Then
            tmpstr += "請輸入收款資料!! \n"
        End If
        If Me.SumDown = 0 Then
            tmpstr += "請輸入沖銷之發票資料!! \n"
        End If
        If Me.TextBox3.Text <> "0" And Me.txtDtFrom1.Text = "" Then
            tmpstr += "請選擇匯款日期!! \n"
        End If
        If Me.SumDown <> SumTop() Then
            tmpstr += "收款金額與沖銷金額不符!! \n"
        End If
        If tmpstr <> "" Then
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
            Exit Sub
        End If
        'modUtil.showMsg(Me.Page, "訊息", Me.SumTop.ToString + " \n" + SumDown.ToString)

        Dim tmpsql As String = ""
        Dim docno As String = getDocNo()
        tmpsql = "insert into MMSReceivablesM values("
        tmpsql += " '" + docno + "'," 'ReceivablesNo
        tmpsql += " '" + Me.TextBox18.Text + "'," 'DocDate
        tmpsql += " '" + Me.txtDtFrom0.Text + "'," 'CustomerNo
        tmpsql += " '" + Me.TextBox9.Text + "'," 'CustomerName
        tmpsql += " '" + Me.txtDtFrom.Text + "'," 'DateOfReceipt
        tmpsql += SumCheck().ToString + "," 'AmountOfCheck(int)
        tmpsql += TextBox2.Text + "," 'Cash(int)
        tmpsql += TextBox3.Text + "," 'AmountOfRemittances(int)
        tmpsql += TextBox4.Text + "," 'RemittanceFee(int)
        tmpsql += " '" + txtDtFrom1.Text + "'," 'RemittanceDate
        tmpsql += TextBox5.Text + "," 'SalesDiscounts(int)
        tmpsql += TextBox6.Text + "," 'SalesTax(int)
        tmpsql += TextBox7.Text + "," 'Other(int)
        tmpsql += SumTop().ToString + "," 'TotalAmount(int)
        tmpsql += " '" + Request.Cookies("STNID").Value + "'," 'PreparedBy
        tmpsql += " '" + Me.DropDownList1.SelectedValue + "'," 'Cashier
        tmpsql += " '" + TextBox8.Text + "'," 'Memo
        tmpsql += " ''," 'ReviewDate
        tmpsql += " ''," 'Reviewers
        tmpsql += " 'N',''" 'Status,SYS
        tmpsql += ")"
        'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql, My.User.Name) 'AZITEST      

        EXE_SQL(tmpsql)

        UpdateInvoice() 'update 已收金額和未收金額

        InsertMMSCheque(docno) '紀錄支票資料

        InsertMMSReceive(docno) '紀錄收款沖帳

        modDB.InsertSignRecord("(ReceiveApply)機械保養收款", "新增 DOCNO =" & docno, My.User.Name)

        Response.Redirect("Receive.aspx?Returnflag=1&msg=收款申請成功!!")
    End Sub

    Private Function getDocNo() As String
        Dim tmpsql As String
        Dim tt As DataTable
        tmpsql = "select max(ReceivablesNo) from MMSReceivablesM where ReceivablesNo like '" + GET_W_DATE() + "%'"
        tt = get_DataTable(tmpsql)
        If IsDBNull(tt.Rows(0).Item(0)) Then
            getDocNo = GET_W_DATE() + "001"
        Else
            getDocNo = Int(tt.Rows(0).Item(0)) + 1
        End If
    End Function

    '取得西元年月
    Public Function GET_W_DATE() As String
        GET_W_DATE = LTrim(RTrim(Today.Year.ToString))
        If Today.Month < 10 Then GET_W_DATE += "0"
        GET_W_DATE += LTrim(RTrim(Today.Month.ToString))
    End Function


    Private Sub UnUpdateInvoice(ByVal docno As String)
        Dim tmpsql As String
        Dim tt As DataTable
        tmpsql = "select * from MMSReceive where DocNo='" + Me.Label10.Text + "'"
        tt = get_DataTable(tmpsql)
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            'If tt.Rows(i).Item("revicea").ToString <> "" Then
            tmpsql = "update MMSInvoiceM set ReceiptAmount=ReceiptAmount-" + tt.Rows(i).Item("AmountOfReceive").ToString + ",UnReceiptAmount=UnReceiptAmount+" + tt.Rows(i).Item("AmountOfReceive").ToString + " where InvoiceNo='" + tt.Rows(i).Item("InvoiceNo").ToString + "'"
            EXE_SQL(tmpsql)
            'End If
        Next
    End Sub

    Private Sub UpdateInvoice()
        Dim tmpsql As String
        Dim tt As DataTable
        tt = Session("BatchGrid3")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("revicea").ToString <> "" Then
                tmpsql = "update MMSInvoiceM set ReceiptAmount=ReceiptAmount+" + tt.Rows(i).Item("revicea").ToString + ",UnReceiptAmount=UnReceiptAmount-" + tt.Rows(i).Item("revicea").ToString + " where InvoiceNo='" + tt.Rows(i).Item("InvoiceNo").ToString + "'"
                EXE_SQL(tmpsql)
            End If
        Next
    End Sub

    Private Sub InsertMMSCheque(ByVal docno As String)
        Dim tmpsql As String
        Dim tt As DataTable
        tt = Session("BatchGrid1")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("ChequeNo").ToString <> "" Then
                tmpsql = "insert into MMSCheque values("
                tmpsql += "'" + docno + "'," 'DocNo
                tmpsql += tt.Rows(i).Item("Serial").ToString + "," 'Serial
                tmpsql += "'" + tt.Rows(i).Item("ChequeNo").ToString + "'," 'ChequeNo
                tmpsql += tt.Rows(i).Item("AmountOfCheck").ToString + "," 'AmountOfCheck
                tmpsql += "'" + tt.Rows(i).Item("PayingBank").ToString + "'," 'PayingBank
                tmpsql += "'" + tt.Rows(i).Item("BankAccount").ToString + "'," 'BankAccount
                tmpsql += "'" + tt.Rows(i).Item("ChequeExpiryDate").ToString + "'" 'ChequeExpiryDate
                tmpsql += ")"
                'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql, My.User.Name) 'AZITEST      
                EXE_SQL(tmpsql)
            End If
        Next
    End Sub

    Private Sub InsertMMSReceive(ByVal docno As String)
        Dim tmpsql As String
        Dim tt As DataTable
        tt = Session("BatchGrid3")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("revicea").ToString <> "" Then
                If tt.Rows(i).Item("revicea") > 0 Then
                    tmpsql = "insert into MMSReceive values("
                    tmpsql += "'" + docno + "'," 'DocNo
                    tmpsql += "'" + tt.Rows(i).Item("InvoiceNo").ToString + "'," 'InvoiceNo
                    tmpsql += "'" + tt.Rows(i).Item("InvoiceDate").ToString + "'," 'InvoiceDate
                    tmpsql += tt.Rows(i).Item("revicea").ToString  'AmountOfReceive
                    tmpsql += ",'')" 'SYS
                    'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql, My.User.Name) 'AZITEST
                    EXE_SQL(tmpsql)
                End If
            End If
        Next
    End Sub

    Protected Sub Button4_Click(sender As Object, e As System.EventArgs) Handles Button4.Click
        '檢查必要欄位
        Dim tmpstr As String = ""
        If Me.txtDtFrom.Text = "" Then
            tmpstr += "請選擇收款日期!! \n"
        End If
        If Me.txtDtFrom0.Text = "" Then
            tmpstr += "請輸入客戶代號!! \n"
        Else
            If Me.TextBox9.Text = "" Then
                tmpstr += "請輸入有效的客戶代號!! \n"
            End If
        End If
        If Me.SumTop = 0 Then
            tmpstr += "請輸入收款資料!! \n"
        End If
        If Me.SumDown = 0 Then
            tmpstr += "請輸入沖銷之發票資料!! \n"
        End If
        If Me.TextBox3.Text <> "0" And Me.txtDtFrom1.Text = "" Then
            tmpstr += "請選擇匯款日期!! \n"
        End If
        If Me.SumDown <> Me.SumTop Then
            tmpstr += "收款金額與沖銷金額不符!! \n"
        End If
        If tmpstr <> "" Then
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
            Exit Sub
        End If
        'modUtil.showMsg(Me.Page, "訊息", Me.SumTop.ToString + " \n" + SumDown.ToString)
        '恢復資料
        Dim docno As String = Me.Label10.Text
        Dim tmpsql As String = ""
        Me.UnUpdateInvoice(docno)
        tmpsql = "delete from MMSCheque where DocNo='" + docno + "'"
        EXE_SQL(tmpsql)
        tmpsql = "delete from MMSReceive where DocNo='" + docno + "'"
        EXE_SQL(tmpsql)

        modDB.InsertSignRecord("(ReceiveApply)機械保養收款", "BTN4:預DEL DOCNO =" & docno, My.User.Name)

        'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql, My.User.Name) 'AZITEST

        '-------------------------------------------------------------------
        tmpsql = "update MMSReceivablesM set "
        tmpsql += " DocDate='" + Me.TextBox18.Text + "'," 'DocDate
        tmpsql += " CustomerNo='" + Me.txtDtFrom0.Text + "'," 'CustomerNo
        tmpsql += " CustomerName='" + Me.TextBox9.Text + "'," 'CustomerName
        tmpsql += " DateOfReceipt='" + Me.txtDtFrom.Text + "'," 'DateOfReceipt
        tmpsql += " AmountOfCheck=" + SumCheck().ToString + "," 'AmountOfCheck(int)
        tmpsql += " Cash=" + TextBox2.Text + "," 'Cash(int)
        tmpsql += "AmountOfRemittances=" + TextBox3.Text + "," 'AmountOfRemittances(int)
        tmpsql += "RemittanceFee=" + TextBox4.Text + "," 'RemittanceFee(int)
        tmpsql += "RemittanceDate='" + txtDtFrom1.Text + "'," 'RemittanceDate
        tmpsql += "SalesDiscounts=" + TextBox5.Text + "," 'SalesDiscounts(int)
        tmpsql += "SalesTax=" + TextBox6.Text + "," 'SalesTax(int)
        tmpsql += "Other=" + TextBox7.Text + "," 'Other(int)
        tmpsql += "TotalAmount=" + SumTop().ToString + "," 'TotalAmount(int)
        tmpsql += " PreparedBy='" + Request.Cookies("STNID").Value + "'," 'PreparedBy
        tmpsql += " Cashier='" + Me.DropDownList1.SelectedValue + "'," 'Cashier
        tmpsql += " Memo='" + TextBox8.Text + "'" 'Memo
        tmpsql += " where ReceivablesNo='" + docno + "'"

        'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql.Substring(0, 90), My.User.Name) 'AZITEST
        'modDB.InsertSignRecord("AziTest", "ReceivablesNo='" + docno, My.User.Name) 'AZITEST

        EXE_SQL(tmpsql)

        UpdateInvoice() 'update 已收金額和未收金額

        InsertMMSCheque(docno) '紀錄支票資料

        InsertMMSReceive(docno) '紀錄收款沖帳
'
        modDB.InsertSignRecord("(ReceiveApply)機械保養收款", "BTN4:修改完成 DOCNO =" & docno, My.User.Name)
        Response.Redirect("Receive.aspx?Returnflag=1&msg=收款資料修改成功!!")
        'Response.Redirect("Receive.aspx?Returnflag=1&msg=" & tmpsql)
    End Sub

    Protected Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Response.Redirect("Receive.aspx?Returnflag=1&msg=")
    End Sub

    Protected Sub Button6_Click(sender As Object, e As System.EventArgs) Handles Button6.Click
        'Retrieve the table from the session object.
        Dim dt As DataTable = CType(Session("BatchGrid1"), DataTable)
        Dim tmpstr As String = ""

        If Me.TextBox10.Text = "" Then
            tmpstr += "支票號碼不可空白!! \n"
        End If

        If Me.TextBox11.Text = "" Then
            tmpstr += "支票金額不可空白!! \n"
        Else
            If Int(Me.TextBox11.Text) = 0 Then
                tmpstr += "支票金額不可為0!! \n"
            End If
        End If

        If Me.TextBox12.Text = "" Then
            tmpstr += "付款銀行不可空白!! \n"
        End If
        If Me.TextBox13.Text = "" Then
            tmpstr += "銀行帳號不可空白!! \n"
        End If
        If Me.TextBox14.Text = "" Then
            tmpstr += "支票到期日不可空白!! \n"
        End If

        If tmpstr <> "" Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & tmpstr)
            Exit Sub
        End If

        If Me.TextBox15.Text <> "" Then
            Dim i As Integer
            For i = 0 To dt.Rows.Count - 1
                If dt.Rows(i).Item(0).ToString.Trim = Me.TextBox15.Text Then
                    dt.Rows(i).Item(1) = Me.TextBox10.Text
                    dt.Rows(i).Item(2) = Me.TextBox11.Text
                    dt.Rows(i).Item(3) = Me.TextBox12.Text
                    dt.Rows(i).Item(4) = Me.TextBox13.Text
                    dt.Rows(i).Item(5) = Me.TextBox14.Text
                    Session("BatchGrid1") = dt
                    Me.GridView1.DataSource = Session("BatchGrid1")
                    Me.GridView1.DataBind()
                End If
            Next
        Else
            'Bind data to the GridView control.
            Session("BatchGrid1") = dt
            Me.GridView1.DataSource = Session("BatchGrid1")
            Me.GridView1.DataBind()

            Me.TextBox16.Text = SumTop.ToString

            Dim tt As DataTable
            Dim tr As DataRow
            tt = Session("BatchGrid1")
            tr = tt.NewRow
            tr.Item(1) = Me.TextBox10.Text
            tr.Item(2) = Me.TextBox11.Text
            tr.Item(3) = Me.TextBox12.Text
            tr.Item(4) = Me.TextBox13.Text
            tr.Item(5) = Me.TextBox14.Text

            tt.Rows.Add(tr)
            Dim i As Integer
            For i = 1 To tt.Rows.Count
                tt.Rows(i - 1).Item("Serial") = i
            Next
            Session("BatchGrid1") = tt
            Me.GridView1.DataSource = Session("BatchGrid1")
            Me.GridView1.DataBind()
        End If
        clean_check()
        Me.TextBox16.Text = SumTop.ToString
    End Sub

    Protected Sub Button8_Click(sender As Object, e As System.EventArgs) Handles Button8.Click
        Response.Redirect("Receive.aspx?Returnflag=1&msg=")
    End Sub

    Protected Sub Button9_Click(sender As Object, e As System.EventArgs) Handles Button9.Click
        SetUnReadOnly()
    End Sub

    Protected Sub Button5_Click(sender As Object, e As System.EventArgs) Handles Button5.Click
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox13.Text = ""
        Me.TextBox14.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox16.Text = ""
        Me.TextBox17.Text = ""
        Me.txtDtFrom1.Text = ""

        Me.txtDtFrom.Text = ""
        Me.TextBox18.Text = ""
        Me.TextBox19.Text = ""
        Me.txtDtFrom0.Text = ""

        Dim dt As DataTable = CType(Session("BatchGrid1"), DataTable)
        dt.Rows.Clear()
        Session("BatchGrid1") = dt
        Me.GridView1.DataSource = dt
        Me.GridView1.DataBind()

        dt = Session("BatchGrid3")
        dt.Rows.Clear()
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
    End Sub

    Protected Sub Button10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim i, mon As Integer
        Dim tt As DataTable
        tt = Session("BatchGrid3")
        mon = SumTop()
        'UnReceiptAmount,revicea
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("UnReceiptAmount") > mon Then
                tt.Rows(i).Item("revicea") = mon
                mon = 0
                Exit For
            Else
                tt.Rows(i).Item("revicea") = tt.Rows(i).Item("UnReceiptAmount")
                mon = mon - tt.Rows(i).Item("UnReceiptAmount")
            End If
        Next
        Session("BatchGrid3") = tt
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click '刪除
        'Add in 20140828
        Dim docno As String = Me.Label10.Text
        Dim tmpsql As String = ""
        modDB.InsertSignRecord("(ReceiveApply)機械保養收款", "刪除 DOCNO =" & docno, My.User.Name)
        Me.UnUpdateInvoice(docno)
        '
        tmpsql = "delete from MMSCheque where DocNo='" + docno + "'"
        EXE_SQL(tmpsql)
        tmpsql = "delete from MMSReceive where DocNo='" + docno + "'"
        EXE_SQL(tmpsql)
        tmpsql = "delete from MMSReceivablesM where RECEIVABLESNO='" + docno + "'"
        EXE_SQL(tmpsql)
        modDB.InsertSignRecord("(ReceiveApply)機械保養收款", "刪除完成 DOCNO =" & docno, My.User.Name)
        Response.Redirect("Receive.aspx?Returnflag=1&msg=收款資料刪除成功!!")
    End Sub

    Protected Sub Button11_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button11.Click
        '20141119_核帳單列印
        Dim sValidOK As Boolean = True
        Dim sMsg As String = ""
        
        Dim Areacode As String = ""
        Areacode = modUnset.GetAreaCode(My.User.Name).Trim
        rptCmp.Visible = True

        Dim sParas(0) As ReportParameter '* 依參數個數產生陣列大小
        sParas(0) = New ReportParameter("DeptName", "測試中") '* 指定參數值
        'sParas(1) = New ReportParameter("BB1", Me.txtAreaItem.Text.Substring(0, 1))
        'sParas(2) = New ReportParameter("BB2", Me.txtAreaItem.Text.Substring(1, Me.txtAreaItem.Text.Length - 1))
        'sParas(3) = New ReportParameter("CC1", Me.TextBox1.Text)
        'sParas(4) = New ReportParameter("CC2", Me.TextBox2.Text)

        '-------------------------------------* 報表路徑及認證
        'lblTitle.Text = "OK! 截止日=" & sParas(5).Values.ToString & " 回饋日=" & sParas(6).Values.ToString 'azitest
        'showMsg(Me.Page, "OK! 截止日=" & sParas(5).Values.ToString & " 回饋日=" & sParas(6).Values.ToString, sMsg)
        rptCmp.ProcessingMode = ProcessingMode.Remote
        rptCmp.ServerReport.ReportServerCredentials = Me.ReportCredentials '* 身分認證 2008/04/03 NEC 杜
        rptCmp.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))

        rptCmp.ServerReport.ReportPath = "/MMS/MMS_R05"
        rptCmp.ServerReport.SetParameters(sParas)
        rptCmp.ShowParameterPrompts = False
        rptCmp.DataBind()
    End Sub

End Class
