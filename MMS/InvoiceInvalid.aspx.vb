Imports System.Data
Partial Class MMS_InvoiceInvalid
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
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                If ViewState("ROL").Substring(1, 1) = "N" Then Me.btnQry.Visible = False
            End If

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            'modDB.SetGridViewStyle(Me.GridView1)          '* 套用gridview樣式
            'DispGridE()
            '--------------------------------------------------* 加入連結欄位

            Me.lblMsg.Text = Request("msg")

            Me.Label2.Text = Format(Now, "yyyy/MM/dd")
            dispArea()
            Me.dispEmployee()
            Me.DropDownList1.Focus()
        End If
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

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        dispEmployee()
    End Sub

    '顯示區域所屬業務員收款員
    Private Sub dispEmployee()
        Dim tmpsql As String
        'If Me.DropDownList1.SelectedValue = "" And Me.DropDownList1.SelectedIndex = -1 Then
        '    Try
        '        tmpsql = "SELECT AreaCode, AreaCode+'_'+AreaName  as AreaName FROM MMSArea WHERE (Effective = 'Y')"
        '        Me.SqlDataSource2.SelectCommand = tmpsql
        '        Me.SqlDataSource2.DataBind()
        '        Me.DropDownList5.SelectedIndex = 0
        '    Catch ex As Exception
        '        Dim sss As String = ""
        '    End Try
        'End If

        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee"
        tmpsql += " WHERE Effective = 'Y' and AreaCode='" + Me.DropDownList1.SelectedValue + "'"
        tmpsql += " and Salesman = 'Y'"
        tt2 = get_DataTable(tmpsql)
        Me.DropDownList2.DataTextField = "EmployeeName"
        Me.DropDownList2.DataValueField = "EmployeeNo"
        DropDownList2.DataSource = tt2
        DropDownList2.DataBind()

    End Sub

    Protected Sub btnQry0_Click(sender As Object, e As System.EventArgs) Handles btnQry0.Click
        Me.DropDownList1.SelectedIndex = -1
        Me.DropDownList2.DataSource = Nothing
        Me.DropDownList2.Items.Clear()
        Me.TextBox1.Text = ""
        Me.Label5.Text = ""
        Me.Label4.Text = ""
        Me.Label3.Text = ""
        Me.Label6.Text = ""
        Me.Label7.Text = ""
        Me.TextBox2.Text = ""
        Me.GridView3.DataSource = Nothing
        Me.GridView3.DataBind()
        dispEmployee()
    End Sub

    Protected Sub btnQry_Click(sender As Object, e As System.EventArgs) Handles btnQry.Click
        If checkEdit() <> "" Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & checkEdit())
            Exit Sub
        End If
        Dim tmpsql As String
        
        tmpsql = "update MMSInvoiceM set Status='N' where InvoiceNo='" + Me.TextBox1.Text + "'"
        EXE_SQL(tmpsql)
        tmpsql = "Insert INTO MMSInvoiceInvalid ([InvoiceDate],[InvoiceNo],[GUINumber],[TitleOfInvoice],[ApplyDepartment],[Applicant],[ApplyDate],[VoidReason])"
        tmpsql += " VALUES ('" + Me.Label6.Text + "','" + Me.TextBox1.Text + "','" + Me.Label7.Text + "','" + Me.Label3.Text + "','" + Me.DropDownList1.SelectedValue + "','" + Me.DropDownList2.SelectedValue + "','" + Me.Label2.Text + "','" + Me.TextBox2.Text + "')"
        'LOGSQL(tmpsql) 'azitest
        EXE_SQL(tmpsql)
        Response.Redirect("InvoiceInvalid.aspx?Returnflag=1&msg=資料新增成功!!")
    End Sub

    'Private Function VALIDCHK(ByVal VINVOICENUM As String) As String
    '    Dim tt_CHK As DataTable
    '    Dim TMPSQL As String
    '    Dim MSGSTR As String
    '    TMPSQL = "Select INVOICENO,STATUS from MMSInvoiceM where InvoiceNo='" + Me.TextBox1.Text + "'"
    '    tt_CHK = get_DataTable(TMPSQL)
    '    MSGSTR = ""
    '    If tt_CHK.Rows.Count = 0 Then
    '        MSGSTR = "查無此發票號碼!!"
    '    Else
    '        If tt_CHK.Rows(0).Item("STATUS").ToString = "N" Then
    '            MSGSTR = "已發票申請中!"
    '        ElseIf tt_CHK.Rows(0).Item("STATUS").ToString = "V" Then
    '            MSGSTR = "發票已作廢!"
    '        ElseIf tt_CHK.Rows(0).Item("ReceiptAmount").ToString > "0" Then
    '            MSGSTR = "發票已作收款，不可申請作廢!"
    '        End If
    '    End If
    '    '
    '    If MSGSTR = "" Then
    '        VALIDCHK = "Y"
    '    Else
    '        VALIDCHK = "N " + MSGSTR
    '    End If
    'End Function
    
    Private Function checkEdit() As String
        checkEdit = ""
        If Me.DropDownList1.SelectedIndex = -1 Then
            checkEdit += "請選擇部門! \n"
        End If
        If Me.DropDownList2.SelectedIndex = -1 Then
            checkEdit += "請選擇申請人! \n"
        End If
        If Me.TextBox1.Text = "" Then
            checkEdit += "請輸入發票號碼! \n"
        End If
        If Me.TextBox2.Text = "" Then
            checkEdit += "請輸入發票作廢原因! \n"
        End If
    End Function

    Protected Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Me.Label5.Text = ""
        Me.Label4.Text = ""
        Me.Label3.Text = ""
        Me.Label6.Text = ""
        Me.Label7.Text = ""
        Dim tmpsql As String
        Dim MSGSTR As String = ""
        tmpsql = "select *,REPLACE(CONVERT(varchar(20), (CAST(Amount AS money)), 1), '.00', '') as Amount1 from MMSInvoiceM where InvoiceNo='" + Me.TextBox1.Text + "' " 'and Status='Y'
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)

        'modDB.InsertSignRecord("AziTest", "tmpsql =" & tmpsql.Substring(0, 90), My.User.Name) 'AZITEST
        'modDB.InsertSignRecord("AziTest", "tmpsql =" & tmpsql.Substring(90, 90), My.User.Name) 'AZITEST

        
        'If tt.Rows.Count = 0 Then
        '    modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "查無此發票號碼!!")
        '    Me.GridView3.DataSource = Nothing
        '    Me.GridView3.DataBind()
        '    Exit Sub
        'End If

        'TMPSQL = "Select INVOICENO,STATUS from MMSInvoiceM where InvoiceNo='" + Me.TextBox1.Text + "'"
        'tt_CHK = get_DataTable(TMPSQL)
        MSGSTR = ""
        If tt.Rows.Count = 0 Then
            MSGSTR = "查無此發票號碼!!"
        Else
            If tt.Rows(0).Item("STATUS").ToString = "N" Then
                MSGSTR = "已發票申請中!"
            ElseIf tt.Rows(0).Item("STATUS").ToString = "V" Then
                MSGSTR = "發票已作廢!"
            ElseIf tt.Rows(0).Item("ReceiptAmount").ToString > "0" Then
                MSGSTR = "發票已作收款，不可申請作廢!"
            End If
        End If
        '
        If MSGSTR = "" Then
            MSGSTR = "Y"
        Else
            MSGSTR = "N " + MSGSTR
        End If


        If MSGSTR <> "Y" Then
            MSGSTR = MSGSTR.Substring(2, MSGSTR.Length - 2).Trim
            btnQry.Visible = False
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & MSGSTR)
            Me.GridView3.DataSource = Nothing
            Me.GridView3.DataBind()
            Exit Sub
        End If

        'If tt.Rows(0).Item("ReceiptAmount").ToString <> "0" Then
        '    modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "此發票已有收款資料無法作廢!!")
        '    Me.GridView3.DataSource = Nothing
        '    Me.GridView3.DataBind()
        '    Exit Sub
        'End If

        '20160118 增加發票狀態檢查
        'If tt.Rows(0).Item("STATUS").ToString <> "Y" Then
        '    If tt.Rows(0).Item("STATUS").ToString = "N" Then
        '        MSGSTR = "此發票已有申請作廢!"
        '    ElseIf tt.Rows(0).Item("STATUS").ToString = "V" Then
        '        MSGSTR = "此發票已作廢!"
        '    End If
        '    modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & MSGSTR)
        '    Me.GridView3.DataSource = Nothing
        '    Me.GridView3.DataBind()
        '    Exit Sub
        'End If

        'modDB.InsertSignRecord("AziTest", "InvoiceDate=" & tt.Rows(0).Item("InvoiceDate").ToString, My.User.Name) 'AZITEST

        Me.Label5.Text = tt.Rows(0).Item("InvoiceDate").ToString
        Me.Label4.Text = tt.Rows(0).Item("Amount1").ToString
        Me.Label3.Text = tt.Rows(0).Item("TitleOfInvoice").ToString
        Me.Label6.Text = tt.Rows(0).Item("InvoicePeriod").ToString
        Me.Label7.Text = tt.Rows(0).Item("GUINumber").ToString

        tmpsql = "select *,REPLACE(CONVERT(varchar(20), (CAST(Amount AS money)), 1), '.00', '') as Amount1"
        tmpsql += ",REPLACE(CONVERT(varchar(20), (CAST(UnitPrice AS money)), 1), '.00', '') as UnitPrice1"
        tmpsql += " from MMSInvoiceD where InvoiceNo='" + Me.TextBox1.Text + "'"
        Session("BatchGrid3") = get_DataTable(tmpsql)
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()

        btnQry.Visible = True
    End Sub

    Protected Sub LOGSQL(ByVal datastr As String)
        Dim TMPSTR As String = datastr
        While (TMPSTR.Length) > 0
            If TMPSTR.Length > 95 Then
                modDB.InsertSignRecord("AziTest", TMPSTR.Substring(0, 95), My.User.Name) 'AZITEST
                TMPSTR = TMPSTR.Substring(95, TMPSTR.Length - 95)
            Else
                modDB.InsertSignRecord("AziTest", TMPSTR.Substring(0, TMPSTR.Length - 1), My.User.Name) 'AZITEST
                TMPSTR = ""
            End If
        End While
    End Sub

End Class