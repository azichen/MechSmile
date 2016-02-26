Imports System.Data
Partial Class MMS_BatchS
    Inherits System.Web.UI.Page

    Dim tt As DataTable

    Protected Sub RadioButtonList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles RadioButtonList1.SelectedIndexChanged
        If Me.RadioButtonList1.SelectedValue = "1" Then
            Me.Lable1.Text = "原業務員"
            Me.Label2.Text = "新業務員"
        Else
            Me.Lable1.Text = "原收款員"
            Me.Label2.Text = "新收款員"
        End If
        disp_emp()
        Session("CHECKED_ITEMS") = Nothing
        'DispGrid()
    End Sub

    Private Sub DispGrid()
        Dim tmpsql As String
        tmpsql = "select 0 as checkf,CustomerNo,CustomerName from MMSCustomers"
        If Me.RadioButtonList1.SelectedValue = "1" Then
            tmpsql += " where Salesman='" + Me.DropDownList1.SelectedValue + "'"
        Else
            tmpsql += " where Cashier='" + Me.DropDownList1.SelectedValue + "'"
        End If
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)
        Session("BatchGrid") = tt
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()
    End Sub

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
                If ViewState("ROL").Substring(1, 1) = "N" Then Me.Button1.Visible = False
            End If
            dispArea()
            Try
                Me.DropDownList3.SelectedIndex = 0
            Catch ex As Exception

            End Try
            disp_emp()
            'DispGrid()
            DropDownList3.Focus()
        End If
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
        Select Case Me.RadioButtonList1.SelectedValue
            Case "1"
                tmpsql += " where Salesman='Y'"
            Case "2"
                tmpsql += " where Cashier='Y'"
        End Select
        tmpsql += " and Effective='Y'"
        tmpsql += " and AreaCode='" + Me.DropDownList3.SelectedValue + "'"
        Me.SqlDataSource1.SelectCommand = tmpsql
        Me.SqlDataSource1.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.DropDownList1.DataSourceID = Me.SqlDataSource1.ID
        Me.DropDownList1.DataTextField = "EmployeeName"
        Me.DropDownList1.DataValueField = "EmployeeNo"
        Me.DropDownList1.DataBind()
        Me.SqlDataSource2.SelectCommand = tmpsql
        Me.SqlDataSource2.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.DropDownList2.DataSourceID = Me.SqlDataSource2.ID
        Me.DropDownList2.DataTextField = "EmployeeName"
        Me.DropDownList2.DataValueField = "EmployeeNo"
        Me.DropDownList2.DataBind()
    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        'DispGrid()
    End Sub

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        RememberOldValues()
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何客戶!!")
            Exit Sub
        End If
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i As Integer
        tmpsql = "update MMSCustomers "
        Select Case Me.RadioButtonList1.SelectedValue
            Case "1"
                tmpsql += " set Salesman='" + Me.DropDownList2.SelectedValue + "'"
            Case "2"
                tmpsql += " set Cashier='" + Me.DropDownList2.SelectedValue + "'"
        End Select
        tmpsql += " where CustomerNo in("
        For i = 0 To categoryIDList.Count - 1
            tmpsql += "'" + categoryIDList.Item(i) + "'"
            If i <> categoryIDList.Count - 1 Then
                tmpsql += ","
            End If
        Next
        tmpsql += ")"
        'modDB.InsertSignRecord("AZITEST", tmpsql, My.User.Name)
        EXE_SQL(tmpsql)
        'add in 20141028
        tmpsql = "update MMSContract "
        Select Case Me.RadioButtonList1.SelectedValue
            Case "1"
                tmpsql += " set Salesman='" + Me.DropDownList2.SelectedValue + "'"
            Case "2"
                tmpsql += " set Cashier='" + Me.DropDownList2.SelectedValue + "'"
        End Select
        tmpsql += " where CustomerNo in("
        For i = 0 To categoryIDList.Count - 1
            tmpsql += "'" + categoryIDList.Item(i) + "'"
            If i <> categoryIDList.Count - 1 Then
                tmpsql += ","
            End If
        Next
        tmpsql += ")"
        'modDB.InsertSignRecord("AZITEST", tmpsql, My.User.Name)
        EXE_SQL(tmpsql)
        modUtil.showMsg(Me.Page, "訊息", "批次更新完成!!")
        'DispGrid()
    End Sub




    Protected Sub GridView1_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        RememberOldValues()
        GridView1.PageIndex = e.NewPageIndex
        Me.GridView1.DataSource = Session("BatchGrid")
        Me.GridView1.DataBind()
        RePopulateValues()
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
        Me.Button1.Enabled = True
    End Sub

    Protected Sub Button3_Click(sender As Object, e As System.EventArgs) Handles Button3.Click
        Session("CHECKED_ITEMS") = Nothing
        Me.Button1.Enabled = False
        RePopulateValues()
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

    Protected Sub DropDownList3_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList3.SelectedIndexChanged
        disp_emp()
        'DispGrid()
    End Sub

    Protected Sub Button5_Click(sender As Object, e As System.EventArgs) Handles Button5.Click
        DispGrid()
        If Me.GridView1.Rows.Count > 0 Then
            Me.Button2.Enabled = True
            Me.Button3.Enabled = True
            Me.Button1.Enabled = False
        Else
            Me.Button2.Enabled = False 
            Me.Button3.Enabled = False
            Me.Button1.Enabled = False
        End If
    End Sub

    Protected Sub CheckBox2_CheckedChanged(sender As Object, e As System.EventArgs)
        RememberOldValues()
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        If categoryIDList.Count > 0 Then
            Me.Button1.Enabled = True
        Else
            Me.Button1.Enabled = False
        End If
    End Sub
End Class
