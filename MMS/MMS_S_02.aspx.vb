Imports System.Data
Partial Class MMS_MMS_S_02
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String
    Dim tt1, tt2 As DataTable

    Private Sub dispArea()
        Dim tmpsql As String
        tmpsql = "SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea]"
        tt1 = get_DataTable(tmpsql)
        Me.DropDownList1.DataTextField = "AreaName"
        Me.DropDownList1.DataValueField = "AreaCode"
        DropDownList1.DataSource = tt1
        DropDownList1.DataBind()

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
            End If

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            modDB.SetGridViewStyle(Me.grdList)          '* 套用gridview樣式
            modDB.SetFields("客戶代號", "客戶代號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("客戶名稱", "客戶名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("統一編號", "統一編號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("發票號碼", "發票號碼", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("發票日期", "發票日期", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("合約編號", "合約編號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("合約編號", "合約編號", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("品名", "品名", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("傳票號碼", "傳票號碼", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("項目金額", "項目金額", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("立帳金額", "立帳金額", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("收款金額", "收款金額", grdList, HorizontalAlign.Left, "", False)
            'modDB.SetFields("收款日期", "收款日期", grdList, HorizontalAlign.Left, "", False)
            '--------------------------------------------------* 加入連結欄位
            Me.grdList.RowStyle.Height = 18

            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 7, Me.ViewState("EmptyTable"))
            Me.MsgLabel.Text = Request("msg")

            dispArea()
            'If Me.DropDownList1.Items.Count > 0 Then
            'dispEmployee()
            'End If

            'modDB.InsertSignRecord("AziTest", "AreaCode=" & AreaCode, My.User.Name) 'AZITEST
            'If AreaCode = "" Then
            '    Me.DropDownList1.Enabled = True
            'Else
            '    Me.DropDownList1.Enabled = False
            '    Dim i As Integer
            '    For i = 0 To Me.DropDownList1.Items.Count - 1
            '        If Me.DropDownList1.Items(i).Value = AreaCode Then
            '            Me.DropDownList1.SelectedIndex = i
            '        End If
            '    Next
            '    Me.dispEmployee()
            'End If
            Dim AreaCode As String = ""
            AreaCode = GetAreaCode()

            If Me.DropDownList1.Items.Count > 0 Then
                If AreaCode <> "" Then
                    Dim i As Integer
                    For i = 0 To Me.DropDownList1.Items.Count - 1
                        If Me.DropDownList1.Items(i).Value = AreaCode Then
                            Me.DropDownList1.SelectedIndex = i
                        End If
                    Next
                End If
                '
                dispEmployee()
            End If


            'If Session("QryField") Is Nothing Then
            '    Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 13, Me.ViewState("EmptyTable"))
            'Else
            '    Call ReflashQryData()
            'End If
            Me.txtDtFrom0.Focus()
        End If

        'dispEmployee()
        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.item_defList.SelectCommand = Me.txtSql.Text
        '--------------------------------------------------* 設定GridView資料來源ID


        modUtil.SetDateObj(Me.txtDtFrom0, False, Me.txtDtTo0, False)
        modUtil.SetDateObj(Me.txtDtTo0, False, Nothing, False)

        Me.txtDtFrom1.Attributes.Add("onkeypress", "return ToUpper()")
        Me.txtDtFrom2.Attributes.Add("onkeypress", "return ToUpper()")
        'Me.txtDtFrom3.Attributes.Add("onkeypress", "return ToUpper()")
        Me.ScriptManager1.RegisterPostBackControl(Me.ExcelButton)
    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        dispEmployee()
    End Sub

    '顯示區域所屬業務員收款員
    Private Sub dispEmployee()
        Dim tmpsql As String

        tmpsql = "SELECT EmployeeNo, EmployeeNo + '-' + EmployeeName AS EmployeeName FROM MMSEmployee"
        tmpsql += " WHERE Effective = 'Y' and AreaCode='" + Me.DropDownList1.SelectedValue + "'"
        tmpsql += " and Salesman = 'Y'"
        tt2 = get_DataTable(tmpsql)
        Me.DropDownList2.DataTextField = "EmployeeName"
        Me.DropDownList2.DataValueField = "EmployeeNo"
        DropDownList2.DataSource = tt2
        DropDownList2.DataBind()

    End Sub


    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub ExcelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExcelButton.Click
        grdList.AllowPaging = False
        grdList.AllowSorting = False
        grdList.DataBind()
        modUtil.GridView2Excel("MMS_S_02.xls", Me.grdList)
        grdList.AllowPaging = True
        grdList.AllowSorting = True
        grdList.DataBind()
    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub QueryButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QueryButton.Click
        Call ShowGrid()
        Me.MsgLabel.Text = ""
        Session("employeeQ") = Me.item_defList.SelectCommand
    End Sub


    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        Dim tmpstr As String, sFields As New Hashtable
        tmpstr = ""
        tmpstr += Me.txtDtFrom0.Text
        tmpstr += Me.txtDtFrom1.Text
        tmpstr += Me.txtDtFrom2.Text
        'tmpstr += Me.txtDtFrom3.Text
        tmpstr += Me.txtDtTo0.Text
        If Me.CheckBox2.Checked = True Then
            tmpstr += " "
        End If
        If tmpstr = "" Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "您未輸入任何查詢條件!!")
            Exit Sub
        End If
        Dim tmpsql As String
        'EXEC MMS_S02 發票日期(起),發票日期(迄),客戶代號(起),客戶代號(迄),業務員,合約編號,1已收款未收款全部
        'tmpsql = "EXEC MMS_S02 '2013/09/01','2013/09/30','','','','','3'"
        tmpsql = "EXEC MMS_S02 "
        If Me.txtDtFrom0.Text <> "" Then '發票日期(起)
            tmpsql += "'" + Me.txtDtFrom0.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtTo0.Text <> "" Then '發票日期(迄)
            tmpsql += "'" + Me.txtDtTo0.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtFrom1.Text <> "" Then '客戶代號(起)
            tmpsql += "'" + Me.txtDtFrom1.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.txtDtFrom2.Text <> "" Then '客戶代號(迄)
            tmpsql += "'" + Me.txtDtFrom2.Text + "',"
        Else
            tmpsql += "'',"
        End If
        If Me.CheckBox2.Checked = True Then '業務員
            tmpsql += "'" + Me.DropDownList2.SelectedValue + "',"
        Else
            tmpsql += "'',"
        End If
        'If Me.txtDtFrom3.Text <> "" Then '合約編號
        '    tmpsql += "'" + Me.txtDtFrom3.Text + "',"
        'Else
        '    tmpsql += "'',"
        'End If
        'tmpsql += "'" + Me.RadioButtonList1.SelectedValue + "',"
        'Add in 20140911
        If (GetAreaCode() <> "") And (GetAreaCode() <> "F") Then
            tmpsql += "'" + GetAreaCode() + "'"
        Else
            tmpsql += "''"
        End If
        'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql, My.User.Name) 'AZITEST
        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        Me.item_defList.SelectCommand = tmpsql
        Me.txtSql.Text = tmpsql '* 保留SQL
        '--------------------------------------------------* 設定GridView資料來源ID
        Me.grdList.DataSource = Nothing
        Me.grdList.DataSourceID = Me.item_defList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 7, Me.ViewState("EmptyTable"))

            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
            Me.ExcelButton.Enabled = False

        Else
            Me.ExcelButton.Enabled = True
        End If
        sFields.Add("txtSql", tmpsql)
        Me.ViewState("QryField") = sFields
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
            Response.Redirect("Employee_001.aspx?FormMode=edit&STNID=" & .Cells(0).Text)
        End With

        Session("iCurPage") = grdList.PageIndex
    End Sub
#End Region

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub ClearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearButton.Click
        Me.txtDtFrom0.Text = ""
        Me.txtDtFrom1.Text = ""
        Me.txtDtFrom2.Text = ""
        'Me.txtDtFrom3.Text = ""
        Me.txtDtTo0.Text = ""
        Me.DropDownList1.SelectedIndex = 0
        'Me.dispEmployee()
        Me.CheckBox2.Checked = False
        Me.DropDownList1.Enabled = False
        Me.DropDownList2.Enabled = False
        Me.ExcelButton.Enabled = False
        'Me.RadioButtonList1.SelectedIndex = 2
        'Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 15, Me.ViewState("EmptyTable"))
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


    Protected Sub CheckBox2_CheckedChanged(sender As Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        'Me.DropDownList1.Enabled = Me.CheckBox2.Checked
        If (GetAreaCode() = "") And (GetAreaCode() = "F") Then 'modify in 20140905
            Me.DropDownList1.Enabled = Me.CheckBox2.Checked
        End If
        Me.DropDownList2.Enabled = Me.CheckBox2.Checked
    End Sub

    Protected Sub txtDtFrom1_TextChanged(sender As Object, e As System.EventArgs) Handles txtDtFrom1.TextChanged
        If Me.txtDtFrom2.Text.Trim = "" Then
            Me.txtDtFrom2.Text = Me.txtDtFrom1.Text
        End If
        txtDtFrom2.Focus()
    End Sub

    Private Function GetAreaCode() As String
        ' Dim Areacode As String = ""
        GetAreaCode = modUnset.GetAreaCode(My.User.Name).Trim
    End Function

    'Dim tmpsql As String
    'Dim STNID As String = Request.Cookies("STNID").Value
    'tmpsql = "select * from MECHSTNM where STNID='" + STNID + "'"
    'If get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "    " Or get_DataTable(tmpsql).Rows(0).Item("STNUID") <> "4601" Then
    '    tmpsql = "select * from MMSArea where UNITCODE='" + get_DataTable(tmpsql).Rows(0).Item("STNUID") + "'"
    '    Dim tmparea As String
    '    Try
    '        GetAreaCode = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
    '    Catch ex As Exception
    '        GetAreaCode = ""
    '    End Try
    'Else
    '    GetAreaCode = ""
    'End If



End Class
