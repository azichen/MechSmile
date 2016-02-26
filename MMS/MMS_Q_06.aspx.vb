Imports System.Data
Imports Microsoft.Reporting.WebForms


Partial Class MMS_Q_06
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

            'modDB.SetGridViewStyle(Me.GridView1)          '* 套用gridview樣式
            'DispGridE()
            '--------------------------------------------------* 加入連結欄位

            Me.lblMsg.Text = Request("msg")
            Me.txtDtFrom0.Attributes.Add("onkeypress", "return ToUpper()")
            Me.txtDtTo0.Attributes.Add("onkeypress", "return ToUpper()")
            txtDtFrom.Focus()
        End If
        modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
        modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
        modUtil.SetDateObj(Me.txtDtFrom1, False, Me.txtDtTo1, False)
        modUtil.SetDateObj(Me.txtDtTo1, False, Nothing, False)
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


    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        If txtDtFrom.Text + txtDtTo.Text + txtDtFrom1.Text + txtDtTo1.Text + txtDtFrom0.Text + txtDtTo0.Text = "" Then
            modUtil.showMsg(Me.Page, "訊息", "尚未輸入任何查詢條件!!")
            Exit Sub
        End If
        'modDB.InsertSignRecord("AziTest", "MMSQ06_btnQry_Click", My.User.Name) 'AZITEST
        Call ShowGrid()
        Me.lblMsg.Text = ""
        Session("CHECKED_ITEMS") = Nothing
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


    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        'modDB.InsertSignRecord("AziTest", "MMSQ06_ShowGrid", My.User.Name) 'AZITEST
        Dim sSql As String
        sSql = "select"
        sSql += " 0 as checkf,"
        sSql += " 0 as sn,"
        sSql += " *,convert(varchar, InvoiceDate, 111) as InvoiceDate1,a.sn as snM,"
        sSql += " REPLACE(CONVERT(varchar(20), (CAST(Amount AS money)), 1), '.00', '') as Amount1,"
        sSql += " case a.Status"
        sSql += " when 'Y' then '有效'"
        sSql += " when 'N' then '申請作廢'"
        sSql += " when 'V' then '作廢'"
        sSql += " end as StatusN,"
        sSql += " case a.InvoiceType"
        sSql += " when '1' then '收銀發票'"
        sSql += " when '2' then '手開發票'"
        sSql += " end as InvoiceTypeM"
        sSql += " from MMSInvoiceM a"
        sSql += " left outer join MMSInvoiceApplyM b"
        sSql += " on a.ApplyDocNo = b.DocNo"
        sSql += " where a.InvoiceNo <> 'XXXXXXXXX'"
        If Me.txtDtFrom.Text <> "" Then
            sSql += " and a.InvoiceDate >= '" + Me.txtDtFrom.Text + "'"
        End If
        If Me.txtDtTo.Text <> "" Then
            sSql += " and a.InvoiceDate <= '" + Me.txtDtTo.Text + "'"
        End If
        If Me.txtDtFrom1.Text <> "" Then
            sSql += " and b.ApplyDate >= '" + Me.txtDtFrom1.Text + "'"
        End If
        If Me.txtDtTo1.Text <> "" Then
            sSql += " and b.ApplyDate <= '" + Me.txtDtTo1.Text + "'"
        End If
        If Me.txtDtFrom0.Text <> "" Then
            sSql += " and a.InvoiceNo >= '" + Me.txtDtFrom0.Text + "'"
        End If
        If Me.txtDtTo0.Text <> "" Then
            sSql += " and a.InvoiceNo <= '" + Me.txtDtTo0.Text + "'"
        End If

        sSql += " and InvoiceNo <> ''"
        sSql += " and a.STATUS IN ('V','N')"
        'If CheckBoxList1.Items(0).Selected Or CheckBoxList1.Items(1).Selected Or CheckBoxList1.Items(2).Selected Then
        'sSql += " and a.STATUS IN ("
        'If CheckBoxList1.Items(0).Selected Then '有效
        '    If CheckBoxList1.Items(1).Selected Then '作廢
        '        If CheckBoxList1.Items(2).Selected Then '申請作廢
        '            sSql += "'Y','V','N'"
        '        Else
        '            sSql += "'Y','V'"
        '        End If
        '    Else
        '        If CheckBoxList1.Items(2).Selected Then '申請作廢
        '            sSql += "'Y','N'"
        '        Else
        '            sSql += "'Y'"
        '        End If

        '    End If
        'Else
        '    If CheckBoxList1.Items(1).Selected Then '作廢
        '        If CheckBoxList1.Items(2).Selected Then '申請作廢
        '            sSql += "'V','N'"
        '        Else
        '            sSql += "'V'"
        '        End If
        '    Else
        '        If CheckBoxList1.Items(2).Selected Then '申請作廢
        '            sSql += "'N'"
        '        End If
        '    End If
        'End If
        'sSql += ")"
        'End If

        'LOGSQL(sSql) 'azitest

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
        Me.GridView4.DataSource = Nothing
        Me.GridView4.DataBind()
        Me.GridView1.Visible = True
        Me.GridView3.Visible = True
        Me.GridView4.Visible = True
        '
        If Me.GridView1.Rows.Count = 0 Then
            'Call modDB.ShowEmptyDataGridHeader(Me.GridView1, 0, 10, Me.ViewState("EmptyTable"))
            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        End If
    End Sub

    Private Function check_printed(ByVal ss As String) As Boolean
        Dim tmpsql As String
        tmpsql = "select count(*) from MMSInvoicePrint where InvoiceNo='" + ss + "'"
        If get_DataTable(tmpsql).Rows(0).Item(0) > 0 Then
            check_printed = True
        Else
            check_printed = False
        End If
    End Function


#Region "GridView Event"

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        RememberOldValues()
        GridView1.PageIndex = e.NewPageIndex
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
        RePopulateValues()
    End Sub

    Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit
        Me.GridView1.EditIndex = -1
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()

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

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Me.GridView1.EditIndex = e.NewEditIndex
        'Dim row = GridView1.Rows(Me.GridView1.EditIndex)
        'modUtil.SetDateObj((CType((row.Cells(9).Controls(0)), TextBox)), False, Nothing, False)
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating

        'Retrieve the table from the session object.
        Dim dt As DataTable = CType(Session("BatchGrid1"), DataTable)

        'Update the values.
        Dim row = GridView1.Rows(Me.GridView1.EditIndex)
        Dim ss As String
        If (CType((row.Cells(11).Controls(0)), TextBox)).Text.Trim <> "" Then
            Try
                Dim dd As Date = (CType((row.Cells(11).Controls(0)), TextBox)).Text.Trim
            Catch ex As Exception
                modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "日期格式錯誤!!")
                Exit Sub
            End Try
        End If
        dt.Rows(row.DataItemIndex)("InvoiceDate") = (CType((row.Cells(11).Controls(0)), TextBox)).Text
        dt.Rows(row.DataItemIndex)("InvoiceNo") = (CType((row.Cells(12).Controls(0)), TextBox)).Text

        'Reset the edit index.
        GridView1.EditIndex = -1

        'Bind data to the GridView control.

        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim tmpsql As String
        tmpsql = "select * from MMSInvoiceD where InvoiceNo='" + Me.GridView1.SelectedRow.Cells(3).Text + "'"
        'modDB.InsertSignRecord("AziTest", "GridView1_SelectedIndexChanged InvoiceNo= " & Me.GridView1.SelectedRow.Cells(3).Text, My.User.Name) 'AZITEST
        Session("BatchGrid3") = get_DataTable(tmpsql)
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()

        tmpsql = "select * from MMSInvoicePrint where InvoiceNo='" + Me.GridView1.SelectedRow.Cells(3).Text + "'"
        Session("BatchGrid4") = get_DataTable(tmpsql)
        Me.GridView4.DataSource = Session("BatchGrid4")
        Me.GridView4.DataBind()
    End Sub
#End Region

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

    Protected Sub ClearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearButton.Click
        txtDtFrom.Text = ""
        txtDtTo.Text = ""
        txtDtFrom0.Text = ""
        txtDtTo0.Text = ""
        CheckBoxList1.Enabled = False
        CheckBoxList1.Items(0).Selected = False
        CheckBoxList1.Items(1).Selected = False
        CheckBoxList1.Items(2).Selected = False
        Me.GridView1.Visible = False
        Me.GridView3.Visible = False
        Me.GridView4.Visible = False
    End Sub

    'Protected Sub LOGSQL(ByVal datastr As String)
    '    Dim TMPSTR As String = datastr
    '    While (TMPSTR.Length) > 0
    '        If TMPSTR.Length > 95 Then
    '            modDB.InsertSignRecord("AziTest", TMPSTR.Substring(0, 95), My.User.Name) 'AZITEST
    '            TMPSTR = TMPSTR.Substring(95, TMPSTR.Length - 95)
    '        Else
    '            modDB.InsertSignRecord("AziTest", TMPSTR.Substring(0, TMPSTR.Length - 1), My.User.Name) 'AZITEST
    '            TMPSTR = ""
    '        End If
    '    End While
    'End Sub
End Class
