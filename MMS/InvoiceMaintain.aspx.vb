Imports System.Data
Imports Microsoft.Reporting.WebForms


Partial Class MMS_InvoiceMaintain
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
        Dim sSql, tmpstr As String
        sSql = "select"
        sSql += " 0 as checkf,"
        sSql += " 0 as sn,"
        sSql += " *,convert(varchar, InvoiceDate, 111) as InvoiceDate1,a.sn as snM,a.Status,"
        sSql += " REPLACE(CONVERT(varchar(20), (CAST(Amount AS money)), 1), '.00', '') as Amount1,"
        sSql += " case a.Status when 'Y' then '有效' when 'N' then '申請作廢' when 'R' then '作廢核准' when 'V' then '作廢' end as StatusN,"        
        sSql += " case a.InvoiceType"
        sSql += " when '1' then '收銀發票'"
        sSql += " when '2' then '手開發票'"
        sSql += " end as InvoiceTypeM,ISNULL(EINV,'') AS EINV"
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
        tmpstr = ""
        Select Case Me.RadioButtonList1.SelectedIndex
            Case 0 '已審核未開立=未給號
                sSql += " and InvoiceNo = ''"
            Case 1 '已開立=已給號有列印紀錄 => 不一定已列印
                sSql += " and InvoiceNo <> ''"
                'sSql += " and (PRNDT>'2000/01/01') " ' MODIFY BY AZI 
                'sSql += " and (select count(*) from MMSInvoicePrint b where b.InvoiceNo=a.InvoiceNo) > 0"
                'mark in 20160130
                'Case 2 '已審核未列印=有給號無列印紀錄
                '    sSql += " and InvoiceNo <> '' AND a.InvoiceType<>'2' "
                '    sSql += " and ((PRNDT is null) or (PRNDT <'2000/01/01')) " ' MODIFY BY AZI

                'sSql += " and (select count(*) from MMSInvoicePrint b where b.InvoiceNo=a.InvoiceNo) = 0"
        End Select

        'add in 20160130
        Select Case Me.RadioButtonList2.SelectedIndex
            Case 1 '有效
                sSql += " and a.STATUS = 'Y'"
            Case 2 '作廢
                sSql += " and a.STATUS = 'V'"
            Case 3 '申請作廢
                sSql += " and a.STATUS = 'N'"
            Case 4 '作廢核准
                sSql += " and a.STATUS = 'R'"
        End Select

        'If CheckBoxList1.Items(0).Selected Or CheckBoxList1.Items(1).Selected Or CheckBoxList1.Items(2).Selected Then
        '    sSql += " and a.STATUS IN ("
        '    If CheckBoxList1.Items(0).Selected Then '有效
        '        If CheckBoxList1.Items(1).Selected Then '作廢
        '            If CheckBoxList1.Items(2).Selected Then '申請作廢
        '                sSql += "'Y','V','N'"
        '            Else
        '                sSql += "'Y','V'"
        '            End If
        '        Else
        '            If CheckBoxList1.Items(2).Selected Then '申請作廢
        '                sSql += "'Y','N'"
        '            Else
        '                sSql += "'Y'"
        '            End If

        '        End If
        '    Else
        '        If CheckBoxList1.Items(1).Selected Then '作廢
        '            If CheckBoxList1.Items(2).Selected Then '申請作廢
        '                sSql += "'V','N'"
        '            Else
        '                sSql += "'V'"
        '            End If
        '        Else
        '            If CheckBoxList1.Items(2).Selected Then '申請作廢
        '                sSql += "'N'"
        '            End If
        '        End If
        '    End If
        '    sSql += ")"
        'End If

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
        '
        Me.Button1.Visible = False
        Me.Button2.Visible = False
        Me.Button3.Visible = False
        Me.Button4.Visible = False
        Me.Button5.Visible = False
        Me.Button6.Visible = False
        Me.Button7.Visible = False
        Me.Button8.Visible = False
        '
        If Me.GridView1.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "查無任何資料!!")
            
            Exit Sub
        Else
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
            Select Case Me.RadioButtonList1.SelectedIndex
                Case 0 '未開立
                    Me.Button1.Visible = True
                    Me.Button2.Visible = True
                    Me.Button5.Visible = True '取消審核
                    Me.Button6.Visible = True '發票給號
                Case 1 '已開立
                    'azitemp
                    Select Case Me.RadioButtonList2.SelectedIndex
                        Case 1 '有效
                            '
                        Case 2 '作廢
                            '
                        Case 3 '申請作廢
                            Me.Button7.Visible = True '核准作廢
                            Me.Button8.Visible = True '取消作廢申請
                        Case 4 '作廢核准
                            Me.Button3.Visible = True '作廢
                    End Select
            End Select
            '
            'Me.Button3.Enabled = True '作廢
            'Me.Button4.Enabled = True '列印發票
            'Me.Button5.Enabled = True '取消審核
            'Me.Button6.Enabled = True '發票給號
        End If
        '
        Me.GridView1.Visible = True
        Me.GridView3.Visible = True
        Me.GridView4.Visible = True
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
        tmpsql = "select * from MMSInvoiceD where InvoiceNo='" + Me.GridView1.SelectedRow.Cells(4).Text + "'"
        Session("BatchGrid3") = get_DataTable(tmpsql)
        Me.GridView3.DataSource = Session("BatchGrid3")
        Me.GridView3.DataBind()

        tmpsql = "select * from MMSInvoicePrint where InvoiceNo='" + Me.GridView1.SelectedRow.Cells(4).Text + "'"
        Session("BatchGrid4") = get_DataTable(tmpsql)
        Me.GridView4.DataSource = Session("BatchGrid4")
        Me.GridView4.DataBind()
    End Sub
#End Region


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim categoryIDList As ArrayList = New ArrayList
        Dim tt As DataTable
        Dim i As Integer
        Session("CHECKED_ITEMS") = Nothing
        tt = Session("BatchGrid1")
        For i = 0 To tt.Rows.Count - 1
            categoryIDList.Add(tt.Rows(i).Item("sn").ToString)
        Next
        Try
            If (categoryIDList.Count > 0) Then
                Session("CHECKED_ITEMS") = categoryIDList
            End If
        Catch ex As Exception

        End Try
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
        RePopulateValues()
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Session("CHECKED_ITEMS") = Nothing
        RePopulateValues()
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Dim VCONTRACTNO, VINVYM As String

        RememberOldValues()
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何資料!!")
            Exit Sub
        End If
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i, j As Integer
        Dim tt, tt1 As DataTable
        Dim tmpstr As String = ""
        tt = Session("BatchGrid1")
        For i = 0 To categoryIDList.Count - 1
            If tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") <> "" Then
                modDB.InsertSignRecord("(InvoiceMaintain)_發票作廢", "InvoiceNo =" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo"), My.User.Name)
                '
                tmpsql = "update MMSInvoiceM set Status = 'V',VOIDDT = GETDATE()"
                tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                EXE_SQL(tmpsql)
                tmpsql = "update MMSInvoiceApplyD set Effective = 'V'"
                tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                EXE_SQL(tmpsql)
                '20151007 需再增加電子發票作廢處理 
                'modDB.InsertSignRecord("AziTest", "InvoiceMaintain InvoiceDate=" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate1"), My.User.Name) 'AZITEST
                If tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate1") >= "2015/10/01" Then
                    '作廢日期處理 
                    Dim MIGVOIDDT As String
                    If DateTime.Now.ToString("yyyy/MM/dd") <= tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate1") Then
                        MIGVOIDDT = DateTime.Parse(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate1")).AddDays(1).ToString("yyyyMMdd")
                    Else
                        MIGVOIDDT = DateTime.Now.ToString("yyyyMMdd")
                    End If
                    '
                    tmpsql = "SELECT ISNULL(MAX(MIGID),'000000000000') AS MIGDATENO FROM INVMIGREC WHERE SUBSTRING(MIGID,1,8)='" + MIGVOIDDT + "'"
                    '
                    tt1 = get_DataTable(tmpsql)
                    Dim vMIGID As String = MIGVOIDDT + (Convert.ToInt32(tt1.Rows(0).Item("MIGDATENO").substring(8, 4)) + 1).ToString("0000")
                    '
                    tmpsql = "INSERT INTO INVMIGREC (STNID,MIGID,INVNUM,INVDATE,INVTIME,MIGKIND,DATAKIND,MIGUPLOAD) VALUES "
                    tmpsql += " ('040301','" + vMIGID + "','" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                    tmpsql += " ,'" + MIGVOIDDT + "','" + DateTime.Now.ToString("HHmmss") + "','C0501','MECHINE','')"
                    EXE_SQL(tmpsql)
                    modDB.InsertSignRecord("(InvoiceMaintain)_發票作廢", "MIGADD OK: InvoiceNo =" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + " MIGID=" + vMIGID, My.User.Name)
                End If
                ''20140609 add by Azi 增加 update 發票開立期別處理
                'VCONTRACTNO = ""
                'VINVYM = ""
                'tmpsql = "SELECT ISNULL(CONTRACTNO,'') AS CONTRACTNO,ISNULL(INVOICEPERIOD,'') AS INVOICEPERIOD FROM MMSInvoiceApplyD"
                'tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                'tmpsql += "   and ContractNo<>''" 'modify in 20140925
                ''LOGSQL(tmpsql)
                'tt1 = get_DataTable(tmpsql)
                ''modDB.InsertSignRecord("AziTest", "InvoiceMaintain tt.Rows.Count=" + tt1.Rows.Count.ToString, My.User.Name) 'AZITEST
                'If tt1.Rows.Count > 0 Then
                '    'modDB.InsertSignRecord("azitest", "作廢合約期別:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " 筆數 = " & tt1.Rows.Count.ToString, My.User.Name)
                '    j = 0
                '    '20151223 一張發票可能有多筆保養合約的資料，增加 For loop處理
                '    For j = 0 To (tt1.Rows.Count - 1)
                '        'modDB.InsertSignRecord("azitest", " j = " & j.ToString, My.User.Name)
                '        VCONTRACTNO = tt1.Rows(j).Item("CONTRACTNO")
                '        VINVYM = tt1.Rows(j).Item("INVOICEPERIOD")
                '        'modDB.InsertSignRecord("azitest", "作廢合約期別:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " CONTRACTNO =" & VCONTRACTNO & " INVYM =" & VINVYM, My.User.Name)
                '        If (VCONTRACTNO.Trim <> "") And (VINVYM.Trim <> "") Then
                '            modDB.InsertSignRecord("(InvoiceMaintain)_發票作廢", "期別處理:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " CONTRACTNO =" & VCONTRACTNO & " INVYM =" & VINVYM, My.User.Name)
                '            tmpsql = "update MMSInvoiceA set YN = 'N',INVOICEDATE='',INVOICENO=''" 'MODIFY IN 20160107
                '            tmpsql += " where CONTRACTNO='" + VCONTRACTNO + "' AND INVOICEPERIOD='" + VINVYM + "'"
                '            '
                '            'LOGSQL(tmpsql) 'azitest
                '            EXE_SQL(tmpsql)
                '            '增加UPDATE合約最後開立期別_20150901 
                '            UpdateLastInvoiceDate(VCONTRACTNO)
                '        End If
                '    Next
                'End If
                tmpstr += tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + " \n"
            End If
        Next
        If tmpstr <> "" Then
            tmpstr = "發票號碼:" + " \n" + tmpstr + "作廢成功!!"
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
        Else
            modUtil.showMsg(Me.Page, "訊息", "無可作廢發票!!")
        End If
        '
        Call ShowGrid()
        Me.lblMsg.Text = ""
        Session("CHECKED_ITEMS") = Nothing        
    End Sub

    'Private Sub UpdateRow(ByVal sn As String)
    '    Dim tt As DataTable
    '    Dim tmpsql As String
    '    tt = Session("BatchGrid1")
    '    Dim i As Integer
    '    For i = 0 To tt.Rows.Count - 1
    '        If tt.Rows(i).Item("sn") = sn Then
    '            tmpsql = "update MMSInvoiceApplyD set "
    '            tmpsql += " InvoiceDate='" + tt.Rows(i).Item("InvoiceDate") + "',"
    '            tmpsql += " InvoiceNo='" + tt.Rows(i).Item("InvoiceNo") + "'"
    '            tmpsql += " where DocNo='" + tt.Rows(i).Item("DocNo") + "' and Serial=" + tt.Rows(i).Item("Serial").ToString
    '            EXE_SQL(tmpsql)
    '            Exit For
    '        End If
    '    Next
    'End Sub

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

    Protected Sub RadioButtonList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButtonList1.SelectedIndexChanged
        If RadioButtonList1.SelectedIndex = 1 Then
            RadioButtonList2.Visible = True
        Else
            RadioButtonList2.Visible = False
        End If
    End Sub

    Private Function changeDate(ByVal dd As String) As String
        Dim tdate As Date
        tdate = dd
        changeDate = LTrim(RTrim(tdate.Year.ToString)) + "/"
        If tdate.Month < 10 Then changeDate += "0"
        changeDate += LTrim(RTrim(tdate.Month.ToString)) + "/"
        If tdate.Day < 10 Then changeDate += "0"
        changeDate += LTrim(RTrim(tdate.Day.ToString))
    End Function

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        RememberOldValues()
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何資料!!")
            Exit Sub
        End If
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i, j, k, m As Integer
        Dim tt, tmptable As DataTable
        Dim tmpstr As String = ""
        Dim tmpInvNO As String = ""
        'Dim applydocno As String

        Dim ITEM(5), PRICE(5), QTY(5), AMT(5) As String

        Dim tts As String = ""
        Dim Itemfirst As Boolean = True

        tt = Session("BatchGrid1")
        tmpsql = "delete from MMSInvoiceTMP"
        EXE_SQL(tmpsql)
        For i = 0 To categoryIDList.Count - 1
            If tt.Rows(categoryIDList.Item(i)).Item("InvoiceTypeM") = "收銀發票" And tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") <> "" And tt.Rows(categoryIDList.Item(i)).Item("StatusN") <> "作廢" Then
                tmpsql = "insert into [MMSInvoicePrint]([InvoiceNo],[InvoiceDate],[PrintDateTime])"
                tmpsql += " values('" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "','" + changeDate(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate")) + "',GETDATE())"
                'EXE_SQL(tmpsql)
                tmpstr += tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + " \n"
                'modDB.InsertSignRecord("azitest", "tmpstr =" & tmpstr & "stnid=" & Request.Cookies("STNID").Value.ToString.Trim, My.User.Name) 'azitest

                '填入列印暫存table
                tmpsql = "delete from MMSInvoiceTMP"
                tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                'modDB.InsertSignRecord("azitest", "tmpsql=" & tmpsql, My.User.Name) 'azitest

                EXE_SQL(tmpsql)

                tmpsql = "INSERT INTO MMSInvoiceTMP "
                tmpsql += " Select '" + Request.Cookies("STNID").Value.ToString.Trim + "' as UserId,InvoiceNo,convert(varchar,year(Invoicedate)-1911)+'-'"
                tmpsql += " +SUBSTRING(Invoicedate,6,2)+'-'+SUBSTRING(Invoicedate,9,2)+' '"
                tmpsql += " +substring(INVTIME,1,2)+':'+substring(INVTIME,3,2)+':'+substring(INVTIME,5,2) As INVTM"
                tmpsql += " ,GUINumber,TitleOfInvoice,CHKNO,'' AS ItemName"
                tmpsql += " ,'' AS ITEM2,'' AS ITEM3,'' AS ITEM4,'' AS ITEM5"
                tmpsql += " ,0 AS UnitPrice,0 AS Quantity,0 AS Amount"
                tmpsql += " ,0 as price2,0 as qty2,0 as amt2,0 as price3,0 as qty3,0 as amt3"
                tmpsql += " ,0 as price4,0 as qty4,0 as amt4,0 as price5,0 as qty5,0 as amt5"
                tmpsql += " ,NotTaxAmount as InvAmount,Tax,Amount as totaleAmount,CustomerNo"
                tmpsql += " ,'一' as PageType1,'存根' as PageType2"
                tmpsql += " From MMSInvoiceM"
                tmpsql += " Where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"

                'tmpsql += " select '" + Request.Cookies("STNID").Value.ToString.Trim + "' as UserId,a.InvoiceNo"
                'tmpsql += " ,convert(varchar,year(Invoicedate)-1911)+"
                'tmpsql += " isnull('-'+substring(convert(varchar,Invoicedate,120),6,5),'')+' '+"
                'tmpsql += " substring(INVTIME convert(varchar,Invoicedate,121),12,8) as INVTIME,"
                'tmpsql += " b.GUINumber,b.TitleOfInvoice,b.CHKNO,a.Serial,a.ItemName,a.UnitPrice,"
                'tmpsql += " a.Quantity,a.Amount,b.NotTaxAmount as InvAmount,b.Tax,"
                'tmpsql += " b.Amount as totaleAmount,CustomerNo,"
                'tmpsql += " '一' as PageType1,"
                'tmpsql += " '存根' as PageType2"
                'tmpsql += " from MMSInvoiceD a left outer join MMSInvoiceM b"
                'tmpsql += " on a.InvoiceNo=b.InvoiceNo"
                'tmpsql += " where a.InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                EXE_SQL(tmpsql)
                '處理品名大於六個中文字
                tmpsql = "Select * From MMSInvoiceD"
                tmpsql += " WHERE InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                tmpsql += " ORDER BY SERIAL"
                tmptable = get_DataTable(tmpsql)
                For k = 1 To 5
                    ITEM(k) = ""
                    PRICE(k) = ""
                    QTY(k) = ""
                    AMT(k) = ""
                Next
                '
                k = 1
                For j = 0 To tmptable.Rows.Count - 1
                    'modDB.InsertSignRecord("AziTest", "j=" & j.ToString, My.User.Name) 'AZITEST
                    If k <= 5 Then '最多5行
                        If System.Text.Encoding.Default.GetBytes(tmptable.Rows(j).Item("ItemName")).Length > 12 Then
                            'modDB.InsertSignRecord("AziTest", "ITEMname length>12", My.User.Name) 'AZITEST
                            Itemfirst = True
                            For m = 1 To tmptable.Rows(j).Item("ItemName").ToString.Length
                                'modDB.InsertSignRecord("AziTest", "m=" & m.ToString, My.User.Name) 'AZITEST
                                tts = ""
                                tts += Mid(tmptable.Rows(j).Item("ItemName").ToString, m, 1)
                                If System.Text.Encoding.Default.GetBytes(tts).Length = 12 Then
                                    ITEM(k) = tts
                                    'modDB.InsertSignRecord("AziTest", "Length=12 k=" & k.ToString & "tts=" & tts & m.ToString, My.User.Name) 'AZITEST
                                    If Itemfirst Then
                                        PRICE(k) = tmptable.Rows(j).Item("UnitPrice").ToString
                                        QTY(k) = tmptable.Rows(j).Item("Quantity").ToString
                                        AMT(k) = tmptable.Rows(j).Item("Amount").ToString
                                        Itemfirst = False
                                    Else
                                        PRICE(k) = ""
                                        QTY(k) = ""
                                        AMT(k) = ""
                                    End If
                                    tts = ""
                                    k = k + 1
                                Else
                                    If j <> tmptable.Rows(i).Item("ItemName").ToString.Length Then
                                        If System.Text.Encoding.Default.GetBytes(tts + Mid(tmptable.Rows(i).Item("ItemName").ToString, j + 1, 1)).Length > 12 Then
                                            If Itemfirst Then
                                                PRICE(k) = tmptable.Rows(i).Item("UnitPrice").ToString
                                                QTY(k) = tmptable.Rows(i).Item("Quantity").ToString
                                                AMT(k) = tmptable.Rows(i).Item("Amount").ToString
                                            Else
                                                PRICE(k) = ""
                                                QTY(k) = ""
                                                AMT(k) = ""
                                            End If
                                            tts = ""
                                            k = k + 1
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            ITEM(k) = tmptable.Rows(j).Item("ItemName").ToString
                            PRICE(k) = tmptable.Rows(j).Item("UnitPrice").ToString
                            QTY(k) = tmptable.Rows(j).Item("Quantity").ToString
                            AMT(k) = tmptable.Rows(j).Item("Amount").ToString
                            Itemfirst = False
                            tts = ""
                            k = k + 1
                        End If
                    End If
                Next
                '
                '處理品名大於六個中文字
                'modDB.InsertSignRecord("AziTest", "ITEM(2)=" & ITEM(2), My.User.Name) 'AZITEST

                tmpsql = "update MMSInvoiceTMP Set ItemName='" + ITEM(1) + "',ITEM2='" + ITEM(2) + "', ITEM3='" + ITEM(3) + "'"
                tmpsql += " ,ITEM4='" + ITEM(4) + "', ITEM5='" + ITEM(5) + "'"
                tmpsql += " ,UnitPrice='" + PRICE(1) + "',Price2='" + PRICE(2) + "',price3='" + PRICE(3) + "'"
                tmpsql += " ,PRICE4='" + PRICE(4) + "',Price5='" + PRICE(5) + "'"
                tmpsql += " ,Quantity='" + QTY(1) + "',qty2='" + QTY(2) + "',QTY3='" + QTY(3) + "',qty4='" + QTY(4) + "',QTY5='" + QTY(5) + "'"
                tmpsql += " ,Amount='" + AMT(1) + "',AMT2='" + AMT(2) + "',AMT3='" + AMT(3) + "',AMT4='" + AMT(4) + "',AMT5='" + AMT(5) + "'"
                tmpsql += " Where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                EXE_SQL(tmpsql)
                'azitemp 20150225
                '        For j = 1 To tmptable.Rows(i).Item("ItemName").ToString.Length
                '            tts += Mid(tmptable.Rows(i).Item("ItemName").ToString, j, 1)
                '            If System.Text.Encoding.Default.GetBytes(tts).Length = 12 Then
                '                tmpser += 1
                '                tmpsql = "INSERT INTO MMSInvoiceTMP values("
                '                tmpsql += "'" + tmptable.Rows(i).Item("UserID").ToString + "'," '<UserID, varchar(10),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("InvoiceNo").ToString + "'," '<InvoiceNo, varchar(10),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("InvDateTime").ToString + "'," '<InvDateTime, varchar(20),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("GUINumber").ToString + "'," '<GUINumber, varchar(8),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("TitleOfInvoice").ToString + "'," '<TitleOfInvoice, nvarchar(150),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("CHKNO").ToString + "'," '<CHKNO, varchar(3),>
                '                tmpsql += tmpser.ToString + "," 'tmptable.Rows(i).Item("Serial").ToString + "," '<Serial, int,>
                '                tmpsql += "'" + tts + "'," '<ItemName, nvarchar(50),>
                '                tmpsql += tmptable.Rows(i).Item("UnitPrice").ToString + "," '<UnitPrice, money,>
                '                tmpsql += tmptable.Rows(i).Item("Quantity").ToString + "," '<Quantity, int,>
                '                tmpsql += tmptable.Rows(i).Item("Amount").ToString + "," '<Amount, int,>
                '                tmpsql += tmptable.Rows(i).Item("InvAmount").ToString + "," '<InvAmount, int,>
                '                tmpsql += tmptable.Rows(i).Item("Tax").ToString + "," '<Tax, int,>
                '                tmpsql += tmptable.Rows(i).Item("TotaleAmount").ToString + "," '<TotaleAmount, int,>
                '                tmpsql += "'" + tmptable.Rows(i).Item("CustomerNo").ToString + "'," '<CustomerNo, varchar(15),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("PageType1").ToString + "'," '<PageType1, nvarchar(10),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("PageType2").ToString + "'," '<PageType2, nvarchar(10),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("Memo1").ToString + "'," '<Memo1, nvarchar(50),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("Memo2").ToString + "'," '<Memo2, nvarchar(50),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("Memo3").ToString + "'," '<Memo3, nvarchar(50),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("Memo4").ToString + "'," '<Memo4, nvarchar(50),>
                '                tmpsql += "'" + tmptable.Rows(i).Item("Memo5").ToString + "'" '<Memo5, nvarchar(150),>
                '                tmpsql += ")"
                '                EXE_SQL(tmpsql)
                '                tts = ""
                '            Else
                '                If j <> tmptable.Rows(i).Item("ItemName").ToString.Length Then
                '                    If System.Text.Encoding.Default.GetBytes(tts + Mid(tmptable.Rows(i).Item("ItemName").ToString, j + 1, 1)).Length > 12 Then
                '                        tmpser += 1
                '                        tmpsql = "INSERT INTO MMSInvoiceTMP values("
                '                        tmpsql += "'" + tmptable.Rows(i).Item("UserID").ToString + "'," '<UserID, varchar(10),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("InvoiceNo").ToString + "'," '<InvoiceNo, varchar(10),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("InvDateTime").ToString + "'," '<InvDateTime, varchar(20),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("GUINumber").ToString + "'," '<GUINumber, varchar(8),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("TitleOfInvoice").ToString + "'," '<TitleOfInvoice, nvarchar(150),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("CHKNO").ToString + "'," '<CHKNO, varchar(3),>
                '                        tmpsql += tmpser.ToString + "," 'tmptable.Rows(i).Item("Serial").ToString + "," '<Serial, int,>
                '                        tmpsql += "'" + tts + "'," '<ItemName, nvarchar(50),>
                '                        tmpsql += tmptable.Rows(i).Item("UnitPrice").ToString + "," '<UnitPrice, money,>
                '                        tmpsql += tmptable.Rows(i).Item("Quantity").ToString + "," '<Quantity, int,>
                '                        tmpsql += tmptable.Rows(i).Item("Amount").ToString + "," '<Amount, int,>
                '                        tmpsql += tmptable.Rows(i).Item("InvAmount").ToString + "," '<InvAmount, int,>
                '                        tmpsql += tmptable.Rows(i).Item("Tax").ToString + "," '<Tax, int,>
                '                        tmpsql += tmptable.Rows(i).Item("TotaleAmount").ToString + "," '<TotaleAmount, int,>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("CustomerNo").ToString + "'," '<CustomerNo, varchar(15),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("PageType1").ToString + "'," '<PageType1, nvarchar(10),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("PageType2").ToString + "'," '<PageType2, nvarchar(10),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo1").ToString + "'," '<Memo1, nvarchar(50),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo2").ToString + "'," '<Memo2, nvarchar(50),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo3").ToString + "'," '<Memo3, nvarchar(50),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo4").ToString + "'," '<Memo4, nvarchar(50),>
                '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo5").ToString + "'" '<Memo5, nvarchar(150),>
                '                        tmpsql += ")"
                '                        EXE_SQL(tmpsql)
                '                        tts = ""
                '                    End If

                '                End If
                '            End If
                '        Next

                '        modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-2", My.User.Name) 'azitest

                '        tmpser += 1
                '        tmpsql = "INSERT INTO MMSInvoiceTMP values("
                '        tmpsql += "'" + tmptable.Rows(i).Item("UserID").ToString + "'," '<UserID, varchar(10),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("InvoiceNo").ToString + "'," '<InvoiceNo, varchar(10),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("InvDateTime").ToString + "'," '<InvDateTime, varchar(20),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("GUINumber").ToString + "'," '<GUINumber, varchar(8),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("TitleOfInvoice").ToString + "'," '<TitleOfInvoice, nvarchar(150),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("CHKNO").ToString + "'," '<CHKNO, varchar(3),>
                '        tmpsql += tmpser.ToString + "," 'tmptable.Rows(i).Item("Serial").ToString + "," '<Serial, int,>
                '        tmpsql += "'" + tts + "'," '<ItemName, nvarchar(50),>
                '        tmpsql += tmptable.Rows(i).Item("UnitPrice").ToString + "," '<UnitPrice, money,>
                '        tmpsql += tmptable.Rows(i).Item("Quantity").ToString + "," '<Quantity, int,>
                '        tmpsql += tmptable.Rows(i).Item("Amount").ToString + "," '<Amount, int,>
                '        tmpsql += tmptable.Rows(i).Item("InvAmount").ToString + "," '<InvAmount, int,>
                '        tmpsql += tmptable.Rows(i).Item("Tax").ToString + "," '<Tax, int,>
                '        tmpsql += tmptable.Rows(i).Item("TotaleAmount").ToString + "," '<TotaleAmount, int,>
                '        tmpsql += "'" + tmptable.Rows(i).Item("CustomerNo").ToString + "'," '<CustomerNo, varchar(15),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("PageType1").ToString + "'," '<PageType1, nvarchar(10),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("PageType2").ToString + "'," '<PageType2, nvarchar(10),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("Memo1").ToString + "'," '<Memo1, nvarchar(50),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("Memo2").ToString + "'," '<Memo2, nvarchar(50),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("Memo3").ToString + "'," '<Memo3, nvarchar(50),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("Memo4").ToString + "'," '<Memo4, nvarchar(50),>
                '        tmpsql += "'" + tmptable.Rows(i).Item("Memo5").ToString + "'" '<Memo5, nvarchar(150),>
                '        tmpsql += ")"
                '        EXE_SQL(tmpsql)
                '        tts = ""
                '        modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-3-A", My.User.Name) 'azitest
                '        tmpsql = "delete from MMSInvoiceTMP where  InvoiceNo='" + tmptable.Rows(i).Item("InvoiceNo").ToString + "' and Serial < 10"
                '        modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-3-B", My.User.Name) 'azitest
                '        EXE_SQL(tmpsql)
                '    End If


            End If
        Next

        'modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-1", My.User.Name) 'azitest

        '處理品名大於六個中文字
        'tmpsql = "Select * from MMSInvoiceTMP"
        'Dim tmptable As DataTable = get_DataTable(tmpsql)
        'Dim tmptable1, tmptable2 As DataTable
        'Dim tts As String
        'Dim j As Integer
        'Dim tmpser As Integer
        'For i = 0 To tmptable.Rows.Count - 1
        '    tts = ""
        '    tmpser = tmptable.Rows(i).Item("Serial") * 10
        '    If System.Text.Encoding.Default.GetBytes(tmptable.Rows(i).Item("ItemName")).Length > 12 Then
        '        For j = 1 To tmptable.Rows(i).Item("ItemName").ToString.Length
        '            tts += Mid(tmptable.Rows(i).Item("ItemName").ToString, j, 1)
        '            If System.Text.Encoding.Default.GetBytes(tts).Length = 12 Then
        '                tmpser += 1
        '                tmpsql = "INSERT INTO MMSInvoiceTMP values("
        '                tmpsql += "'" + tmptable.Rows(i).Item("UserID").ToString + "'," '<UserID, varchar(10),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("InvoiceNo").ToString + "'," '<InvoiceNo, varchar(10),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("InvDateTime").ToString + "'," '<InvDateTime, varchar(20),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("GUINumber").ToString + "'," '<GUINumber, varchar(8),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("TitleOfInvoice").ToString + "'," '<TitleOfInvoice, nvarchar(150),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("CHKNO").ToString + "'," '<CHKNO, varchar(3),>
        '                tmpsql += tmpser.ToString + "," 'tmptable.Rows(i).Item("Serial").ToString + "," '<Serial, int,>
        '                tmpsql += "'" + tts + "'," '<ItemName, nvarchar(50),>
        '                tmpsql += tmptable.Rows(i).Item("UnitPrice").ToString + "," '<UnitPrice, money,>
        '                tmpsql += tmptable.Rows(i).Item("Quantity").ToString + "," '<Quantity, int,>
        '                tmpsql += tmptable.Rows(i).Item("Amount").ToString + "," '<Amount, int,>
        '                tmpsql += tmptable.Rows(i).Item("InvAmount").ToString + "," '<InvAmount, int,>
        '                tmpsql += tmptable.Rows(i).Item("Tax").ToString + "," '<Tax, int,>
        '                tmpsql += tmptable.Rows(i).Item("TotaleAmount").ToString + "," '<TotaleAmount, int,>
        '                tmpsql += "'" + tmptable.Rows(i).Item("CustomerNo").ToString + "'," '<CustomerNo, varchar(15),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("PageType1").ToString + "'," '<PageType1, nvarchar(10),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("PageType2").ToString + "'," '<PageType2, nvarchar(10),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("Memo1").ToString + "'," '<Memo1, nvarchar(50),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("Memo2").ToString + "'," '<Memo2, nvarchar(50),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("Memo3").ToString + "'," '<Memo3, nvarchar(50),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("Memo4").ToString + "'," '<Memo4, nvarchar(50),>
        '                tmpsql += "'" + tmptable.Rows(i).Item("Memo5").ToString + "'" '<Memo5, nvarchar(150),>
        '                tmpsql += ")"
        '                EXE_SQL(tmpsql)
        '                tts = ""
        '            Else
        '                If j <> tmptable.Rows(i).Item("ItemName").ToString.Length Then
        '                    If System.Text.Encoding.Default.GetBytes(tts + Mid(tmptable.Rows(i).Item("ItemName").ToString, j + 1, 1)).Length > 12 Then
        '                        tmpser += 1
        '                        tmpsql = "INSERT INTO MMSInvoiceTMP values("
        '                        tmpsql += "'" + tmptable.Rows(i).Item("UserID").ToString + "'," '<UserID, varchar(10),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("InvoiceNo").ToString + "'," '<InvoiceNo, varchar(10),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("InvDateTime").ToString + "'," '<InvDateTime, varchar(20),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("GUINumber").ToString + "'," '<GUINumber, varchar(8),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("TitleOfInvoice").ToString + "'," '<TitleOfInvoice, nvarchar(150),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("CHKNO").ToString + "'," '<CHKNO, varchar(3),>
        '                        tmpsql += tmpser.ToString + "," 'tmptable.Rows(i).Item("Serial").ToString + "," '<Serial, int,>
        '                        tmpsql += "'" + tts + "'," '<ItemName, nvarchar(50),>
        '                        tmpsql += tmptable.Rows(i).Item("UnitPrice").ToString + "," '<UnitPrice, money,>
        '                        tmpsql += tmptable.Rows(i).Item("Quantity").ToString + "," '<Quantity, int,>
        '                        tmpsql += tmptable.Rows(i).Item("Amount").ToString + "," '<Amount, int,>
        '                        tmpsql += tmptable.Rows(i).Item("InvAmount").ToString + "," '<InvAmount, int,>
        '                        tmpsql += tmptable.Rows(i).Item("Tax").ToString + "," '<Tax, int,>
        '                        tmpsql += tmptable.Rows(i).Item("TotaleAmount").ToString + "," '<TotaleAmount, int,>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("CustomerNo").ToString + "'," '<CustomerNo, varchar(15),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("PageType1").ToString + "'," '<PageType1, nvarchar(10),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("PageType2").ToString + "'," '<PageType2, nvarchar(10),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo1").ToString + "'," '<Memo1, nvarchar(50),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo2").ToString + "'," '<Memo2, nvarchar(50),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo3").ToString + "'," '<Memo3, nvarchar(50),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo4").ToString + "'," '<Memo4, nvarchar(50),>
        '                        tmpsql += "'" + tmptable.Rows(i).Item("Memo5").ToString + "'" '<Memo5, nvarchar(150),>
        '                        tmpsql += ")"
        '                        EXE_SQL(tmpsql)
        '                        tts = ""
        '                    End If

        '                End If
        '            End If
        '        Next

        '        modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-2", My.User.Name) 'azitest

        '        tmpser += 1
        '        tmpsql = "INSERT INTO MMSInvoiceTMP values("
        '        tmpsql += "'" + tmptable.Rows(i).Item("UserID").ToString + "'," '<UserID, varchar(10),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("InvoiceNo").ToString + "'," '<InvoiceNo, varchar(10),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("InvDateTime").ToString + "'," '<InvDateTime, varchar(20),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("GUINumber").ToString + "'," '<GUINumber, varchar(8),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("TitleOfInvoice").ToString + "'," '<TitleOfInvoice, nvarchar(150),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("CHKNO").ToString + "'," '<CHKNO, varchar(3),>
        '        tmpsql += tmpser.ToString + "," 'tmptable.Rows(i).Item("Serial").ToString + "," '<Serial, int,>
        '        tmpsql += "'" + tts + "'," '<ItemName, nvarchar(50),>
        '        tmpsql += tmptable.Rows(i).Item("UnitPrice").ToString + "," '<UnitPrice, money,>
        '        tmpsql += tmptable.Rows(i).Item("Quantity").ToString + "," '<Quantity, int,>
        '        tmpsql += tmptable.Rows(i).Item("Amount").ToString + "," '<Amount, int,>
        '        tmpsql += tmptable.Rows(i).Item("InvAmount").ToString + "," '<InvAmount, int,>
        '        tmpsql += tmptable.Rows(i).Item("Tax").ToString + "," '<Tax, int,>
        '        tmpsql += tmptable.Rows(i).Item("TotaleAmount").ToString + "," '<TotaleAmount, int,>
        '        tmpsql += "'" + tmptable.Rows(i).Item("CustomerNo").ToString + "'," '<CustomerNo, varchar(15),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("PageType1").ToString + "'," '<PageType1, nvarchar(10),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("PageType2").ToString + "'," '<PageType2, nvarchar(10),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("Memo1").ToString + "'," '<Memo1, nvarchar(50),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("Memo2").ToString + "'," '<Memo2, nvarchar(50),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("Memo3").ToString + "'," '<Memo3, nvarchar(50),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("Memo4").ToString + "'," '<Memo4, nvarchar(50),>
        '        tmpsql += "'" + tmptable.Rows(i).Item("Memo5").ToString + "'" '<Memo5, nvarchar(150),>
        '        tmpsql += ")"
        '        EXE_SQL(tmpsql)
        '        tts = ""
        '        modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-3-A", My.User.Name) 'azitest
        '        tmpsql = "delete from MMSInvoiceTMP where  InvoiceNo='" + tmptable.Rows(i).Item("InvoiceNo").ToString + "' and Serial < 10"
        '        modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-3-B", My.User.Name) 'azitest
        '        EXE_SQL(tmpsql)
        '    End If
        'Next

        'modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-4", My.User.Name) 'azitest

        'tmpsql = "update MMSInvoiceTMP set UnitPrice=NULL,Quantity=NULL,Amount=NULL,InvAmount=Null where Serial <> 11"
        'EXE_SQL(tmpsql)

        'tmpsql = "select distinct InvoiceNo from MMSInvoiceTMP"
        'tmptable = get_DataTable(tmpsql)
        'Dim k As Integer
        'For i = 0 To tmptable.Rows.Count - 1
        '    tmpsql = "select count(*) from MMSInvoiceTMP where InvoiceNo='" + tmptable.Rows(i).Item(0) + "'"
        '    k = get_DataTable(tmpsql).Rows(0).Item(0)
        '    For j = 0 To (6 - (k Mod 6)) - 1
        '        tmpsql = "select distinct UserID,InvoiceNo,InvDateTime,GUINumber,TitleOfInvoice,CHKNO,CustomerNo from MMSInvoiceTMP  where InvoiceNo='" + tmptable.Rows(i).Item(0) + "'"
        '        tmptable2 = get_DataTable(tmpsql)
        '        tmpsql = "INSERT INTO MMSInvoiceTMP(UserID,InvoiceNo,InvDateTime,GUINumber,TitleOfInvoice,CHKNO,CustomerNo,Serial,PageType1) values("
        '        tmpsql += "'" + tmptable2.Rows(i).Item("UserID").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("InvoiceNo").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("InvDateTime").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("GUINumber").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("TitleOfInvoice").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("CHKNO").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("CustomerNo").ToString + "'," '
        '        tmpsql += "999,'一'"
        '        tmpsql += ")"
        '        EXE_SQL(tmpsql)
        '        tmpsql = "INSERT INTO MMSInvoiceTMP(UserID,InvoiceNo,InvDateTime,GUINumber,TitleOfInvoice,CHKNO,CustomerNo,Serial,PageType1) values("
        '        tmpsql += "'" + tmptable2.Rows(i).Item("UserID").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("InvoiceNo").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("InvDateTime").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("GUINumber").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("TitleOfInvoice").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("CHKNO").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("CustomerNo").ToString + "'," '
        '        tmpsql += "999,'二'"
        '        tmpsql += ")"
        '        EXE_SQL(tmpsql)
        '        tmpsql = "INSERT INTO MMSInvoiceTMP(UserID,InvoiceNo,InvDateTime,GUINumber,TitleOfInvoice,CHKNO,CustomerNo,Serial,PageType1) values("
        '        tmpsql += "'" + tmptable2.Rows(i).Item("UserID").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("InvoiceNo").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("InvDateTime").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("GUINumber").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("TitleOfInvoice").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("CHKNO").ToString + "'," '
        '        tmpsql += "'" + tmptable2.Rows(i).Item("CustomerNo").ToString + "'," '
        '        tmpsql += "999,'三'"
        '        tmpsql += ")"
        '        EXE_SQL(tmpsql)
        '    Next
        'Next


        'modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-5", My.User.Name) 'azitest
        Dim sScript As String
        If tmpstr <> "" Then
            sScript = "window.open('MMSRptPrint.aspx?ReportPath=/MMS/Invoice');"
            ScriptManager.RegisterStartupScript(Me, Me.GetType(), "PrnOut", sScript, True)

            'Me.rptViewer.Visible = True
            'Try
            '    '-------------------------------------* 報表路徑及認證
            '    Me.rptViewer.ProcessingMode = ProcessingMode.Remote
            '    Me.rptViewer.ServerReport.ReportServerCredentials = Me.ReportCredentials
            '    modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-6", My.User.Name) 'azitest

            '    Me.rptViewer.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))
            '    Me.rptViewer.ServerReport.ReportPath = "/MMS/Invoice"
            '    modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-7", My.User.Name) 'azitest

            '    Me.rptViewer.DataBind()
            '    modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-8", My.User.Name) 'azitest

            '    modDB.InsertSignRecord("MMSInvoiceMaintain", "發票號碼:" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + ",列印完成!!", My.User.Name)
            'Catch ex As Exception
            '    modUtil.showMsg(Me.Page, "錯誤訊息", ex.Message)
            'End Try

            'rptCmp.ProcessingMode = ProcessingMode.Remote
            'rptCmp.ServerReport.ReportServerCredentials = Me.ReportCredentials '* 身分認證 2008/04/03 NEC 杜
            'rptCmp.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))

            'rptCmp.ServerReport.ReportPath = "/MMS/Invoice"
            'rptCmp.ShowParameterPrompts = False
            'rptCmp.DataBind()
            'tmpstr = "發票號碼:" + " \n" + tmpstr + ",列印成功!!"
            'modUtil.showMsg(Me.Page, "訊息", tmpstr)
        Else
            modUtil.showMsg(Me.Page, "訊息", "無可列印發票!!")
        End If
        'Azitest
        'modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-9-1", My.User.Name) 'azitest

        Call ShowGrid()
        'modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-9-2", My.User.Name) 'azitest
        Me.lblMsg.Text = ""
        Session("CHECKED_ITEMS") = Nothing
        'modDB.InsertSignRecord("azitest", "mmsInvMaintain ok-9-3", My.User.Name) 'azitest
    End Sub

    'GET_INVNUM(,年月)
    Private Function GET_INVNUM(ByVal VINVNO As String, ByVal VINVDT As String) As String
        '20150106 應增加超過發票範圍的限制警告
        Dim VINVYM As String
        Dim VINVFW As String
        Dim VINVNUMST, VINVNUMED As String
        Dim UINVNUMST_NEW, UINVNUMED_NEW As String
        Dim INVFLAG As Boolean
        Dim tmpsql As String
        Dim tt, tt2 As DataTable

        '先決定發票年月
        Select Case Mid(VINVDT, 6, 2)
            Case "01", "02"
                VINVYM = "0102"
            Case "03", "04"
                VINVYM = "0304"
            Case "05", "06"
                VINVYM = "0506"
            Case "07", "08"
                VINVYM = "0708"
            Case "09", "10"
                VINVYM = "0910"
            Case "11", "12"
                VINVYM = "1112"
        End Select
        VINVYM = Mid(VINVDT, 1, 4) + VINVYM
        '找出字軌
        tmpsql = "SELECT Nirw,NirsStart,NirsEnd FROM InvoiceNum "
        tmpsql += " WHERE STNID='050101' and FLG_H = '3' AND ((Invtype='') OR (Invtype IS NULL))"
        tmpsql += "   AND USEYEAR='" + Mid(VINVYM, 1, 4) + "'" 'MODIFY ADD IN 20150420
        tmpsql += "   AND MonStart='" + Mid(VINVYM, 5, 2) + "'"
        tmpsql += "   AND MonEnd='" + Mid(VINVYM, 7, 2) + "'"
        '
        tt = get_DataTable(tmpsql)
        INVFLAG = True
        If tt.Rows.Count = 0 Then
            INVFLAG = False
            modUtil.showMsg(Me.Page, "訊息", "查無 :'" + VINVYM + "' 月份的發票使用登錄！")
        Else
            VINVFW = tt.Rows(0).Item("Nirw").ToString '字軌
            VINVNUMST = tt.Rows(0).Item("NirsStart").ToString '起號
            VINVNUMED = tt.Rows(0).Item("NirsEnd").ToString '訖號
        End If

        If INVFLAG Then
            '找出可用的卷號
            tmpsql = "SELECT MIN(INVNUMST) AS INVNUMST FROM MMSInvUsed"
            tmpsql += " WHERE INVYM='" + VINVYM + "'"
            tmpsql += " AND (USENUMED<INVNUMED)" '尚未最後一號
            tmpsql += " AND LINVDT<='" + VINVDT + "' AND ((EINV='') OR (EINV IS NULL))" '同一卷需按日期作排序
            tt = get_DataTable(tmpsql)
            INVFLAG = True
            If tt.Rows(0).Item("INVNUMST").ToString <> "" Then
                '拉出該卷的最終號
                tmpsql = "SELECT INVFW,USENUMED FROM MMSInvUsed"
                tmpsql += " WHERE INVYM='" + VINVYM + "'"
                tmpsql += " AND INVFW='" + VINVFW + "'"
                tmpsql += " AND INVNUMST='" + tt.Rows(0).Item("INVNUMST").ToString.Trim + "' AND ((EINV='') OR (EINV IS NULL))"
                tt2 = get_DataTable(tmpsql)
                UINVNUMST_NEW = Format(Int(tt2.Rows(0).Item("USENUMED")) + 1, "00000000")

                tmpsql = "UPDATE MMSInvUsed SET USENUMED='" + UINVNUMST_NEW + "'"
                tmpsql += ",LINVDT='" + VINVDT + "'"
                tmpsql += " WHERE INVYM='" + VINVYM + "' AND INVFW='" + VINVFW + "'"
                tmpsql += " AND INVNUMST='" + tt.Rows(0).Item("INVNUMST").ToString.Trim + "' AND ((EINV='') OR (EINV IS NULL))"
                EXE_SQL(tmpsql)
            Else  '找不到需開新的一卷
                'INVFLAG = False
                '找出可用的卷號
                tmpsql = "SELECT MAX(INVNUMED) AS INVNUMED FROM MMSInvUsed "
                tmpsql += " WHERE INVYM='" + VINVYM + "' AND INVFW='" + VINVFW + "' AND ((EINV='') OR (EINV IS NULL))"
                tt = get_DataTable(tmpsql)
                If tt.Rows(0).Item("INVNUMED").ToString <> "" Then
                    UINVNUMST_NEW = Format(Int(tt.Rows(0).Item("INVNUMED")) + 1, "00000000")
                Else   '當月尚無紀錄
                    UINVNUMST_NEW = VINVNUMST
                End If
                UINVNUMED_NEW = Format(Int(UINVNUMST_NEW) + 49, "00000000")
                '
                If (UINVNUMED_NEW >= VINVNUMST) And (UINVNUMED_NEW <= VINVNUMED) Then
                    '新增使用卷資料
                    '找出可用的卷號
                    tmpsql = "INSERT INTO MMSInvUsed (INVYM,STNID,INVFW,LINVDT,INVNUMST,INVNUMED,USENUMST,USENUMED,EINV)"
                    tmpsql += " VALUES ('" + VINVYM + "','" + Request.Cookies("STNID").Value + "','" + VINVFW + "','" + VINVDT + "','" + UINVNUMST_NEW + "','" + UINVNUMED_NEW + "'"
                    tmpsql += " ,'" + UINVNUMST_NEW + "','" + UINVNUMST_NEW + "','')"
                    EXE_SQL(tmpsql)
                Else
                    INVFLAG = False
                    modDB.InsertSignRecord("(InvoiceMaintain)_GET_INVNUM", "計算機發票配號已不足:" + UINVNUMED_NEW, My.User.Name)
                    modUtil.showMsg(Me.Page, "訊息", "發票配號已不足:" + UINVNUMED_NEW)
                End If
            End If
            '
            'UINVNUMED_NEW
            '先作發票號碼booking
            'tmpsql = "INSERT INTO INVMH (STNID,INVNO,INVDT,INVTIME,INVNUM,GETDT,GETER,INVTITLE,TAXID,INVAMT,INVTXAMT'"
            'tmpsql += ",INVTAMT,PRNDT,VOID,VOIDDT,VOIDER,MEMOA,MEMOB,MEMOC,MEMOD,MEMOE,CHKNO,CUSTID)'"
            'tmpsql += " VALUES "
            'tmpsql += "('" + Request.Cookies("STNID").Value + "','" + VINVNO + "','" + VINVDT + "','','" + VINVFW + UINVNUMST_NEW + "','" + Now.ToString("yyyy/MM/dd HH:mm:ss") + "'"
            'tmpsql += " ,'GETERID','','',0,0,0,'','','','','','','','','','','')"

            'EXE_SQL(tmpsql)
            If INVFLAG Then
                GET_INVNUM = VINVFW + UINVNUMST_NEW
            Else
                GET_INVNUM = "ZZ99999999"
            End If
        Else
            GET_INVNUM = "ZZ99999999"
        End If
    End Function

    '20150925 電子發票取號
    Private Function GET_eINVNUM(ByVal VINVNO As String, ByVal VINVDT As String) As String
        '20150106 應增加超過發票範圍的限制警告
        Dim VINVYM As String
        Dim VINVFW As String
        Dim VINVNUMST, VINVNUMED As String
        Dim UINVNUMST_NEW, UINVNUMED_NEW As String
        Dim INVFLAG As Boolean
        Dim tmpsql As String
        Dim tt, tt2 As DataTable

        '先決定發票年月
        Select Case Mid(VINVDT, 6, 2)
            Case "01", "02"
                VINVYM = "0102"
            Case "03", "04"
                VINVYM = "0304"
            Case "05", "06"
                VINVYM = "0506"
            Case "07", "08"
                VINVYM = "0708"
            Case "09", "10"
                VINVYM = "0910"
            Case "11", "12"
                VINVYM = "1112"
        End Select
        VINVYM = Mid(VINVDT, 1, 4) + VINVYM
        '找出字軌
        tmpsql = "SELECT Nirw,NirsStart,NirsEnd FROM InvoiceNum "
        tmpsql += " WHERE STNID='050101' and FLG_H = '3' AND Invtype='35'"
        tmpsql += "   AND USEYEAR='" + Mid(VINVYM, 1, 4) + "'" 'MODIFY ADD IN 20150420
        tmpsql += "   AND MonStart='" + Mid(VINVYM, 5, 2) + "'"
        tmpsql += "   AND MonEnd='" + Mid(VINVYM, 7, 2) + "'"
        '
        tt = get_DataTable(tmpsql)
        INVFLAG = True
        If tt.Rows.Count = 0 Then
            INVFLAG = False
            modUtil.showMsg(Me.Page, "訊息", "查無 :'" + VINVYM + "' 月份的發票使用登錄！")
        Else
            VINVFW = tt.Rows(0).Item("Nirw").ToString '字軌
            VINVNUMST = tt.Rows(0).Item("NirsStart").ToString '起號
            VINVNUMED = tt.Rows(0).Item("NirsEnd").ToString '訖號
        End If

        If INVFLAG Then
            '找出可用的卷號
            tmpsql = "SELECT MIN(INVNUMST) AS INVNUMST FROM MMSInvUsed"
            tmpsql += " WHERE INVYM='" + VINVYM + "'"
            tmpsql += " AND (USENUMED<INVNUMED)" '尚未最後一號
            tmpsql += " AND LINVDT<='" + VINVDT + "'" '同一卷需按日期作排序
            tmpsql += " AND EINV='Y'"
            tt = get_DataTable(tmpsql)
            INVFLAG = True
            If tt.Rows(0).Item("INVNUMST").ToString <> "" Then
                '拉出該卷的最終號
                tmpsql = "SELECT INVFW,USENUMED FROM MMSInvUsed"
                tmpsql += " WHERE INVYM='" + VINVYM + "'"
                tmpsql += " AND INVFW='" + VINVFW + "'"
                tmpsql += " AND INVNUMST='" + tt.Rows(0).Item("INVNUMST").ToString.Trim + "' AND EINV='Y'"
                tt2 = get_DataTable(tmpsql)
                UINVNUMST_NEW = Format(Int(tt2.Rows(0).Item("USENUMED")) + 1, "00000000")

                tmpsql = "UPDATE MMSInvUsed SET USENUMED='" + UINVNUMST_NEW + "'"
                tmpsql += ",LINVDT='" + VINVDT + "'"
                tmpsql += " WHERE INVYM='" + VINVYM + "' AND INVFW='" + VINVFW + "'"
                tmpsql += " AND INVNUMST='" + tt.Rows(0).Item("INVNUMST").ToString.Trim + "' AND EINV='Y'"
                EXE_SQL(tmpsql)
            Else  '找不到需開新的一卷
                'INVFLAG = False
                '找出可用的卷號
                tmpsql = "SELECT MAX(INVNUMED) AS INVNUMED FROM MMSInvUsed "
                tmpsql += " WHERE INVYM='" + VINVYM + "' AND INVFW='" + VINVFW + "' AND EINV='Y'"
                tt = get_DataTable(tmpsql)
                If tt.Rows(0).Item("INVNUMED").ToString <> "" Then
                    UINVNUMST_NEW = Format(Int(tt.Rows(0).Item("INVNUMED")) + 1, "00000000")
                Else   '當月尚無紀錄
                    UINVNUMST_NEW = VINVNUMST
                End If
                UINVNUMED_NEW = Format(Int(UINVNUMST_NEW) + 49, "00000000")
                '
                If (UINVNUMED_NEW >= VINVNUMST) And (UINVNUMED_NEW <= VINVNUMED) Then
                    '新增使用卷資料
                    '找出可用的卷號
                    tmpsql = "INSERT INTO MMSInvUsed (INVYM,STNID,INVFW,LINVDT,INVNUMST,INVNUMED,USENUMST,USENUMED,EINV)"
                    tmpsql += " VALUES ('" + VINVYM + "','" + Request.Cookies("STNID").Value + "','" + VINVFW + "','" + VINVDT + "','" + UINVNUMST_NEW + "','" + UINVNUMED_NEW + "'"
                    tmpsql += " ,'" + UINVNUMST_NEW + "','" + UINVNUMST_NEW + "','Y')"
                    EXE_SQL(tmpsql)
                Else
                    INVFLAG = False
                    modDB.InsertSignRecord("(InvoiceMaintain)_GET_eINVNUM", "電子發票配號已不足:" + UINVNUMED_NEW, My.User.Name)
                    modUtil.showMsg(Me.Page, "訊息", "發票配號已不足:" + UINVNUMED_NEW)
                End If
            End If
            '
            If INVFLAG Then
                GET_eINVNUM = VINVFW + UINVNUMST_NEW
            Else
                GET_eINVNUM = "ZZ99999999"
            End If
        Else
            GET_eINVNUM = "ZZ99999999"
        End If
    End Function

    Protected Sub ClearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearButton.Click
        txtDtFrom.Text = ""
        txtDtTo.Text = ""
        txtDtFrom0.Text = ""
        txtDtTo0.Text = ""
        RadioButtonList1.SelectedIndex = 0
        'CheckBoxList1.Enabled = False
        'CheckBoxList1.Items(0).Selected = False
        'CheckBoxList1.Items(1).Selected = False
        'CheckBoxList1.Items(2).Selected = False
        Me.GridView1.Visible = False
        Me.GridView3.Visible = False
        Me.GridView4.Visible = False

    End Sub

    Protected Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        '取消審核
        '20150106 應增加如以有發票號碼則不可取消審核 
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
        Dim docnotmp As String
        tt = Session("BatchGrid1")
        For i = 0 To categoryIDList.Count - 1
            If tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") = "" Then
                docnotmp = get_DataTable("select [ApplyDocNo] from MMSInvoiceM where sn='" + tt.Rows(categoryIDList.Item(i)).Item("snM").ToString + "'").Rows(0).Item(0)
                tmpsql = "update [MMSInvoiceApplyM] set [Status]='N',[ReviewDate]='',[Reviewers]=''"
                tmpsql += " where [DocNo]='" + docnotmp + "'"
                EXE_SQL(tmpsql)
                'modify IN 20160129
                tmpsql = "update [MMSInvoiceApplyD] set [EFFECTIVE]='N'"
                tmpsql += " where [DocNo]='" + docnotmp + "'"
                EXE_SQL(tmpsql)
                tmpsql = "delete from MMSInvoiceM "
                tmpsql += " where sn='" + tt.Rows(categoryIDList.Item(i)).Item("snM").ToString + "'"
                EXE_SQL(tmpsql)
                tmpsql = "delete from MMSInvoiceD "
                tmpsql += " where snM='" + tt.Rows(categoryIDList.Item(i)).Item("snM").ToString + "'"
                EXE_SQL(tmpsql)
                tmpstr = " "
            End If
        Next
        '
        If tmpstr <> "" Then
            tmpstr = "取消審核成功!!"
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
        Else
            modUtil.showMsg(Me.Page, "訊息", "無可取消審核!!")
        End If
        '
        Call ShowGrid()
        Me.lblMsg.Text = ""
        Session("CHECKED_ITEMS") = Nothing        
    End Sub


    '原給號程式
    ''GET_INVNUM(,年月)
    'Private Function GET_INVNUM(ByVal VINVNO As String, ByVal VINVDT As String) As String
    '    Dim VINVYM As String
    '    Dim VINVFW As String
    '    Dim VINVNUMST, VINVNUMED As String
    '    Dim UINVNUMST_NEW, UINVNUMED_NEW As String
    '    Dim INVFLAG As Boolean
    '    Dim tmpsql As String
    '    Dim tt, tt2 As DataTable

    '    '先決定發票年月
    '    Select Case Mid(VINVDT, 6, 2)
    '        Case "01", "02"
    '            VINVYM = "0102"
    '        Case "03", "04"
    '            VINVYM = "0304"
    '        Case "05", "06"
    '            VINVYM = "0506"
    '        Case "07", "08"
    '            VINVYM = "0708"
    '        Case "09", "10"
    '            VINVYM = "0910"
    '        Case "11", "12"
    '            VINVYM = "1112"
    '    End Select
    '    VINVYM = Mid(VINVDT, 1, 4) + VINVYM
    '    '找出字軌
    '    tmpsql = "SELECT INVFW,INVNUMST,INVNUMED FROM InvoiceM"
    '    tmpsql += " WHERE INVYM='" + VINVYM + "'"
    '    tmpsql += " AND STNID='" + Request.Cookies("STNID").Value + "' "
    '    tt = get_DataTable(tmpsql)
    '    INVFLAG = True
    '    If tt.Rows.Count = 0 Then
    '        INVFLAG = False
    '        modUtil.showMsg(Me.Page, "訊息", "查無 :'" + VINVYM + "' 月份的發票使用登錄！")
    '    Else
    '        VINVFW = tt.Rows(0).Item("INVFW").ToString
    '        VINVNUMST = tt.Rows(0).Item("INVNUMST").ToString
    '        VINVNUMED = tt.Rows(0).Item("INVNUMED").ToString
    '    End If

    '    If INVFLAG Then
    '        '找出可用的卷號
    '        tmpsql = "SELECT MIN(INVNUMST) AS INVNUMST FROM MMSInvUsed"
    '        tmpsql += " WHERE INVYM='" + VINVYM + "'"
    '        tmpsql += " AND (USENUMED<INVNUMED)" '尚未最後一號
    '        tmpsql += " AND LINVDT<='" + VINVDT + "'" '同一卷需按日期作排序
    '        tt = get_DataTable(tmpsql)
    '        INVFLAG = True
    '        If tt.Rows(0).Item("INVNUMST").ToString <> "" Then
    '            '拉出該卷的最終號
    '            tmpsql = "SELECT INVFW,USENUMED FROM MMSInvUsed"
    '            tmpsql += " WHERE INVYM='" + VINVYM + "'"
    '            tmpsql += " AND INVFW='" + VINVFW + "'"
    '            tmpsql += " AND INVNUMST='" + tt.Rows(0).Item("INVNUMST").ToString.Trim + "'"
    '            tt2 = get_DataTable(tmpsql)
    '            UINVNUMST_NEW = Format(Int(tt2.Rows(0).Item("USENUMED")) + 1, "00000000")

    '            tmpsql = "UPDATE MMSInvUsed SET USENUMED='" + UINVNUMST_NEW + "'"
    '            tmpsql += ",LINVDT='" + VINVDT + "'"
    '            tmpsql += " WHERE INVYM='" + VINVYM + "' AND INVFW='" + VINVFW + "'"
    '            tmpsql += " AND INVNUMST='" + tt.Rows(0).Item("INVNUMST").ToString.Trim + "'"
    '            EXE_SQL(tmpsql)
    '        Else  '找不到需開新的一卷
    '            INVFLAG = False
    '            '找出可用的卷號
    '            tmpsql = "SELECT MAX(INVNUMED) AS INVNUMED FROM MMSInvUsed"
    '            tmpsql += " WHERE INVYM='" + VINVYM + "' AND INVFW='" + VINVFW + "'"
    '            tt = get_DataTable(tmpsql)
    '            If tt.Rows(0).Item("INVNUMED").ToString <> "" Then
    '                UINVNUMST_NEW = Format(Int(tt.Rows(0).Item("INVNUMED")) + 1, "00000000")
    '            Else   '當月尚無紀錄
    '                UINVNUMST_NEW = VINVNUMST
    '            End If
    '            UINVNUMED_NEW = Format(Int(UINVNUMST_NEW) + 49, "00000000")
    '            '新增使用卷資料
    '            '找出可用的卷號
    '            tmpsql = "INSERT INTO MMSInvUsed (INVYM,STNID,INVFW,LINVDT,INVNUMST,INVNUMED,USENUMST,USENUMED)"
    '            tmpsql += " VALUES('" + VINVYM + "','" + Request.Cookies("STNID").Value + "','" + VINVFW + "','" + VINVDT + "','" + UINVNUMST_NEW + "','" + UINVNUMED_NEW + "'"
    '            tmpsql += " ,'" + UINVNUMST_NEW + "','" + UINVNUMST_NEW + "')"
    '            EXE_SQL(tmpsql)
    '        End If
    '        'UINVNUMED_NEW
    '        '先作發票號碼booking
    '        tmpsql = "INSERT INTO INVMH (STNID,INVNO,INVDT,INVTIME,INVNUM,GETDT,GETER,INVTITLE,TAXID,INVAMT,INVTXAMT'"
    '        tmpsql += ",INVTAMT,PRNDT,VOID,VOIDDT,VOIDER,MEMOA,MEMOB,MEMOC,MEMOD,MEMOE,CHKNO,CUSTID)'"
    '        tmpsql += " VALUES "
    '        tmpsql += "('" + Request.Cookies("STNID").Value + "','" + VINVNO + "','" + VINVDT + "','','" + VINVFW + UINVNUMST_NEW + "','" + Now.ToString("yyyy/MM/dd HH:mm:ss") + "'"
    '        tmpsql += " ,'GETERID','','',0,0,0,'','','','','','','','','','','')"

    '        EXE_SQL(tmpsql)
    '        GET_INVNUM = VINVFW + UINVNUMST_NEW
    '    Else
    '        GET_INVNUM = "ZZ99999999"
    '    End If
    'End Function

    '發票給號

    '更新合約最後發票開立期別
    Private Sub UpdateLastInvoiceDate(ByVal cno As String)
        Dim tt As DataTable
        Dim LASTINVOICEPERIOD, tmpsql As String
        tmpsql = "Select ISNULL(MAX(INVOICEPERIOD),'') AS INVOICEPERIOD FROM MMSinvoiceA where ContractNo='" + cno + "'"
        tt = get_DataTable(tmpsql)
        LASTINVOICEPERIOD = tt.Rows(0).Item("INVOICEPERIOD").ToString
        '
        tmpsql = "Update MMSContract set LastInvoiceDate='" + LASTINVOICEPERIOD + "' where ContractNo='" + cno + "'"
        EXE_SQL(tmpsql)
    End Sub

    Protected Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        RememberOldValues()
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何資料!!")
            Exit Sub
        End If
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i, j As Integer
        Dim tt, tta As DataTable
        Dim tmpstr As String = ""
        Dim tmpInvNO As String = ""
        Dim applydocno As String
        Dim VEINV As String = ""
        tt = Session("BatchGrid1")
        For i = 0 To categoryIDList.Count - 1
            If tt.Rows(categoryIDList.Item(i)).Item("InvoiceTypeM") = "收銀發票" And tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") = "" And tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate") <> "" And tt.Rows(categoryIDList.Item(i)).Item("StatusN") <> "作廢" Then
                If changeDate(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate")) <= "2015/09/30" Then
                    tmpInvNO = GET_INVNUM("", changeDate(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate")))
                Else
                    VEINV = "Y"
                    tmpInvNO = GET_eINVNUM("", changeDate(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate")))
                End If
                'tmpInvNO = GET_eINVNUM("", changeDate(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate")))
                If tmpInvNO = "ZZ99999999" Then GoTo aa
                tmpstr += tmpInvNO + " \n"
                tmpsql = "update MMSInvoiceM set InvoiceNo='" + tmpInvNO + "',EINV='" + VEINV + "' where sn=" + tt.Rows(categoryIDList.Item(i)).Item("snM").ToString
                EXE_SQL(tmpsql)
                tmpsql = "update MMSInvoiceD set InvoiceNo='" + tmpInvNO + "' where snM=" + tt.Rows(categoryIDList.Item(i)).Item("snM").ToString
                EXE_SQL(tmpsql)
                'tmpsql = "insert into [MMSInvoicePrint]([InvoiceNo],[InvoiceDate],[PrintDateTime])"
                'tmpsql += " values('" + tmpInvNO + "','" + changeDate(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate")) + "',GETDATE())"
                'EXE_SQL(tmpsql)
                tmpsql = "select ApplyDocNo from MMSInvoiceM where InvoiceNo='" + tmpInvNO + "'"
                applydocno = get_DataTable(tmpsql).Rows(0).Item(0)
                tmpsql = "update MMSInvoiceApplyD set InvoiceNo='" + tmpInvNO + "',Effective='Y' where DocNo='" + applydocno + "'"
                EXE_SQL(tmpsql)
                'AZITEMP ADD IN 20160107
                tmpsql = "Select CONTRACTNO,INVOICEPERIOD From MMSInvoiceApplyD Where DocNo='" + applydocno + "' AND CONTRACTNO<>''"
                tta = get_DataTable(tmpsql)
                If tta.Rows.Count > 0 Then
                    j = 0
                    '
                    For j = 0 To (tta.Rows.Count - 1)
                        tmpsql = "UPDATE MMSINVOICEA SET INVOICEDATE='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate") + "',INVOICENO='" + tmpInvNO + "'"
                        tmpsql +=" Where CONTRACTNO='" + tta.Rows(j).Item("CONTRACTNO") + "' AND INVOICEPERIOD='" + tta.Rows(j).Item("INVOICEPERIOD") + "'" 
                        EXE_SQL(tmpsql)
                    Next
                End If
                tmpstr += tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + " \n"
            End If

            If tt.Rows(categoryIDList.Item(i)).Item("InvoiceTypeM") = "收銀發票" And tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") <> "" And tt.Rows(categoryIDList.Item(i)).Item("StatusN") <> "作廢" Then
                'tmpsql = "insert into [MMSInvoicePrint]([InvoiceNo],[InvoiceDate],[PrintDateTime])"
                'tmpsql += " values('" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "','" + changeDate(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate")) + "',GETDATE())"
                ''tmpsql += " values('" + tmpInvNO + "','" + changeDate(tt.Rows(categoryIDList.Item(i)).Item("InvoiceDate")) + "',GETDATE())"
                'EXE_SQL(tmpsql)
                tmpstr += tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + " \n"
            End If
aa:     Next 'aa:

        If tmpstr <> "" Then
            tmpstr = "發票號碼:" + " \n" + tmpstr + ",給號成功!!2015/10/01 之後的發票一律配號電子發票。麻煩請詳細確認。"
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
        Else
            modUtil.showMsg(Me.Page, "訊息", "無可給號之發票開立申請!!")
        End If
        Call ShowGrid()
        Me.lblMsg.Text = ""
        Session("CHECKED_ITEMS") = Nothing
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

    '核准作廢
    Protected Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim VCONTRACTNO, VINVYM As String

        RememberOldValues()
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何資料!!")
            Exit Sub
        End If
        '
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i, j As Integer
        Dim tt, tt1 As DataTable
        Dim tmpstr As String = ""
        tt = Session("BatchGrid1")
        For i = 0 To categoryIDList.Count - 1
            If tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") <> "" Then
                modDB.InsertSignRecord("(InvoiceMaintain)_核准發票作廢", "InvoiceNo =" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo"), My.User.Name)
                '
                tmpsql = "update MMSInvoiceM set Status = 'R'"
                tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                EXE_SQL(tmpsql)
                tmpsql = "update MMSInvoiceApplyD set Effective = 'R'"
                tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                EXE_SQL(tmpsql)
                '20140609 add by Azi 增加 update 發票開立期別處理
                VCONTRACTNO = ""
                VINVYM = ""
                tmpsql = "SELECT ISNULL(CONTRACTNO,'') AS CONTRACTNO,ISNULL(INVOICEPERIOD,'') AS INVOICEPERIOD FROM MMSInvoiceApplyD"
                tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                tmpsql += "   and ContractNo<>''" 'modify in 20140925
                'LOGSQL(tmpsql)
                tt1 = get_DataTable(tmpsql)
                'modDB.InsertSignRecord("AziTest", "InvoiceMaintain tt.Rows.Count=" + tt1.Rows.Count.ToString, My.User.Name) 'AZITEST
                If tt1.Rows.Count > 0 Then
                    'modDB.InsertSignRecord("azitest", "作廢合約期別:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " 筆數 = " & tt1.Rows.Count.ToString, My.User.Name)
                    j = 0
                    '20151223 一張發票可能有多筆保養合約的資料，增加 For loop處理
                    For j = 0 To (tt1.Rows.Count - 1)
                        'modDB.InsertSignRecord("azitest", " j = " & j.ToString, My.User.Name)
                        VCONTRACTNO = tt1.Rows(j).Item("CONTRACTNO")
                        VINVYM = tt1.Rows(j).Item("INVOICEPERIOD")
                        modDB.InsertSignRecord("MMSInsMainTain", "核准作廢合約期別:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " CONTRACTNO =" & VCONTRACTNO & " INVYM =" & VINVYM, My.User.Name)
                        If (VCONTRACTNO.Trim <> "") And (VINVYM.Trim <> "") Then
                            modDB.InsertSignRecord("(InvoiceMaintain)_發票作廢", "期別處理:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " CONTRACTNO =" & VCONTRACTNO & " INVYM =" & VINVYM, My.User.Name)
                            tmpsql = "update MMSInvoiceA set YN = 'N',INVOICEDATE='',INVOICENO=''" 'MODIFY IN 20160107
                            tmpsql += " where CONTRACTNO='" + VCONTRACTNO + "' AND INVOICEPERIOD='" + VINVYM + "'"
                            '
                            'LOGSQL(tmpsql) 'azitest
                            EXE_SQL(tmpsql)
                            ''增加UPDATE合約最後開立期別_20150901 
                            'UpdateLastInvoiceDate(VCONTRACTNO)
                        End If
                    Next
                End If
                tmpstr += tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + " \n"
            End If
        Next
        If tmpstr <> "" Then
            tmpstr = "發票號碼:" + " \n" + tmpstr + "核准作廢成功!!"
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
        Else
            modUtil.showMsg(Me.Page, "訊息", "無可核准作廢發票!!")
        End If
        '
        Call ShowGrid()
        Me.lblMsg.Text = ""
        Session("CHECKED_ITEMS") = Nothing
    End Sub

    '取消作廢申請
    Protected Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        RememberOldValues()
        If Session("CHECKED_ITEMS") Is Nothing Then
            modUtil.showMsg(Me.Page, "訊息", "尚未選擇任何資料!!")
            Exit Sub
        End If
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        Dim tmpsql As String
        Dim i, j As Integer
        Dim tt, tt1 As DataTable
        Dim tmpstr As String = ""
        tt = Session("BatchGrid1")
        For i = 0 To categoryIDList.Count - 1
            If tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") <> "" Then
                modDB.InsertSignRecord("(InvoiceMaintain)_取消發票(作廢申請)", "InvoiceNo =" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo"), My.User.Name)
                '
                tmpsql = "update MMSInvoiceM set Status = 'Y'"
                tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                EXE_SQL(tmpsql)
                tmpsql = "update MMSInvoiceApplyD set Effective = 'Y'"
                tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                EXE_SQL(tmpsql)
                ''20140609 add by Azi 增加 update 發票開立期別處理
                'VCONTRACTNO = ""
                'VINVYM = ""
                'tmpsql = "SELECT ISNULL(CONTRACTNO,'') AS CONTRACTNO,ISNULL(INVOICEPERIOD,'') AS INVOICEPERIOD FROM MMSInvoiceApplyD"
                'tmpsql += " where InvoiceNo='" + tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + "'"
                'tmpsql += "   and ContractNo<>''" 'modify in 20140925
                ''LOGSQL(tmpsql)
                'tt1 = get_DataTable(tmpsql)
                ''modDB.InsertSignRecord("AziTest", "InvoiceMaintain tt.Rows.Count=" + tt1.Rows.Count.ToString, My.User.Name) 'AZITEST
                'If tt1.Rows.Count > 0 Then
                '    'modDB.InsertSignRecord("azitest", "作廢合約期別:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " 筆數 = " & tt1.Rows.Count.ToString, My.User.Name)
                '    j = 0
                '    '20151223 一張發票可能有多筆保養合約的資料，增加 For loop處理
                '    For j = 0 To (tt1.Rows.Count - 1)
                '        'modDB.InsertSignRecord("azitest", " j = " & j.ToString, My.User.Name)
                '        VCONTRACTNO = tt1.Rows(j).Item("CONTRACTNO")
                '        VINVYM = tt1.Rows(j).Item("INVOICEPERIOD")
                '        modDB.InsertSignRecord("MMSInsMainTain", "核准作廢合約期別:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " CONTRACTNO =" & VCONTRACTNO & " INVYM =" & VINVYM, My.User.Name)
                '        If (VCONTRACTNO.Trim <> "") And (VINVYM.Trim <> "") Then
                '            modDB.InsertSignRecord("(InvoiceMaintain)_發票作廢", "期別處理:InvoiceNo=" & tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") & " CONTRACTNO =" & VCONTRACTNO & " INVYM =" & VINVYM, My.User.Name)
                '            tmpsql = "update MMSInvoiceA set YN = 'N',INVOICEDATE='',INVOICENO=''" 'MODIFY IN 20160107
                '            tmpsql += " where CONTRACTNO='" + VCONTRACTNO + "' AND INVOICEPERIOD='" + VINVYM + "'"
                '            '
                '            'LOGSQL(tmpsql) 'azitest
                '            EXE_SQL(tmpsql)
                '            ''增加UPDATE合約最後開立期別_20150901 
                '            'UpdateLastInvoiceDate(VCONTRACTNO)
                '        End If
                '    Next
                'End If
                tmpstr += tt.Rows(categoryIDList.Item(i)).Item("InvoiceNo") + " \n"
            End If
        Next
        If tmpstr <> "" Then
            tmpstr = "發票號碼:" + " \n" + tmpstr + "取消(作廢申請)成功!!"
            modUtil.showMsg(Me.Page, "訊息", tmpstr)
        Else
            modUtil.showMsg(Me.Page, "訊息", "無可取消(作廢申請)發票!!")
        End If
        '
        Call ShowGrid()
        Me.lblMsg.Text = ""
        Session("CHECKED_ITEMS") = Nothing
    End Sub
End Class
