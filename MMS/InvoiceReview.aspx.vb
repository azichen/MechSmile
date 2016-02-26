Imports System.Data
Partial Class MMS_InvoiceReview
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
                If ViewState("ROL").Substring(2, 1) = "N" Then Me.Button3.Visible = False
            End If

            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            'modDB.SetGridViewStyle(Me.GridView1)          '* 套用gridview樣式
            'DispGridE()
            '--------------------------------------------------* 加入連結欄位

            Me.lblMsg.Text = Request("msg")
            Me.Button3.Enabled = False
            txtDtFrom.Focus()

            Me.CustFrom.Attributes.Add("onkeypress", "return ToUpper()")
            Me.CustTo.Attributes.Add("onkeypress", "return ToUpper()")
        End If

        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        'Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        'Me.item_defList.SelectCommand = Me.txtSql.Text
        '--------------------------------------------------* 設定GridView資料來源ID
        modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
        modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
        modUtil.SetDateObj(Me.txtDtFrom0, False, Me.txtDtTo0, False)
        modUtil.SetDateObj(Me.txtDtTo0, False, Nothing, False)
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
                Me.Button3.Enabled = True
            Else
                Me.Button3.Enabled = False
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
    Protected Sub btnQry_Click(sender As Object, e As System.EventArgs) Handles btnQry.Click
        Call ShowGrid()
        Me.lblMsg.Text = ""
        Session("CHECKED_ITEMS") = Nothing
        If Me.GridView1.Rows.Count > 0 Then
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
        Else
            Me.Button1.Enabled = False
            Me.Button2.Enabled = False
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


    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        Dim sSql As String
        sSql = "select "
        sSql += "0 as checkf,"
        sSql += " 0 as sn,"
        sSql += " a.DocNo,"
        sSql += " a.CustomerNo,"
        sSql += " '' as ContractNo,"
        sSql += " a.GUINumber,"
        sSql += " a.TitleOfInvoice,"
        sSql += " (select top 1 b.InvoiceDate from MMSInvoiceApplyD b where b.DocNo=a.DocNo) as InvoiceDate,"
        sSql += " (select top 1 b.InvoiceNo from MMSInvoiceApplyD b where b.DocNo=a.DocNo) as InvoiceNo,"
        sSql += " case a.InvoiceType when '1' then '收銀發票'  when '2' then  '手開發票'  end  as InvoiceType, "
        sSql += "REPLACE(CONVERT(varchar(20), (CAST((select sum(c.amount) from MMSInvoiceApplyD c where c.DocNo=a.DocNo) AS money)), 1), '.00', '') as amt"
        sSql += " from MMSInvoiceApplyM a"
        sSql += " left outer join MMSCustomers c on a.CustomerNo=c.CustomerNo"
        sSql += " where a.ReviewDate =''"
        If Me.txtDtFrom.Text.Trim <> "" Or Me.txtDtTo.Text.Trim <> "" Then
            sSql += " and a.ApplyDate >= '" + Me.txtDtFrom.Text.Trim + "'"
            sSql += " and a.ApplyDate <= '" + Me.txtDtTo.Text.Trim + "'"
        End If
        If Me.txtDtFrom0.Text.Trim <> "" Or Me.txtDtTo0.Text.Trim <> "" Then
            sSql += "and a.DocNo in (select c.DocNo from MMSInvoiceApplyD c where c.InvoiceDate >='" + Me.txtDtFrom0.Text.Trim + "' and c.InvoiceDate <= '" + Me.txtDtTo0.Text.Trim + "')"
        End If
        If Me.CustFrom.Text.Trim <> "" Then
            sSql += " and a.CustomerNo >= '" + Me.CustFrom.Text.Trim + "' "
        End If
        If Me.CustTo.Text.Trim <> "" Then
            sSql += " and a.CustomerNo <= '" + Me.CustTo.Text.Trim + "' "
        End If
        If GetAreaCode() <> "" Then
            sSql += " and c.AreaCode ='" + GetAreaCode() + "'"
        End If
        sSql += " Order by InvoiceDate"
        Dim tt As DataTable = get_DataTable(sSql)
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            tt.Rows(i).Item("sn") = i
        Next
        Session("BatchGrid1") = tt
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
        If tt.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
            Exit Sub
        End If
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
        End Select

    End Sub

    Protected Sub GridView1_RowEditing(sender As Object, e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Me.GridView1.EditIndex = e.NewEditIndex
        'Dim row = GridView1.Rows(Me.GridView1.EditIndex)
        'modUtil.SetDateObj((CType((row.Cells(9).Controls(0)), TextBox)), False, Nothing, False)
        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowUpdating(sender As Object, e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating

        'Retrieve the table from the session object.
        Dim dt As DataTable = CType(Session("BatchGrid1"), DataTable)

        'Update the values.
        Dim row = GridView1.Rows(Me.GridView1.EditIndex)
        Dim ss As String
        If (CType((row.Cells(5).Controls(0)), TextBox)).Text.Trim <> "" Then
            Try
                Dim dd As Date = (CType((row.Cells(5).Controls(0)), TextBox)).Text.Trim
            Catch ex As Exception
                modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & "日期格式錯誤!!")
                Exit Sub
            End Try
        End If
        dt.Rows(row.DataItemIndex)("InvoiceDate") = (CType((row.Cells(5).Controls(0)), TextBox)).Text
        dt.Rows(row.DataItemIndex)("InvoiceNo") = (CType((row.Cells(6).Controls(0)), TextBox)).Text.ToUpper

        'Reset the edit index.
        GridView1.EditIndex = -1

        'Bind data to the GridView control.

        Me.GridView1.DataSource = Session("BatchGrid1")
        Me.GridView1.DataBind()
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Session("QryField") = Me.ViewState("QryField")
        Dim sSql As String
        sSql = "select distinct "
        sSql += "0 as checkf,"
        sSql += "0 as sn,"
        sSql += "a.DocNo,"
        'sSql += "a.Serial,"
        sSql += "b.CustomerNo,"
        sSql += "a.ContractNo,"
        sSql += "b.GUINumber,"
        sSql += "b.TitleOfInvoice,"
        sSql += "a.ItemName,"
        sSql += "REPLACE(CONVERT(varchar(20), (CAST(a.UnitPrice AS money)), 1), '.00', '') as UnitPrice,"
        sSql += "a.Quantity,"
        sSql += "REPLACE(CONVERT(varchar(20), (CAST(a.Amount AS money)), 1), '.00', '') as Amount,"
        sSql += "a.Memo,"
        sSql += "a.InvoiceDate,"
        sSql += "a.InvoiceNo"
        sSql += " from dbo.MMSInvoiceApplyD a"
        sSql += " left outer join dbo.MMSInvoiceApplyM b on a.DocNo=b.DocNo"
        sSql += " where b.DocNo ='" + CType(Me.GridView1.SelectedRow.FindControl("Label5"), Label).Text + "'"
        Dim tt As DataTable = get_DataTable(sSql)
        Dim i As Integer
        Session("BatchGrid2") = tt
        Me.GridView2.DataSource = Session("BatchGrid2")
        Me.GridView2.DataBind()
    End Sub
#End Region


    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
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

    Protected Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Session("CHECKED_ITEMS") = Nothing
        RePopulateValues()
        Me.Button3.Enabled = False
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

        For i = 0 To categoryIDList.Count - 1
            If checkDoc(categoryIDList.Item(i)) = False Then
                tmpstr += getDocNo(categoryIDList.Item(i)) + " \n"
            End If
        Next
        If tmpstr.Trim.Length <> 0 Then
            modUtil.showMsg(Me.Page, "訊息", "有下列申請單號有資料沒選到!! \n" + tmpstr)
            Exit Sub
        End If

        For i = 0 To categoryIDList.Count - 1
            If checkDocDate(categoryIDList.Item(i)) = False Then
                tmpstr += getDocNo(categoryIDList.Item(i)) + " \n"
            End If
        Next
        If tmpstr.Trim.Length <> 0 Then
            modUtil.showMsg(Me.Page, "訊息", "有下列申請單號有發票日期沒輸入!! \n" + tmpstr)
            Exit Sub
        End If

        Dim docnoF As String = ""
        Dim docno As String
        For i = 0 To categoryIDList.Count - 1
            tmpsql = "update MMSInvoiceApplyM set ReviewDate='" + GET_FW_DATE() + "',Reviewers='" + My.User.Name + "'"
            tmpsql += ",Status='Y'"
            tmpsql += " where DocNo='" + getDocNo(categoryIDList.Item(i)) + "'"
            EXE_SQL(tmpsql)
            docno = getDocNo(categoryIDList.Item(i))
            UpdateRow(categoryIDList.Item(i), docno)
            If docnoF <> docno Then
                InsertInvoice(docno)
            End If
            docnoF = docno
        Next
        Response.Redirect("InvoiceReview.aspx?Returnflag=1&msg=申請審核成功!!")
    End Sub

    Private Function checkDocDate(ByVal sn As String) As Boolean
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
            If tt.Rows(i).Item("DocNo") = dd Then
                If IsDBNull(tt.Rows(i).Item("InvoiceDate")) Then
                    checkDocDate = False
                    Exit Function
                Else
                    If tt.Rows(i).Item("InvoiceDate") = "" Then
                        checkDocDate = False
                        Exit Function
                    End If
                End If
            End If
        Next
        checkDocDate = True
    End Function

    Private Sub UpdateRow(ByVal sn As String, ByVal docno As String)
        Dim tt As DataTable
        Dim tmpsql As String
        Dim effect As String
        tt = Session("BatchGrid1")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("sn") = sn Then
                'add in 20140623
                If (tt.Rows(i).Item("InvoiceNo") <> "") And (tt.Rows(i).Item("InvoiceDate") <> "") Then
                    effect = "Y"
                Else
                    effect = "N"
                End If
                tmpsql = "update MMSInvoiceApplyD set "
                tmpsql += " InvoiceDate='" + tt.Rows(i).Item("InvoiceDate") + "',"
                tmpsql += " InvoiceNo='" + tt.Rows(i).Item("InvoiceNo") + "'"
                tmpsql += " where DocNo='" + tt.Rows(i).Item("DocNo") + "'" ' and Serial=" + tt.Rows(i).Item("Serial").ToString
                EXE_SQL(tmpsql)
                Exit For
            End If
        Next
    End Sub

    Private Sub InsertInvoice(ByVal sn As String)
        Dim ttm, ttd As DataTable ', ttq
        Dim tmpsql As String
        Dim ti As Integer
        Dim CUSDAYALLOW As Integer = 0

        'tmpsql = "select * from MMSInvoiceApplyM where DocNo='" + sn + "'"
        tmpsql = "select A.CustomerNo,A.GUINumber,A.TitleOfInvoice,B.GUINumber AS APPLYGUIN,B.TitleOfInvoice AS APPLYTITLE  from MMSCustomers A inner join MMSInvoiceApplyM B on A.CustomerNo = B.CustomerNo where B.DocNo='" + sn + "'"
        ttm = get_DataTable(tmpsql)
        tmpsql = "select * from MMSInvoiceApplyD where DocNo='" + sn + "'"
        ttd = get_DataTable(tmpsql)
        tmpsql = "select sum(Amount) from MMSInvoiceApplyD where DocNo='" + sn + "'"
        ti = get_DataTable(tmpsql).Rows(0).Item(0)
        'tmpsql = "SELECT ISNULL(MAX(DAYSALLOWED),0) AS DAYSALLOWED FROM MMSCONTRACT "
        'tmpsql += " WHERE CONTRACTNO IN "
        'tmpsql += " (SELECT CONTRACTNO FROM MMSInvoiceApplyD WHERE DocNo='" + sn + "')"
        'ttq = get_DataTable(tmpsql)
        'If ttq.Rows(0).Item("DAYSALLOWED") > 0 Then
        '    CUSDAYALLOW = ttq.Rows(0).Item("DAYSALLOWED")
        'Else
        '    '發票內沒有合約保養項目，改以發票日期檢視該客戶是否有其他合約資料
        '    tmpsql = "SELECT ISNULL(MAX(DAYSALLOWED),0) AS DAYSALLOWED FROM MMSCONTRACT "
        '    tmpsql += " WHERE CUSTOMERNO='" + ttm.Rows(0).Item("CustomerNo") + "' AND '" + ttd.Rows(0).Item("InvoiceDate") + "' BETWEEN STARTDATEC AND ENDDATEC "
        '    ttq = get_DataTable(tmpsql)
        '    If ttq.Rows(0).Item("DAYSALLOWED") > 0 Then
        '        CUSDAYALLOW = ttq.Rows(0).Item("DAYSALLOWED")
        '    Else
        '        '沒有合約以客戶基本資料為主
        '        tmpsql = "SELECT DAYSALLOWED FROM MMSCUSTOMERS "
        '        tmpsql += " WHERE CUSTOMERNO='" + ttm.Rows(0).Item("CustomerNo") + "'"
        '        ttq = get_DataTable(tmpsql)
        '        CUSDAYALLOW = ttq.Rows(0).Item("DAYSALLOWED")
        '    End If
        'End If
        'AZITEMP 20151007 要增加發票收款日數檢核抓取
        '發票資料
        tmpsql = "insert into MMSInvoiceM([InvoicePeriod],[InvoiceNo],[InvoiceType],[InvoiceDate],[CustomerNo]"
        tmpsql += ",[GUINumber],[TitleOfInvoice],[NotTaxAmount],[Tax],[Amount],[ReceiptAmount],[UnReceiptAmount]"
        tmpsql += ",[Status],[ApplyDocNo],DaysAllowed)"
        tmpsql += " values('" + ttd.Rows(0).Item("InvoicePeriod") + "'," 'InvoicePeriod
        tmpsql += "'" + ttd.Rows(0).Item("InvoiceNo") + "','" 'InvoiceNo
        If ttd.Rows(0).Item("InvoiceNo") = "" Then 'InvoiceType
            tmpsql += "1"
        Else
            tmpsql += "2"
        End If
        tmpsql += "',"
        tmpsql += "'" + ttd.Rows(0).Item("InvoiceDate") + "'," 'InvoiceDate
        tmpsql += "'" + ttm.Rows(0).Item("CustomerNo") + "'," 'CustomerNo

        If (ttm.Rows(0).Item("CustomerNo") = "A00136") Or (ttm.Rows(0).Item("CustomerNo") = "C00029") Or (ttm.Rows(0).Item("CustomerNo") = "D00004") _
                                                       Or (ttm.Rows(0).Item("CustomerNo") = "E00089") Or (ttm.Rows(0).Item("CustomerNo") = "F00007") Then '20140814 零售客戶
            tmpsql += "'" + ttm.Rows(0).Item("APPLYGUIN") + "'," 'GUINumber
            tmpsql += "'" + ttm.Rows(0).Item("APPLYTITLE") + "'," 'TitleOfInvoice
        Else
            tmpsql += "'" + ttm.Rows(0).Item("GUINumber") + "'," 'GUINumber
            tmpsql += "'" + ttm.Rows(0).Item("TitleOfInvoice") + "'," 'TitleOfInvoice
        End If
        tmpsql += Math.Round(ti / 1.05).ToString.Trim + "," 'NotTaxAmount
        tmpsql += Str(ti - Math.Round(ti / 1.05)).Trim + "," 'Tax
        tmpsql += ti.ToString.Trim + "," 'Amount
        tmpsql += "0," 'ReceiptAmount
        tmpsql += ti.ToString.Trim + "," 'UnReceiptAmount
        tmpsql += "'Y','" + sn + "',0)" 'Status
        'tmpsql += "'Y','" + sn + "'," + CUSDAYALLOW.ToString + ")" 'Status
        EXE_SQL(tmpsql)
        Dim ts As String
        tmpsql = "select max(sn) from MMSInvoiceM "
        ts = get_DataTable(tmpsql).Rows(0).Item(0)
        '發票明細   
        Dim i As Integer
        For i = 0 To ttd.Rows.Count - 1
            tmpsql = "insert into MMSInvoiceD([InvoicePeriod],[InvoiceNo],[Serial],[ItemName],[UnitPrice],[Quantity]"
            tmpsql += ",[Amount],[Memo],[snM])"
            tmpsql += " values('" + ttd.Rows(i).Item("InvoicePeriod") + "'," 'InvoicePeriod
            tmpsql += "'" + ttd.Rows(i).Item("InvoiceNo") + "'," 'InvoiceNo
            tmpsql += "'" + ttd.Rows(i).Item("Serial").ToString.Trim + "'," 'Serial
            tmpsql += "'" + ttd.Rows(i).Item("ItemName") + "'," 'ItemName
            tmpsql += ttd.Rows(i).Item("UnitPrice").ToString.Trim + "," 'UnitPrice
            tmpsql += ttd.Rows(i).Item("Quantity").ToString.Trim + "," 'Quantity
            tmpsql += ttd.Rows(i).Item("Amount").ToString.Trim + "," 'Amount
            tmpsql += "'" + ttd.Rows(i).Item("Memo") + "',+'" + ts + "')" 'Memo
            EXE_SQL(tmpsql)
        Next
    End Sub

    Private Function getDocNo(ByVal sn As String) As String
        Dim tt As DataTable
        tt = Session("BatchGrid1")
        Dim i As Integer
        For i = 0 To tt.Rows.Count - 1
            If tt.Rows(i).Item("sn") = sn Then
                getDocNo = tt.Rows(i).Item("DocNo")
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
        Dim CONNSTR1 As String = modDB.GetDataSource.ConnectionString
        Dim CONN1 As New System.Data.SqlClient.SqlConnection(CONNSTR1)
        Dim SQLCOMMAND1 As New System.Data.SqlClient.SqlCommand
        Try
            SQLCOMMAND1.CommandText = TMPSQL1
            If CONN1.State = ConnectionState.Open Then CONN1.Close()
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

    Protected Sub CheckBox2_CheckedChanged(sender As Object, e As System.EventArgs)
        RememberOldValues()
        Dim categoryIDList As ArrayList = Session("CHECKED_ITEMS")
        If categoryIDList.Count > 0 Then
            Me.Button3.Enabled = True
        Else
            Me.Button3.Enabled = False
        End If
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

    Protected Sub CustFrom_TextChanged(sender As Object, e As System.EventArgs) Handles CustFrom.TextChanged
        If Me.CustTo.Text.Trim = "" Then
            Me.CustTo.Text = Me.CustFrom.Text
        End If
    End Sub
End Class
