'******************************************************************************************************
'* 程式：ATT_Q_059 四周出勤周期設定維護
'* 作成：陳盈志
'* 版次：2015/03/27 (VER1.01)：新開發
'******************************************************************************************************
Imports System.Data
Partial Class ATT_Q_059
    Inherits System.Web.UI.Page

#Region "Page Event"
    '******************************************************************************************************
    '* 初始化及權限 設定
    '******************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '--------------------------------------------------* 檢查是否有登錄User.Identity.Name
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else '---------------------------------------------* 檢查是否有 讀取/新增 權限
                modUtil.GetRolData(Me.Request, Me.ViewState("ROL"), Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                Me.txtDtFrom.Focus()
            End If

            Me.lblTitle.Text = "雙週出勤周期設定-查詢"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList, 20)   '* 套用gridview樣式
            Me.grdList.Width = 500

            Dim vYear As Integer
            ddlYear.Items.Clear()
            For vYear = Year(Now()) To (Year(Now()) + 5)
                ddlYear.Items.Add(vYear.ToString)
            Next
            txtDtFrom.Text = ddlYear.Text + "/01/01"
            txtDtTo.Text = ddlYear.Text + "/12/31"
            '--------------------------------------------------* 加入欄位並標題顯示中文
            'modDB.SetFields("STNID", "排班單位", grdList, , HorizontalAlign.Center)
            modDB.SetFields("FWDS_ID", "周期序", grdList, HorizontalAlign.Center)
            modDB.SetFields("FWDS_DATEST", "起始日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("FWDS_DATEEN", "迄止日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("FWDS_STNWH", "標準時數", grdList, HorizontalAlign.Center)
            '--------------------------------------------------* 加入連結欄位
            Dim sField As New CommandField
            sField.ShowSelectButton = True
            sField.HeaderText = "功能"
            sField.SelectText = "瀏覽操作"
            sField.HeaderStyle.ForeColor = Drawing.Color.White
            sField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            Me.grdList.Columns.Add(sField)
            Me.grdList.RowStyle.Height = 20
            '--------------------------------------------------* 設定欄位屬性
            modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
            modUtil.SetDateImgObj(Me.imgDtFrom, Me.txtDtFrom, True, False, Me.txtDtTo)
            modUtil.SetDateImgObj(Me.imgDtTo, Me.txtDtTo, True, False)
            '
            'modDB.InsertSignRecord("AziTest", "Q059_Rol=" + Me.ViewState("ROL"), My.User.Name) 'AZITEST

            Button1.Visible = False
            If (Mid(Me.ViewState("ROL"), 5, 2) = "01") Or (Mid(Me.ViewState("ROL"), 5, 2) = "07") Then
                Button1.Visible = True
            End If
            '----------------------------------------------* 重新顯示查詢結果
            Me.lblMsg.Text = Request("msg")

            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 4, Me.ViewState("EmptyTable"))
            Else
                Call ReflashQryData()
            End If
        End If

        Me.dscList.SelectCommand = ViewState("Sql")
        Me.ScriptManager1.RegisterPostBackControl(Me.btnExcel)
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

#End Region

#Region "GridView / Text Event"
    '******************************************************************************************************
    '* GridView_DataBound => 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.DataBound
        modDB.SetGridPageNum(Me.grdList, PagerButtons.NumericFirstLast, HorizontalAlign.Left)
    End Sub

    '******************************************************************************************************
    '* 保存瀏覽頁碼
    '******************************************************************************************************
    Protected Sub grdList_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.PageIndexChanged
        Me.ViewState("iCurPage") = grdList.PageIndex
    End Sub

    '******************************************************************************************************
    '* 執行 瀏覽操作 
    '******************************************************************************************************
    Protected Sub grdList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.SelectedIndexChanged
        If (Mid(Me.ViewState("ROL"), 5, 2) = "01") Or (Mid(Me.ViewState("ROL"), 5, 2) = "07") Then
            Session("QryField") = Me.ViewState("QryField")
            Session("iCurPage") = Me.ViewState("iCurPage")
            With Me.grdList.SelectedRow
                Session("FWDS_DATEST") = .Cells(1).Text
                Session("FWDS_DATEEN") = .Cells(2).Text
                Response.Redirect("ATT_M_059.aspx?FWDS_DATEST=" & .Cells(1).Text)
            End With
        Else
            modUtil.showMsg(Me.Page, "訊息", "不可操作")
        End If
    End Sub

    '******************************************************************************************************
    '* GridView_RowDataBound => 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub grdList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdList.RowDataBound
        modDB.SetGridLightPen(e)
    End Sub

#End Region

#Region "Button Event"

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        Response.Redirect("ATT_Q_059.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        modUtil.GridView2Excel("ATTQ059.xls", Me.grdList)
    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        Dim sMsg As String = ""
        Dim sValidOK As Boolean = True

        Try
            Me.lblMsg.Text = ""

            If sValidOK Then
                If ((Me.txtDtFrom.Text = "") Or (Me.txtDtTo.Text = "")) Then
                    sValidOK = False
                    sMsg = sMsg + "\n 需輸入查詢日期範圍！"
                End If
            End If

            '---------------------------------------------* 檢核欄位正確性
            If Not sValidOK Then
                modUtil.showMsg(Me.Page, "訊息", sMsg) : Exit Sub
            End If
            Call ShowGrid() '* 顯示查詢結果

        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(查詢)", ex.Message)
        End Try
    End Sub

    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        Dim sSql As String = "", sFields As New Hashtable
        sFields.Add("txtDtFrom", txtDtFrom.Text)
        sFields.Add("txtDtTo", txtDtTo.Text)

        sSql = "SELECT FWDS_ID,FWDS_DATEST,FWDS_DATEEN" _
             & "      ,ISNULL(FWDS_STNWH,0) AS FWDS_STNWH FROM FWDUTYSH " _
             & " WHERE FWDS_DATEEN>='" & txtDtFrom.Text & "'  AND FWDS_DATEST<='" & txtDtTo.Text & "'" _
             & " Order By FWDS_DATEST"

        Me.ViewState("QryField") = sFields
        Me.dscList.SelectCommand = sSql
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 4, Me.ViewState("EmptyTable"))
            If IsPostBack Then modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            ViewState("Sql") = sSql '* 用於 Excel/換頁
        End If
    End Sub

#End Region

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim BaseDate As Date = Convert.ToDateTime("2015/03/29")
        Dim VDATEST, VDATEEN As String
        Dim sSql As String = ""
        Dim tt As DataTable
        Dim Sid As String = "00"
        Dim VSTNWH As Double = 0
        'modDB.InsertSignRecord("AziTest", "Q059_OK-1", My.User.Name) 'AZITEST
        sSql = "SELECT FWDS_ID,FWDS_DATEEN FROM FWDUTYSH WHERE FWDS_DATEEN =(SELECT ISNULL(MAX(FWDS_DATEEN),'') AS LASTDATE FROM FWDUTYSH) "
        'LOGSQL(sSql)
        tt = get_DataTable(sSql)
        If tt.Rows.Count > 0 Then
            'modDB.InsertSignRecord("AziTest", "Q059_OK-2", My.User.Name) 'AZITEST
            'If tt.Rows(0).Item("LASTDATE") > Convert.ToDateTime("2010/01/01") Then
            Sid = tt.Rows(0).Item("FWDS_ID")
            BaseDate = Convert.ToDateTime(tt.Rows(0).Item("FWDS_DATEEN"))
            'modDB.InsertSignRecord("AziTest", "Q059_LASTDATE=" + tt.Rows(0).Item("LASTDATE").ToString, My.User.Name) 'AZITEST
            'End If
        Else
            'modDB.InsertSignRecord("AziTest", "Q059_OK-3", My.User.Name) 'AZITEST
            Sid = Year(BaseDate).ToString + "00"
        End If
        'modDB.InsertSignRecord("AziTest", "Q059_LASTDATE=" + " BaseDate=" + BaseDate.ToString, My.User.Name) 'AZITEST
        '
        If Year(BaseDate) = Convert.ToInt16(ddlYear.SelectedValue) Then
            While (Year(BaseDate.AddDays(1)) = Convert.ToInt16(ddlYear.SelectedValue))
                Sid = Sid.Substring(0, 4) + Format(Convert.ToInt16(Sid.Substring(4, 2)) + 1, "00")
                VDATEST = BaseDate.AddDays(1).ToString("yyyy/MM/dd")
                VDATEEN = BaseDate.AddDays(14).ToString("yyyy/MM/dd")
                VSTNWH = 84
                sSql = "SELECT COUNT(*) AS HOLIQTY FROM MP_HR.DBO.HOLIDAY " _
                     & " WHERE HOLI_DATE BETWEEN '" & VDATEST & "' AND '" & VDATEEN & "'"
                tt = get_DataTable(sSql)
                VSTNWH = 84 - (tt.Rows(0).Item("HOLIQTY") * 8)
                sSql = "INSERT INTO FWDUTYSH (FWDS_ID,FWDS_DATEST,FWDS_DATEEN,FWDS_STNWH) VALUES " _
                     & " ('" + Sid + "','" + VDATEST + "','" + VDATEEN + "'," + VSTNWH.ToString + ")"
                'LOGSQL(sSql)
                EXE_SQL(sSql)
                BaseDate = BaseDate.AddDays(14) '改以雙周作設定
            End While
            ShowGrid()
            modUtil.showMsg(Me.Page, "訊息", "設定完成!")
        Else
            modUtil.showMsg(Me.Page, "訊息", "目前最後設定日期為: " + BaseDate.ToString + " ,需設定年度為: " + Year(BaseDate.AddDays(1)).ToString)
        End If
    End Sub

    '連接資料庫取得datatable
    Public Function get_DataTable(ByVal SQL1 As String) As System.Data.DataTable
        Dim CONNSTR1 As String = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
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
        Dim CONNSTR1 As String = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
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

    Protected Sub ddlYear_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtDtFrom.Text = ddlYear.Text + "/01/01"
        txtDtTo.Text = ddlYear.Text + "/12/31"
    End Sub

End Class
