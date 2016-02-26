'******************************************************************************************************
'* 程式：ATT_Q_060 加油站合理工時維護
'* 作成：
'* 版次：2015/07/31 (VER1.01)：新開發
'******************************************************************************************************
Imports System.Data

Partial Class ATT_Q_060
    Inherits System.Web.UI.Page
    Private cWriteLog As Boolean = False '* 正式上線時須設成 False

#Region "Page Event"
    '******************************************************************************************************
    '* 初始化及權限 設定 
    '******************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '----------------------------------------* 檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else '-----------------------------------* 檢查是否有 讀取/新增 權限
                modUtil.GetRolData(Me.Request, Me.ViewState("ROL"), Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                'If ViewState("ROL").Substring(1, 1) = "N" Then Me.btnNew.Enabled = False
            End If

            Me.TitleLabel.Text = "考勤鎖檔設定作業"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList2, 20)   '* 套用gridview樣式
            Me.grdList2.Width = 500

            modDB.SetFields("STNNAME", "單位", grdList2, HorizontalAlign.Center)
            modDB.SetFields("ACCYM", "計薪月份", grdList2, HorizontalAlign.Center)
            modDB.SetFields("KINDNAME", "週期別", grdList2, HorizontalAlign.Center)
            modDB.SetFields("DATESTEN", "起迄區間", grdList2, HorizontalAlign.Center)
            modDB.SetFields("STATUS", "狀態", grdList2, HorizontalAlign.Center)

            Dim TDATE As String = ""
            TDATE = DateTime.Now.ToString("yyyy/MM/dd")
            SHOWPERIOD(TDATE)
            'TitleLabel.Text = Me.ViewState("RolDeptID")
            'Call modDB.ShowEmptyDataGridHeader(Me.grdList2, 0, 2, Me.ViewState("EmptyTable"))


            Me.lblMsg.Text = "以下單位已將此週期設定為「鎖定」"
            'btnLOCK.Attributes.Add("OnClick", "return confirm('是否確定鎖定?');")
           
            '* 業務處/責任區/加油站之畫面設定控制
            modUtil.SetRolScreen(Me.txtStnID, Me.txtStnName, Me.btnStnID, Me.ddlBus, Me.ddlArea, Me.ViewState("ROL"), Me.ViewState("RolDeptID"))
        End If

        Me.dscList.SelectCommand = ViewState("Sql")
        'Me.ScriptManager1.RegisterPostBackControl(Me.btnExcel)
    End Sub


#End Region

#Region "GridView /TextBox /DropDownList Event"
    '******************************************************************************************************
    '* GridView_DataBound => 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList2.DataBound
        modDB.SetGridPageNum(Me.grdList2, PagerButtons.NumericFirstLast, HorizontalAlign.Left)
    End Sub

    '******************************************************************************************************
    '* 保存瀏覽頁碼
    '******************************************************************************************************
    Protected Sub grdList_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList2.PageIndexChanged
        Me.ViewState("iCurPage") = grdList2.PageIndex
    End Sub

    '******************************************************************************************************
    '* 回主畫面時，重新顯示查詢結果
    '******************************************************************************************************
    'Protected Sub ddlArea_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlBus.DataBound, ddlArea.DataBound
    '    If Not IsPostBack Then
    '        If Session("QryField") IsNot Nothing Then Call ReflashQryData(sender)
    '    End If
    'End Sub

    '******************************************************************************************************
    '* 選取業務處/責任區時，清除加油站
    '******************************************************************************************************
    Protected Sub ddlBus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlBus.SelectedIndexChanged, ddlArea.SelectedIndexChanged
        If Me.ViewState("RolType") <> "STN" Then Me.txtStnID.Text = "" : Me.txtStnName.Text = ""
        SHOWPERIOD(LblDATEEN.Text)
    End Sub

    '******************************************************************************************************
    '* GridView_RowDataBound => 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub grdList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdList2.RowDataBound
        modDB.SetGridLightPen(e)
    End Sub


    '******************************************************************************************************
    '* 選取業務處時，先清除舊責任區，再附加
    '******************************************************************************************************
    Protected Sub dscArea_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles dscArea.Selecting
        If (Not IsPostBack) And (Mid(Me.ViewState("ROL"), 5, 1) > "1") Then '* 0:總 1:業 2:區 3,4:站
            e.Cancel = True '* 已依權限之設定畫面，不可重新讀取
        Else
            Me.ddlArea.Items.Clear()
            Me.ddlArea.Items.Add("")
        End If
        SHOWPERIOD(LblDATEEN.Text)
    End Sub

    '******************************************************************************************************
    '* 顯示加油站名稱(依所選取之業務處/責任區過濾)
    '******************************************************************************************************
    Protected Sub txtStnID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStnID.TextChanged
        'modDB.InsertSignRecord("AziTest", "txtStnID_TextChanged-1", My.User.Name) 'AZITEST
        modCharge.GetStnName(txtStnID, Me.txtStnName, IIf(Trim(Me.ddlArea.SelectedValue) = "", Me.ddlBus.SelectedValue, Me.ddlArea.SelectedValue))
        SHOWPERIOD(LblDATEEN.Text)
    End Sub

#End Region

    Protected Sub SHOWPERIOD(ByVal SPDATE As String)
        Dim tmpsql As String
        Dim tt As DataTable
        Dim HRLOCKDT As String
        Dim CHKDEPT, LOCKDEPT, LOCKDEPTNAME, VLOCK As String
        Dim VPKIND As String
        If Me.rdoType.SelectedValue = "1" Then
            VPKIND = "M"
        Else
            VPKIND = "W"
        End If
        '20150804
        If (SPDATE = "2015/07/20") And (VPKIND = "W") Then
            SPDATE = "2015/07/21"
        End If
        Dim SDATE As String = SPDATE.Substring(0, 4) + SPDATE.Substring(5, 2) + SPDATE.Substring(8, 2)
        '先確認檢查單位
        If Me.txtStnID.Text <> "" Then
            CHKDEPT = Me.txtStnID.Text
        ElseIf Me.ddlBus.SelectedValue <> "" Then
            CHKDEPT = Me.ddlBus.SelectedValue + "10"
        Else
            CHKDEPT = "000000"
        End If
        '
        VLOCK = ""
        LOCKDEPT = ""
        tmpsql = "SELECT CLCT_YEAR,CLCT_MONTH FROM MP_HR.DBO.CLOSE_CONTROL WHERE CLCT_COMPID='01' AND CLCT_KIND='0' "
        tt = get_DataTable(tmpsql)
        HRLOCKDT = tt.Rows(0).Item(0).ToString + "/" + String.Format("{0:00}", tt.Rows(0).Item(1)) + "/20"
        '
        If VPKIND = "M" Then
            If SDATE.Substring(6, 2) > "20" Then
                LblACCYM.Text = DateTime.Parse(SPDATE).AddMonths(1).ToString("yyyyMM")
            Else
                LblACCYM.Text = DateTime.Parse(SPDATE).ToString("yyyyMM")
                'LblDATEEN.Text.Substring(0, 4) + LblDATEEN.Text.Substring(5, 2)
            End If
            '
            LblFWDSID.Text = ""
            LblDATEST.Text = DateTime.Parse(LblACCYM.Text.Substring(0, 4) + "/" + LblACCYM.Text.Substring(4, 2) + "/01").AddMonths(-1).ToString("yyyy/MM") + "/21"
            LblDATEEN.Text = LblACCYM.Text.Substring(0, 4) + "/" + LblACCYM.Text.Substring(4, 2) + "/20"
            tmpsql = "SELECT ACCYM,STNID,STDATE,EDDATE " _
                   + "      ,CASE LOCKCODE WHEN 'Y' THEN 'Y' ELSE 'N' END AS LOCK" _
                   + "  FROM MECHSYA " _
                   + " WHERE '" + SDATE + "' BETWEEN STDATE AND EDDATE " _
                   + "   AND STNID IN ('000000','" + CHKDEPT.Substring(0, 2) + "0110','" + CHKDEPT + "') AND LOCKCODE='Y'" _
                   + "   AND PKIND='M'" _
                   + " ORDER BY STNID DESC "
            '
            tt.Clear()
            tt = get_DataTable(tmpsql)
            '
            If tt.Rows.Count > 0 Then
                VLOCK = "Y"
                LOCKDEPT = tt.Rows(0).Item(1).ToString
            Else
                VLOCK = "N"
                LOCKDEPT = ""
            End If
            '
            'modDB.InsertSignRecord("AziTest", "kind=M 薪月 SDATE=" + SDATE + " Accym=" + LblACCYM.Text, My.User.Name) 'AZITEST
        Else
            tmpsql = "SELECT FWDS_ID,FWDS_DATEST,FWDS_DATEEN " _
                       + "      ,CASE LOCKCODE WHEN 'Y' THEN 'Y' ELSE 'N' END AS LOCK,STNID" _
                       + "  FROM FWDUTYSH A LEFT JOIN MECHSYA ON '" + SDATE + "' BETWEEN STDATE AND EDDATE " _
                       + "                AND STNID IN ('000000','" + CHKDEPT.Substring(0, 2) + "0110','" + CHKDEPT + "') AND LOCKCODE='Y'" _
                       + "                AND PKIND='" + VPKIND + "'" _
                       + "  LEFT JOIN MP_HR.DBO.CLOSE_CONTROL B ON CLCT_COMPID='01' AND CLCT_KIND='0' " _
                       + " WHERE '" + SPDATE + "' BETWEEN FWDS_DATEST AND FWDS_DATEEN " _
                       + " ORDER BY STNID DESC "
            '
            tt = get_DataTable(tmpsql)
            '
            LblFWDSID.Text = tt.Rows(0).Item(0)
            LblDATEST.Text = tt.Rows(0).Item(1)
            LblDATEEN.Text = tt.Rows(0).Item(2)
            VLOCK = tt.Rows(0).Item(3)
            'HRLOCKDT = tt.Rows(0).Item(4).ToString + "/" + String.Format("{0:00}", tt.Rows(0).Item(5)) + "/20"
            LOCKDEPT = tt.Rows(0).Item(4).ToString
            If LblDATEEN.Text.Substring(8, 2) > "20" Then
                LblACCYM.Text = DateTime.Parse(LblDATEEN.Text).AddMonths(1).ToString("yyyyMM")
            Else
                LblACCYM.Text = DateTime.Parse(LblDATEEN.Text).ToString("yyyyMM")
                'LblDATEEN.Text.Substring(0, 4) + LblDATEEN.Text.Substring(5, 2)
            End If
            'modDB.InsertSignRecord("AziTest", "kind=W 雙週 SPDATE=" + SPDATE, My.User.Name) 'AZITEST
        End If
        'modDB.InsertSignRecord("AziTest", "SPDATE=" + SPDATE + " LblDATEST.Text=" + LblDATEST.Text + " LblDATEen.Text=" + LblDATEEN.Text + " LOCK=" + tt.Rows(0).Item(3), My.User.Name) 'AZITEST
        '
        btnLOCK.Visible = False
        btnUNLOCK.Visible = False
        LOCKDEPTNAME = ""
        Dim MAINLOCK As Boolean = False
        If VLOCK = "Y" Then
            LblSTATUS.ForeColor = Drawing.Color.Red
            LblSTATUS.Text = "已鎖定"
            If (LOCKDEPT = "000000") Then
                LblSTATUS.Text = LblSTATUS.Text
                LOCKDEPTNAME = "(總部)"
                MAINLOCK = True
            ElseIf (LOCKDEPT.Substring(2, 4) = "0110") Then
                LblSTATUS.Text = LblSTATUS.Text
                If (LOCKDEPT.Substring(1, 2) = "41") Then
                    LOCKDEPTNAME = "(台北業務處)"
                ElseIf (LOCKDEPT.Substring(1, 2) = "43") Then
                    LOCKDEPTNAME = "(台中業務處)"
                ElseIf (LOCKDEPT.Substring(1, 2) = "44") Then
                    LOCKDEPTNAME = "(嘉南業務處)"
                ElseIf (LOCKDEPT.Substring(1, 2) = "45") Then
                    LOCKDEPTNAME = "(高雄業務處)"
                End If
            Else
                LblSTATUS.Text = LblSTATUS.Text
                LOCKDEPTNAME = "(" + Me.txtStnName.Text + ")"
            End If
            '
            'LblSTATUS.Text = LblSTATUS.Text + " " + CHKDEPT + " " + tt.Rows(0).Item(3)
            LblSTATUS.Text = LblSTATUS.Text + " " + LOCKDEPTNAME
            btnLOCK.Visible = False
            'modDB.InsertSignRecord("AziTest", "SPDATE=" + SPDATE + " HRLOCKDT=" + HRLOCKDT, My.User.Name) 'AZITEST
            If SPDATE > HRLOCKDT Then
                '業務處以上權限才可解鎖
                'modDB.InsertSignRecord("AziTest", "ROL=" + Me.ViewState("ROL") + " chk=" + Mid(Me.ViewState("ROL"), 5, 2) + " CHKDEPT=" + CHKDEPT, My.User.Name) 'AZITEST
                '只有人資與總部管理者
                If (Mid(Me.ViewState("ROL"), 5, 2) = "01") Or (Mid(Me.ViewState("ROL"), 5, 2) = "07") Or ((Mid(Me.ViewState("ROL"), 5, 2) = "11") And (CHKDEPT <> "000000") And Not MAINLOCK) Then
                    'modDB.InsertSignRecord("AziTest", "unlock is enabled", My.User.Name) 'AZITEST
                    btnUNLOCK.Visible = True
                End If
            End If
        Else
            LblSTATUS.ForeColor = Drawing.Color.Green
            LblSTATUS.Text = "未鎖定"
            If DateTime.Now.ToString("yyyy/MM/dd") > LblDATEEN.Text Then
                btnLOCK.Visible = True
            End If
            btnUNLOCK.Visible = False
            If Me.txtStnName.Text <> "" Then
                LOCKDEPTNAME = Me.txtStnName.Text
            ElseIf CHKDEPT.Substring(2, 4) = "0110" Then
                If CHKDEPT.Substring(0, 2) = "41" Then
                    LOCKDEPTNAME = "[台北業務處]"
                ElseIf CHKDEPT.Substring(0, 2) = "43" Then
                    LOCKDEPTNAME = "[台中業務處]"
                ElseIf CHKDEPT.Substring(0, 2) = "44" Then
                    LOCKDEPTNAME = "[嘉南業務處]"
                ElseIf CHKDEPT.Substring(0, 2) = "45" Then
                    LOCKDEPTNAME = "[高雄業務處]"
                Else
                    LOCKDEPTNAME = "[業務處]"
                End If
            Else
                LOCKDEPTNAME = "[總部]"
            End If
            '
            btnLOCK.OnClientClick = "return confirm('麻煩「再次」確認:是否將「" + LOCKDEPTNAME + "」 " + LblDATEST.Text + "~" + LblDATEEN.Text + " 期間 考勤鎖定');"
        End If
        '應該要增加是否已人資鎖檔,若已鎖檔則不再顯示「解鎖檔」
        '增加資料查詢
        Dim sSql As String = ""
        sSql = " SELECT A.STNID,A.ACCYM" _
             + "   ,CASE A.STNID WHEN '000000' THEN '總公司' " _
             + "                 WHEN '410110' THEN '台北業務處' WHEN '430110' THEN '台中業務處'" _
             + "                 WHEN '440110' THEN '嘉南業務處' WHEN '450110' THEN '高雄業務處'" _
             + "                 ELSE B.STNNAME END AS STNNAME" _
             + "   ,CASE PKIND WHEN 'M' THEN '薪月' WHEN 'W' THEN '雙週' ELSE '' END AS KINDNAME" _
             + "   ,CONVERT(VARCHAR(10),CONVERT(DATETIME,STDATE,112),111)+' ~ '+CONVERT(VARCHAR(10),CONVERT(DATETIME,EDDATE,112),111) AS DATESTEN" _
             + "   ,CASE LOCKCODE WHEN 'Y' THEN '鎖定' ELSE '未鎖定' END AS STATUS" _
             + "   FROM MECHSYA A LEFT JOIN MECHSTNM B ON A.STNID=B.STNID" _
             + "  WHERE ('" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd") + "' BETWEEN STDATE AND EDDATE) AND LOCKCODE='Y' "
        '+ "   ,STDATE+'~'+EDDATE AS DATESTEN" _
        '+ "  WHERE STDATE='" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd") + "' AND LOCKCODE='Y'"
        If CHKDEPT.Substring(2, 4) = "0110" Then '業務處
            sSql = sSql + "    AND ((A.STNID='000000') OR (SUBSTRING(A.STNID,1,2)='" + CHKDEPT.Substring(0, 2) + "'))"
        Else
            If CHKDEPT <> "000000" Then '單站
                sSql = sSql + "    AND A.STNID IN ('000000','" + CHKDEPT.Substring(0, 2) + "0110','" + CHKDEPT + "')"
            End If
        End If
        sSql = sSql + "  ORDER BY STNID"

        Me.dscList.SelectCommand = sSql
        'Me.grdList2.DataSourceID = Me.dscList.ID
        Me.grdList2.DataSource = Me.dscList
        Me.grdList2.DataBind()

        If Me.grdList2.Rows.Count = 0 Then
            'modDB.InsertSignRecord("AziTest", "grdList.Rows.Count = 0 STDATE=" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd"), My.User.Name) 'AZITEST
            Call modDB.ShowEmptyDataGridHeader(Me.grdList2, 0, 4, Me.ViewState("EmptyTable"))
            'If IsPostBack Then modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            ViewState("Sql") = sSql '* 用於 Excel/換頁
            'modDB.InsertSignRecord("AziTest", "grdList.Rows.Count =" + grdList2.Rows.Count.ToString + " datest=" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd"), My.User.Name) 'AZITEST
        End If

    End Sub

#Region "Button Event"
    '******************************************************************************************************
    '* 選擇加油站(依所選取之業務處/責任區過濾)
    '******************************************************************************************************
    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStnID.Click
        'modDB.InsertSignRecord("AziTest", "btnStnID_Click-1", My.User.Name) 'AZITEST
        Me.QryStn.ShowBySel(Me.ddlBus.SelectedValue, Me.ddlArea.SelectedValue)
    End Sub

#End Region

    '連接資料庫取得datatable
    Public Function get_DataTable(ByVal SQL1 As String) As System.Data.DataTable
        Dim CONNSTR1 As String = modUnset.PA_GetDataSource.ConnectionString
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
        Dim CONNSTR1 As String = modUnset.PA_GetDataSource.ConnectionString
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

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '前一週期
        Dim TDATE As String = ""
        'Dim DDATE As DateTime = DateTime.Parse(LblDATEST.Text)
        TDATE = DateTime.Parse(LblDATEST.Text).AddDays(-1).ToString("yyyy/MM/dd")
        'modDB.InsertSignRecord("AziTest", "TDATE = " + TDATE, My.User.Name) 'AZITEST
        SHOWPERIOD(TDATE)
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '下一週期
        Dim TDATE As String = ""
        'Dim DDATE As DateTime = DateTime.Parse(LblDATEST.Text)
        TDATE = DateTime.Parse(LblDATEEN.Text).AddDays(1).ToString("yyyy/MM/dd")
        'modDB.InsertSignRecord("AziTest", "TDATE = " + TDATE, My.User.Name) 'AZITEST
        SHOWPERIOD(TDATE)
    End Sub

    Protected Sub btnUNLOCK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim CHKDEPT As String
        Dim VACCYM As String
        If LblDATEEN.Text.Substring(8, 2) > "20" Then
            VACCYM = DateTime.Parse(LblDATEEN.Text).AddMonths(1).ToString("yyyyMM")
        Else
            VACCYM = LblDATEEN.Text.Substring(0, 4) + LblDATEEN.Text.Substring(5, 2)
        End If
        '先確認檢查單位
        If Me.txtStnID.Text <> "" Then
            CHKDEPT = Me.txtStnID.Text
        ElseIf Me.ddlBus.SelectedValue <> "" Then
            CHKDEPT = Me.ddlBus.SelectedValue + "10"
        Else
            CHKDEPT = "000000"
        End If
        '
        Dim VPKIND As String
        If Me.rdoType.SelectedValue = "1" Then
            VPKIND = "M"
        Else
            VPKIND = "W"
        End If
        modDB.InsertSignRecord("ATTQ060", "PKIND=" + VPKIND + " PERIOD_DTST=" + LblDATEST.Text + " STNID=" + CHKDEPT + " UNLOCK", My.User.Name)
        Dim tmpsql As String
        tmpsql = "UPDATE MECHSYA SET LOCKCODE='N',INUSER='" + My.User.Name + "'" _
               + " ,INDATE='" + DateTime.Now.ToString("yyyyMMdd") + "'" _
               + " ,INTIME='" + DateTime.Now.ToString("HHmmss") + "'" _
               + " WHERE ACCYM='" + VACCYM + "' AND STNID='" + CHKDEPT + "' AND STDATE='" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd") + "'" _
               + "   AND PKIND='" + VPKIND + "'"
        EXE_SQL(tmpsql)        
        '重新顯示
        SHOWPERIOD(LblDATEST.Text)
    End Sub

    Protected Sub btnLOCK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '增加檢查是否已過週期最後日
        If DateTime.Now.ToString("yyyy/MM/dd") > LblDATEEN.Text Then

            LblSTATUS.Text = "鎖定處理"
            Dim CHKDEPT As String
            Dim VACCYM As String
            If LblDATEEN.Text.Substring(8, 2) > "20" Then
                VACCYM = DateTime.Parse(LblDATEEN.Text).AddMonths(1).ToString("yyyyMM")
            Else
                VACCYM = LblDATEEN.Text.Substring(0, 4) + LblDATEEN.Text.Substring(5, 2)
            End If
            '先確認檢查單位
            If Me.txtStnID.Text <> "" Then
                CHKDEPT = Me.txtStnID.Text
            ElseIf Me.ddlBus.SelectedValue <> "" Then
                CHKDEPT = Me.ddlBus.SelectedValue + "10"
            Else
                CHKDEPT = "000000"
            End If
            '
            Dim tmpsql As String
            '先刪除後存檔
            'modDB.InsertSignRecord("AziTest", "INUSER = " + My.User.Name, My.User.Name) 'AZITEST
            '
            Dim VPKIND As String
            If Me.rdoType.SelectedValue = "1" Then
                VPKIND = "M"
            Else
                VPKIND = "W"
            End If
            modDB.InsertSignRecord("ATTQ060", "PKIND=" + VPKIND + " PERIOD_DTST=" + LblDATEST.Text + " STNID=" + CHKDEPT + " LOCK", My.User.Name)
            '
            'modify in 20151210
            Dim tt As DataTable
            tmpsql = "SELECT ACCYM FROM MECHSYA WHERE ACCYM='" + VACCYM + "' AND STDATE='" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd") + "' AND STNID='" + CHKDEPT + "'" _
                   + "                      AND PKIND='" + VPKIND + "'"
            '
            'tt.Clear()
            tt = get_DataTable(tmpsql)
            '
            If tt.Rows.Count > 0 Then
                tmpsql = "UPDATE MECHSYA SET LOCKCODE='Y',INUSER='" + My.User.Name + "'" _
                       + " ,INDATE='" + DateTime.Now.ToString("yyyyMMdd") + "'" _
                       + " ,INTIME='" + DateTime.Now.ToString("HHmmss") + "'" _
                       + " WHERE ACCYM='" + VACCYM + "' AND STNID='" + CHKDEPT + "' AND STDATE='" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd") + "'" _
                       + "   AND PKIND='" + VPKIND + "'"
            Else
                'tmpsql = "DELETE FROM MECHSYA WHERE ACCYM='" + VACCYM + "' AND STDATE='" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd") + "' AND STNID='" + CHKDEPT + "'" _
                '       + "                      AND PKIND='" + VPKIND + "'"
                tmpsql = "INSERT INTO MECHSYA (ACCYM,STNID,LOCKCODE,FWDSID,STDATE,EDDATE,INUSER,INDATE,INTIME,PKIND,SETTLEDT,SETTLEUSER) VALUES " _
                   + " ('" + VACCYM + "','" + CHKDEPT + "','Y','" + LblFWDSID.Text + "'" _
                   + " ,'" + DateTime.Parse(LblDATEST.Text).ToString("yyyyMMdd") + "','" + DateTime.Parse(LblDATEEN.Text).ToString("yyyyMMdd") + "'" _
                   + " ,'" + My.User.Name + "'" _
                   + " ,'" + DateTime.Now.ToString("yyyyMMdd") + "','" + DateTime.Now.ToString("hhmmss") + "','" + VPKIND + "','','')"
            End If
            '
            EXE_SQL(tmpsql)
            '重新顯示
            SHOWPERIOD(LblDATEST.Text)
        Else
            modUtil.showMsg(Me.Page, "訊息", "不可在周期迄止日前鎖檔!!")
        End If
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Button3.ForeColor = Drawing.Color.DarkGreen
        Button3.Font.Bold = True
        Button4.ForeColor = Drawing.Color.DarkGray
        Button4.Font.Bold = False
        rdoType.SelectedValue = 1
        Dim TDATE As String = ""
        'Dim DDATE As DateTime = DateTime.Parse(LblDATEST.Text)
        TDATE = DateTime.Parse(LblDATEST.Text).ToString("yyyy/MM/dd")
        SHOWPERIOD(TDATE)
    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Button3.ForeColor = Drawing.Color.DarkGray
        Button3.Font.Bold = False
        Button4.ForeColor = Drawing.Color.DarkGreen
        Button4.Font.Bold = True
        rdoType.SelectedValue = 2
        Dim TDATE As String = ""
        'Dim DDATE As DateTime = DateTime.Parse(LblDATEST.Text)
        TDATE = DateTime.Parse(LblDATEST.Text).ToString("yyyy/MM/dd")
        SHOWPERIOD(TDATE)
    End Sub
End Class
