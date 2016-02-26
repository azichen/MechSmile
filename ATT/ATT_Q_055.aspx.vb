'******************************************************************************************************
'* 程式：ATT_Q_055 加油站合理工時維護
'* 作成：
'* 版次：2013/12/30 (VER1.01)：新開發
'******************************************************************************************************

Partial Class ATT_Q_055
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
                If ViewState("ROL").Substring(1, 1) = "N" Then Me.btnNew.Enabled = False
            End If

            Me.TitleLabel.Text = "加油站合理工時維護"
            '----------------------------------------* 設定GridView樣式
            modDB.SetGridViewStyle(Me.grdList, 15)
            Me.grdList.RowStyle.Height = 20
            modDB.SetFields("STNID", "站號", grdList, False)
            modDB.SetFields("STNNAME", "站名", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("HR11", "晴平日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR12", "晴假日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR13", "晴漲日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR14", "晴跌日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR21", "陰平日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR22", "陰假日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR23", "陰漲日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR24", "陰跌日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR31", "雨平日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR32", "雨假日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR33", "雨漲日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("HR34", "雨跌日", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("DIFFADD", "誤差增加時數", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("DIFFSUB", "誤差減少時數", grdList, HorizontalAlign.Right, "", False)
            modDB.SetFields("CHECKFLAG", "是否檢查", grdList, HorizontalAlign.Right, "", False)
            '----------------------------------------* 加入連結欄位
            Dim addColumn1 As New ButtonField()
            addColumn1.CommandName = "cmdBrowse"
            addColumn1.Text = "瀏覽操作"
            addColumn1.HeaderStyle.ForeColor = Drawing.Color.White
            addColumn1.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            addColumn1.ButtonType = ButtonType.Link
            addColumn1.HeaderText = "功能"
            grdList.Columns.Add(addColumn1)
            '----------------------------------------* 設定欄位屬性
            Me.txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()")

            '----------------------------------------* 重新顯示查詢結果
            Me.lblMsg.Text = Request("msg")
            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 16, Me.ViewState("EmptyTable"))
            Else
                'Call ReflashQryData() '* 移至責任區的DataBound
            End If
            '* 業務處/責任區/加油站之畫面設定控制
            modUtil.SetRolScreen(Me.txtStnID, Me.txtStnName, Me.btnStnID, Me.ddlBus, Me.ddlArea, Me.ViewState("ROL"), Me.ViewState("RolDeptID"))
        End If

        Me.dscList.SelectCommand = ViewState("Sql")
        Me.ScriptManager1.RegisterPostBackControl(Me.btnExcel)
    End Sub

    '******************************************************************************************************
    '* 重新顯示查詢結果 
    '******************************************************************************************************
    Private Sub ReflashQryData(ByVal sender As Object)
        'modDB.InsertSignRecord("AziTest", "Reflash-1", My.User.Name) 'AZITEST
        Dim sCol As Hashtable = Session("QryField")
        Dim sKey(sCol.Count), sVal(sCol.Count) As String
        Dim I As Integer
        sCol.Keys.CopyTo(sKey, 0)
        sCol.Values.CopyTo(sVal, 0)
        'modDB.InsertSignRecord("AziTest", "Reflash-2", My.User.Name) 'AZITEST

        For I = 0 To sCol.Count - 1
            If Mid(sKey(I), 1, 3) = "ddl" Then
                'modDB.InsertSignRecord("AziTest", "DropDownList-" & I.ToString & "=" & sVal(I), My.User.Name) 'AZITEST
                CType(Me.form1.FindControl(sKey(I)), DropDownList).SelectedValue = sVal(I)
                'modDB.InsertSignRecord("AziTest", "DropDownList-" & I.ToString & "=" & sVal(I) & " ok", My.User.Name) 'AZITEST
            Else
                'modDB.InsertSignRecord("AziTest", "not DropDownList-" & I.ToString & "=" & sVal(I), My.User.Name) 'AZITEST
                CType(Me.form1.FindControl(sKey(I)), TextBox).Text = sVal(I)
            End If
        Next
        'modDB.InsertSignRecord("AziTest", "Reflash-3", My.User.Name) 'AZITEST

        If sender Is Me.ddlBus Then Exit Sub '* 責任區更新時才重新顯示

        'modDB.InsertSignRecord("AziTest", "Reflash-3-1", My.User.Name) 'AZITEST

        Call ShowGrid()
        grdList.PageIndex = Session("iCurPage")
        Me.ViewState("iCurPage") = Session("iCurPage")
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        'modDB.InsertSignRecord("AziTest", "Reflash-4", My.User.Name) 'AZITEST

    End Sub
#End Region

#Region "GridView /TextBox /DropDownList Event"
    '******************************************************************************************************
    '* 回主畫面時，重新顯示查詢結果
    '******************************************************************************************************
    Protected Sub ddlArea_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlBus.DataBound, ddlArea.DataBound
        If Not IsPostBack Then
            If Session("QryField") IsNot Nothing Then Call ReflashQryData(sender)
        End If
    End Sub

    '******************************************************************************************************
    '* 選取業務處/責任區時，清除加油站
    '******************************************************************************************************
    Protected Sub ddlBus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlBus.SelectedIndexChanged, ddlArea.SelectedIndexChanged
        If Me.ViewState("RolType") <> "STN" Then Me.txtStnID.Text = "" : Me.txtStnName.Text = ""
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
    End Sub

    '******************************************************************************************************
    '* 顯示加油站名稱(依所選取之業務處/責任區過濾)
    '******************************************************************************************************
    Protected Sub txtStnID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStnID.TextChanged
        'modDB.InsertSignRecord("AziTest", "txtStnID_TextChanged-1", My.User.Name) 'AZITEST
        modCharge.GetStnName(txtStnID, Me.txtStnName, IIf(Trim(Me.ddlArea.SelectedValue) = "", Me.ddlBus.SelectedValue, Me.ddlArea.SelectedValue))
    End Sub

    '******************************************************************************************************
    '* 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.DataBound
        'modDB.InsertSignRecord("AziTest", "grdList_DataBound-1", My.User.Name) 'AZITEST
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
    Protected Sub grdList_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles grdList.RowCommand
        Session("QryField") = Me.ViewState("QryField")
        Session("iCurPage") = Me.ViewState("iCurPage")
        Dim index As Integer = Convert.ToInt32(e.CommandArgument)
        Dim row As GridViewRow = grdList.Rows(index)
        If cWriteLog Then modUtil.WriteLog("ATT_M_055.aspx?STNID=" & row.Cells(0).Text.Trim, "ATT_Q_055")
        If e.CommandName = "cmdBrowse" Then
            Response.Redirect("ATT_M_055.aspx?STNID=" & row.Cells(0).Text.Trim)
        End If
    End Sub

    '******************************************************************************************************
    '* 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub grdList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdList.RowDataBound
        modDB.SetGridLightPen(e)
    End Sub
#End Region

#Region "Button Event"
    '******************************************************************************************************
    '* 選擇加油站(依所選取之業務處/責任區過濾)
    '******************************************************************************************************
    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStnID.Click
        'modDB.InsertSignRecord("AziTest", "btnStnID_Click-1", My.User.Name) 'AZITEST
        Me.QryStn.ShowBySel(Me.ddlBus.SelectedValue, Me.ddlArea.SelectedValue)
    End Sub


    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        Response.Redirect("ATT_Q_055.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        modUtil.GridView2Excel("station.xls", Me.grdList)
    End Sub

    '******************************************************************************************************
    '* 新增一筆資料
    '******************************************************************************************************
    Protected Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Session("QryField") = Me.ViewState("QryField")
        Session("iCurPage") = Me.ViewState("iCurPage")
        Response.Redirect("ATT_M_055.aspx?FormMode=add")
    End Sub

    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        Call ShowGrid() '* 顯示查詢結果
        Session("iCurPage") = Nothing
    End Sub

    '******************************************************************************************************
    '* 顯示查詢結果
    '******************************************************************************************************
    Private Sub ShowGrid()
        'modDB.InsertSignRecord("AziTest", "ShowGrid-1", My.User.Name) 'AZITEST
        Dim sSQL As String = "", sFields As New Hashtable

        sSQL = "Select M.STNID, M.STNNAME, HR11,HR12,HR13,HR14,HR21,HR22,HR23,HR24,HR31,HR32,HR33,HR34,DIFFADD,DIFFSUB,CHECKFLAG " _
             + "  From Duty_STNREAL D " _
             + "  Left Join MECHSTNM M ON D.STNID=M.STNID " _
             + "  Left Join STN_AREA R ON M.ARE_ID=R.ARE_ID " _
             + " Where 1=1 "

        'modDB.InsertSignRecord("AziTest", "Me.ddlArea.SelectedValue=" & Me.ddlArea.SelectedValue, My.User.Name) 'AZITEST
        'modDB.InsertSignRecord("AziTest", "Me.ddlBus.SelectedValue=" & Me.ddlBus.SelectedValue, My.User.Name) 'AZITEST

        If Me.txtStnID.Text <> "" Then
            sSQL = sSQL & "And M.StnID = '" & Me.txtStnID.Text & "' "
            sFields.Add("txtStnID", txtStnID.Text)
            sFields.Add("txtStnName", txtStnName.Text)
        ElseIf Trim(Me.ddlArea.SelectedValue) <> "" Then
            sSQL = sSQL & "And M.ARE_ID = '" & Me.ddlArea.SelectedValue & "' "
            sFields.Add("ddlArea", Me.ddlArea.SelectedValue)
        ElseIf Trim(Me.ddlBus.SelectedValue) <> "" Then
            sSQL = sSQL & "And R.Bus_ID = '" & Me.ddlBus.SelectedValue & "' "
            sFields.Add("ddlBus", Me.ddlBus.SelectedValue)
        End If

        'modDB.InsertSignRecord("AziTest", "ShowGrid-2", My.User.Name) 'AZITEST

        Me.ViewState("QryField") = sFields

        '--------------------------------------------* 加入業務處/責任區之過濾

        'If Me.DropDownList3.SelectedValue.Trim <> "" Then
        '    sSQL = sSQL & "And M.BUS_TYPE = '" & Me.DropDownList3.SelectedValue.Trim & "' "
        '    sFields.Add("DropDownList3", DropDownList3.SelectedValue)
        'End If
        'If Trim(Me.ddlBus.SelectedValue) <> "" Then sFields.Add("ddlBus", Me.ddlBus.SelectedValue)
        'If Trim(Me.ddlArea.SelectedValue) <> "" Then sFields.Add("ddlArea", Me.ddlArea.SelectedValue)
        'Me.ViewState("QryField") = sFields '* For 維護AP更新資料後之重新查詢

        sSQL = sSQL & "Order By M.StnID"
        Me.dscList.SelectCommand = sSQL
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 16, Me.ViewState("EmptyTable"))
            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            ViewState("Sql") = sSQL '* 用於 Excel/換頁
        End If
    End Sub
#End Region

End Class
