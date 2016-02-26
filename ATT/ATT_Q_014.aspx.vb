'******************************************************************************************************
'* 程式：OIL_Q_021 區請假檢核查詢
'* 作成：AZICHEN
'* 版次：2014/09/16(VER1.01)：新開發
'* 非完成!!
'******************************************************************************************************

Partial Class ATT_Q_014
    Inherits System.Web.UI.Page

#Region "Page Event"

    '******************************************************************************************************
    '* 初始化及權限 設定 
    '******************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '--------------------------------------------------* 檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else '---------------------------------------------* 檢查是否有讀取權限
                modUtil.GetRolData(Me.Request, Me.ViewState("ROL"), Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
            End If

            Me.lblTitle.Text = "區請假檢核查詢"
            '--------------------------------------------------* 設定GridView樣式
            modDB.SetGridViewStyle(Me.grdList, 15)
            Me.grdList.RowStyle.Height = 20
            modDB.SetFields("StnName", "站名", grdList, HorizontalAlign.Left)
            modDB.SetFields("EMPID", "員工代號", grdList, HorizontalAlign.Center)
            modDB.SetFields("EMPNAME", "員工姓名", grdList, HorizontalAlign.Center)
            modDB.SetFields("DtTrade", "請假別", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATMDATEST", "請假起始日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATMTIMEST", "請假起始時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATMDATEEN", "請假迄止日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATMTIMEEN", "請假迄止時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATM_Hours", "請假時數", grdList, HorizontalAlign.Right, "{0:N0}")
            '
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 10, Me.ViewState("EmptyTable"))

            '--------------------------------------------------* 設定欄位屬性
            Me.txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
            modUtil.SetDateImgObj(Me.imgDtFrom, Me.txtDtFrom, True, False, Me.txtDtTo)
            modUtil.SetDateImgObj(Me.imgDtTo, Me.txtDtTo, True, False)

            '* 業務處/責任區/加油站之畫面設定控制
            modUtil.SetRolScreen(Me.txtStnID, Me.txtStnName, Me.btnStnID, Me.ddlBus, Me.ddlArea, Me.ViewState("ROL"), Me.ViewState("RolDeptID"))
            If Me.ViewState("RolType") = "STN" Then Me.txtDtFrom.Focus()
        End If

        Me.dscList.SelectCommand = ViewState("Sql")
        Me.ScriptManager1.RegisterPostBackControl(Me.btnExcel)
    End Sub

#End Region

#Region "GridView /TextBox /DropDownList Event"
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
        modCharge.GetStnName(txtStnID, Me.txtStnName, IIf(Trim(Me.ddlArea.SelectedValue) = "", Me.ddlBus.SelectedValue, Me.ddlArea.SelectedValue))
    End Sub

    '******************************************************************************************************
    '* 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.DataBound
        modDB.SetGridPageNum(Me.grdList, PagerButtons.NumericFirstLast, HorizontalAlign.Left)
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
        Me.QryStn.ShowBySel(Me.ddlBus.SelectedValue, Me.ddlArea.SelectedValue)
    End Sub

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Response.Redirect("OIL_Q_021.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        modUtil.GridView2Excel("OilTankLeft.xls", Me.grdList)
    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        Dim sSQL As String, sMsg As String = ""
        Dim sValidOK As Boolean = True

        Try
            Me.txtStnID.Text = Me.txtStnID.Text.Trim
            sValidOK = modUtil.Check2DateObj(Me, "txtDtFrom", "txtDtTo", sMsg, "調撥日期")
            If Me.txtStnName.Text = "" And Me.txtStnID.Text <> "" Then sValidOK = False : sMsg = "\n查無此加油站！"

            If Not sValidOK Then
                modUtil.showMsg(Me.Page, "訊息", sMsg) : Exit Sub
            End If

            sSQL = "Select A.DtAllot, A.AllotID, A.StnOut, C.StnName, D.SName, A.OilQty," _
                 & "(B.DtTrade+' '+B.TmTrade) As DtTrade, B.TradeID, B.TankID, " _
                 & "B.OilQty As RealOilQty, B.Inda From OilAllotM A " _
                 & "Left Join OilAllotTrade B ON B.AllotID = A.AllotID " _
                 & "Inner Join MechStnM_View C ON C.StnID = A.StnOut "

            '------------------* 加入業務處/責任區/加油站之過濾
            If Me.txtStnID.Text <> "" Then
                sSQL = sSQL & "And C.StnID = '" & Me.txtStnID.Text & "' "
            ElseIf Trim(Me.ddlArea.SelectedValue) <> "" Then
                sSQL = sSQL & "And C.Are_ID = '" & Me.ddlArea.SelectedValue & "' "
            ElseIf Trim(Me.ddlBus.SelectedValue) <> "" Then
                sSQL = sSQL & "And C.Bus_ID = '" & Me.ddlBus.SelectedValue & "' "
            End If

            sSQL = sSQL & "Inner Join OilItem D ON D.OilID = A.OilID " _
                 & "Where DtAllot Between '" & txtDtFrom.Text & "'  AND '" & txtDtTo.Text & "' "

            Select Case Me.rdoType.SelectedValue
                Case "1" '--------* 全部
                Case "2" : sSQL = sSQL & "And B.DtTrade Is Null " '* 未發油者
                Case "3" : sSQL = sSQL & "And Not (B.DtTrade Is Null) " '* 已發油者
                Case "4" : sSQL = sSQL & "And A.OilQty <> (Select IsNull(Round(Sum(OilQty), 0), 0) From OilAllotTrade " _
                                & "Where OilAllotTrade.AllotID = A.AllotID And OilAllotTrade.Inda='Y')"
            End Select
            sSQL = sSQL & "Order By A.DtAllot, A.StnOut"
            Call ShowGrid(sSQL) '* 顯示查詢結果

        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(查詢)", ex.Message)
        End Try
    End Sub

    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid(ByVal pSql As String)
        Me.dscList.SelectCommand = pSql
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 10, Me.ViewState("EmptyTable"))
            modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            ViewState("Sql") = pSql '* 用於 Excel/換頁
        End If
    End Sub

#End Region


End Class
