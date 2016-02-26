'******************************************************************************************************
'* 程式：ATT_Q_053 打卡資料查詢與補登維護
'* 作成：陳盈志
'* 版次：2013/05/31(VER1.01)：新開發
'******************************************************************************************************

Partial Class ATT_Q_053
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
                If ViewState("ROL").Substring(1, 1) = "N" Then Me.btnNew.Enabled = False
                modUtil.SetObjReadOnly(form1, "txtStnNAME")
                modUtil.SetObjReadOnly(form1, "txtEmpNAME")
                'If modUtil.IsStnRol(Me.Request) Then '* 加油站權限
                If (modUtil.IsStnRol(Me.Request) Or modUnset.IsPAUnitRol(Me.Request)) Then
                    Me.txtStnID.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                    Me.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                    Me.btnStnID.Visible = False
                    modUtil.SetObjReadOnly(form1, "txtStnID")
                    Me.txtDtFrom.Focus()
                Else
                    Me.txtStnID.Focus()
                End If
            End If

            Me.lblTitle.Text = "打卡資料-查詢"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList, 20)   '* 套用gridview樣式
            Me.grdList.Width = 900

            '--------------------------------------------------* 加入欄位並標題顯示中文
            'modDB.SetFields("STNID", "排班單位", grdList, , HorizontalAlign.Center)
            modDB.SetFields("STNNAME", "打卡單位", grdList, HorizontalAlign.Center)
            modDB.SetFields("EMPID", "員工職號", grdList, HorizontalAlign.Center)
            modDB.SetFields("EMPNAME", "員工姓名", grdList, HorizontalAlign.Center)
            modDB.SetFields("FINDATE", "打卡日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("FINTIME", "打卡時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("FINCLANM", "打卡類別", grdList, HorizontalAlign.Center)
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

            Me.txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            modUtil.SetObjReadOnly(Me, "txtSName")

            modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
            modUtil.SetDateImgObj(Me.imgDtFrom, Me.txtDtFrom, True, False, Me.txtDtTo)
            modUtil.SetDateImgObj(Me.imgDtTo, Me.txtDtTo, True, False)

            '* 業務處/責任區/加油站之畫面設定控制
            modUtil.SetRolScreen(Me.txtStnID, Me.txtStnName, Me.btnStnID, Me.ddlBus, Me.ddlArea, Me.ViewState("ROL"), Me.ViewState("RolDeptID"))
            If Me.ViewState("RolType") = "STN" Then Me.txtDtFrom.Focus()

            '----------------------------------------------* 重新顯示查詢結果
            Me.lblMsg.Text = Request("msg")

            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 5, Me.ViewState("EmptyTable"))
            Else
                'modUtil.showMsg(Me.Page, "(QryField)", "QryField IS NOT CLEAR")
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
        Session("QryField") = Me.ViewState("QryField")
        Session("iCurPage") = Me.ViewState("iCurPage")
        'modDB.InsertSignRecord("AziTest", "Rol=" + Request.Cookies("ROL").Value, My.User.Name) 'HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
        With Me.grdList.SelectedRow
            If Not ((Request.Cookies("ROL").Value = "01") Or (Request.Cookies("ROL").Value = "07")) And .Cells(5).Text = "正常" Then
                modUtil.showMsg(Me.Page, "注意", "正常打卡資料不可作異動!")
            Else
                Session("FINDATE") = .Cells(3).Text
                If .Cells(5).Text = "正常" Then '正常打卡資料只能改員工代號
                    'modDB.InsertSignRecord("AziTest", "FINTIME=" & .Cells(4).Text, My.User.Name)
                    Response.Redirect("ATT_M_053.aspx?STNID=" & txtStnID.Text & "&EMPID=" & .Cells(1).Text & "&FINDATE=" & .Cells(3).Text & "&FINTIME=" & .Cells(4).Text & "&FormMode=edit")
                Else '補登資料只能刪除
                    Response.Redirect("ATT_M_053.aspx?STNID=" & txtStnID.Text & "&EMPID=" & .Cells(1).Text & "&FINDATE=" & .Cells(3).Text & "&FINTIME=" & .Cells(4).Text)
                End If
            End If
            '
            'If .Cells(5).Text = "正常" Then
            '    If (Request.Cookies("ROL").Value = "01") Or (Request.Cookies("ROL").Value = "07") Then
            '        Response.Redirect("ATT_M_053.aspx?STNID=" & txtStnID.Text & "&EMPID=" & .Cells(1).Text & "&FINDATE=" & .Cells(3).Text & "&FINTIME=" & .Cells(4).Text)
            '    Else

            '    End If
            'Else
            '    Session("FINDATE") = .Cells(3).Text
            '    'Session("VSTNID") = txtStnID.Text
            '    'Session("VSTNNAME") = txtStnName.Text
            '    Response.Redirect("ATT_M_053.aspx?STNID=" & txtStnID.Text & "&EMPID=" & .Cells(1).Text & "&FINDATE=" & .Cells(3).Text & "&FINTIME=" & .Cells(4).Text)
            'End If
        End With
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
    '* 新增一筆資料
    '******************************************************************************************************
    Protected Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        If txtStnID.Text <> "" Then
            Session("QryField") = Me.ViewState("QryField")
            Session("iCurPage") = Me.ViewState("iCurPage")
            Response.Redirect("ATT_M_053.aspx?STNID=" & txtStnID.Text & "&STNNAME=" & txtStnName.Text & "&FormMode=add")
        Else
            modUtil.showMsg(Me.Page, "錯誤訊息(新增補登資料)", "需先指定單位")
        End If
        'Response.Redirect("ATT_M_051.aspx?FormMode=add&DEPTID=" & Me.ViewState("RolDeptID") & "&STNID=" & Me.ViewState("STNID") & "&STNNAME=" & Me.ViewState("STNNAME"))
        'modUtil.showMsg(Me.Page, "TEST", "Request.Cookies(STNID).Value= " & HttpUtility.UrlDecode(Request.Cookies("STNID").Value))
    End Sub

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        Response.Redirect("ATT_Q_053.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        modUtil.GridView2Excel("ATTQ053.xls", Me.grdList)
    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        Dim sMsg As String = ""
        Dim sValidOK As Boolean = True

        Try
            Me.lblMsg.Text = ""

            sValidOK = modUtil.Check2DateObj(Me, "txtDtFrom", "txtDtTo", sMsg, "查詢日期")
            '
            If (modUtil.IsStnRol(Me.Request) Or modUnset.IsPAUnitRol(Me.Request)) Then
                If ((Me.txtStnID.Text = "") Or (Me.txtStnName.Text.Trim() = "")) Then
                    sValidOK = False
                    sMsg = sMsg + "\n查無此加油站！"
                End If
            End If
            '
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
            'modDB.InsertSignRecord("AziTest", "uSER RolDeptID=" & Me.ViewState("RolDeptID") & " ddlBus=" & Me.ddlBus.SelectedValue & " ddlarea=" & Me.ddlArea.SelectedValue, My.User.Name)
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(查詢)", ex.Message)
        End Try
    End Sub

    '******************************************************************************************************
    '* 顯示明細
    '******************************************************************************************************
    Private Sub ShowGrid()
        Dim sSql As String = "", sFields As New Hashtable
        Dim DTFROM As String = txtDtFrom.Text.Substring(0, 4) & txtDtFrom.Text.Substring(5, 2) & txtDtFrom.Text.Substring(8, 2)
        Dim DTtO As String = txtDtTo.Text.Substring(0, 4) & txtDtTo.Text.Substring(5, 2) & txtDtTo.Text.Substring(8, 2)
        sFields.Add("txtStnID", txtStnID.Text)
        sFields.Add("txtStnNAME", txtStnName.Text)
        sFields.Add("txtDtFrom", txtDtFrom.Text)
        sFields.Add("txtDtTo", txtDtTo.Text)
        sFields.Add("txtEMPID", txtEmpID.Text)
        sFields.Add("txtEMPNAME", txtEmpName.Text)
        'modUtil.showMsg(Me.Page, "Test", "txtStnID= " & txtStnID.Text & " txtDtFrom= " & txtDtFrom.Text & " txtDtTo.Text= " & txtDtTo.Text)
        sSql = "SELECT A.STNID,STNNAME,A.EMPID AS EMPID,M.EMPL_NAME AS EMPNAME,FINDATE,FINTIME" _
             & "      ,CASE SUBSTRING(FINCLA,1,1) WHEN '1' THEN '補登' ELSE '正常' END AS FINCLANM FROM FINGER A " _
             & "  LEFT JOIN [MP_HR].DBO.EMPLOYEE M ON A.EMPID=M.EMPL_ID " _
             & "  LEFT JOIN MECHSTNM S ON A.STNID=S.STNID "
        sSql = sSql & "WHERE (1=1) "
        If txtStnID.Text <> "" Then
            sSql = sSql & "AND ((A.STNID='" & txtStnID.Text & "')" _
                    & "    OR (A.EMPID IN (SELECT EMPL_ID FROM [MP_HR].DBO.EMPLOYEE " _
                    & "                      WHERE EMPL_DEPTID='" & txtStnID.Text & "')))"
        ElseIf Me.ddlArea.SelectedValue <> "" Then
            sSql = sSql & "AND A.STNID IN (SELECT STNID FROM [10.1.1.100].SMILE_HQ.DBO.MECHSTNM_VIEW" _
                        & "                  WHERE ARE_ID='" & Me.ddlArea.SelectedValue & "')"
        ElseIf Me.ddlBus.SelectedValue <> "" Then
            sSql = sSql & "AND A.STNID IN (SELECT STNID FROM [10.1.1.100].SMILE_HQ.DBO.MECHSTNM_VIEW" _
                        & "                  WHERE BUS_ID='" & Me.ddlBus.SelectedValue & "')"
        End If
        '
        sSql = sSql & "  AND FINDATE>='" & DTFROM & "'" _
                    & "  AND FINDATE<='" & DTtO & "'"
        If txtEmpID.Text <> "" Then
            sSql = sSql & " AND EMPID='" & txtEmpID.Text & "'"
        End If

        'FOR 油品部查詢
        If (Me.ddlArea.SelectedValue = "") And (Me.ddlBus.SelectedValue = "") And (Me.ViewState("RolDeptID").ToString.Substring(0, 4) = "4701") Then
            sSql = sSql & "  AND EMPID IN (SELECT EMPL_ID FROM MP_HR.DBO.EMPLOYEE WHERE EMPL_LEV_DATE IS NULL AND EMPL_POSIID='00007')"
        End If

        If Me.rdoType.SelectedValue = "1" Then
            sSql = sSql & " Order By A.STNID,Empid,FINDATE,FINTIME"
        Else
            sSql = sSql & " Order By FINDATE,FINTIME,EMPID"
        End If

        Me.ViewState("QryField") = sFields
        Me.dscList.SelectCommand = sSql
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 5, Me.ViewState("EmptyTable"))
            'If IsPostBack Then modUtil.showMsg(Me.Page, "無資料", sSql)
            If IsPostBack Then modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            ViewState("Sql") = sSql '* 用於 Excel/換頁
        End If
    End Sub

#End Region
    Protected Sub btnEmp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryEmp.Show(Me.txtStnID.Text)
    End Sub

    Protected Sub txtEmpID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call modCharge.GetEmpName(Me.txtEmpID, Me.txtEmpName, True)
    End Sub

    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryStn.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
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

    Protected Sub txtStnID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'modCharge.GetStnName(Me.txtStnID, Me.txtStnName, Me.ViewState("RolDeptID"))
        modCharge.GetStnName(txtStnID, Me.txtStnName, IIf(Trim(Me.ddlArea.SelectedValue) = "", Me.ddlBus.SelectedValue, Me.ddlArea.SelectedValue))
    End Sub
End Class
