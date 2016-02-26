'******************************************************************************************************
'* 程式：ATT_Q_051 排班資料查詢
'* Rebuilded from 陳盈志
'******************************************************************************************************

Partial Class ATT_Q_051
    Inherits System.Web.UI.Page

#Region "Page Event"
    '******************************************************************************************************
    '* 初始化及權限 設定
    '******************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'modDB.InsertSignRecord("AziTest", "att_i_051:ok-9-a", My.User.Name)
        If Not IsPostBack Then
            '--------------------------------------------------* 檢查是否有登錄User.Identity.Name
            'modDB.InsertSignRecord("AziTest", "att_i_051:ok-9-0", My.User.Name)
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else '---------------------------------------------* 檢查是否有 讀取/新增 權限
                modUtil.GetRolData(Me.Request, Me.ViewState("ROL"), Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                If ViewState("ROL").Substring(1, 1) = "N" Then Me.btnNew.Enabled = False
                'modDB.InsertSignRecord("AziTest", "att_i_051:ok-9-1", My.User.Name)
                modUtil.SetObjReadOnly(form1, "txtStnNAME")
                modUtil.SetObjReadOnly(form1, "txtEmpNAME")
                If modUtil.IsStnRol(Me.Request) Then '* 加油站權限
                    'modDB.InsertSignRecord("AziTest", "att_i_051:ok-9-2", My.User.Name)
                    Me.txtStnID.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                    Me.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                    Me.btnStnID.Visible = False
                    modUtil.SetObjReadOnly(form1, "txtStnID")
                    Me.txtDtFrom.Focus()
                Else
                    'modDB.InsertSignRecord("AziTest", "att_i_051:ok-9-3", My.User.Name)
                    Me.txtStnID.Focus()
                End If
            End If

            Me.lblTitle.Text = "排班資料-查詢"       '* 功能標題
            modDB.SetGridViewStyle(Me.grdList, 20)   '* 套用gridview樣式
            Me.grdList.Width = 900

            '--------------------------------------------------* 加入欄位並標題顯示中文
            'modDB.SetFields("STNID", "排班單位", grdList, , HorizontalAlign.Center)
            modDB.SetFields("STNNAME", "排班單位", grdList, HorizontalAlign.Center)
            modDB.SetFields("EMPID", "員工職號", grdList, HorizontalAlign.Center)
            modDB.SetFields("EMPNAME", "員工姓名", grdList, HorizontalAlign.Center)
            modDB.SetFields("SHSTDATE", "排班日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("SHSTTIME", "起始時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("SHEDTIME", "終止時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("WORKHOUR", "出勤時數", grdList, HorizontalAlign.Center)
            modDB.SetFields("RtSTTIME", "午休起時", grdList, HorizontalAlign.Center)
            modDB.SetFields("RtEDTIME", "午休終時", grdList, HorizontalAlign.Center)

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


            '----------------------------------------------* 重新顯示查詢結果
            Me.lblMsg.Text = Request("msg")

            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 8, Me.ViewState("EmptyTable"))
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
        With Me.grdList.SelectedRow
            Response.Redirect("ATT_M_051.aspx?STNID=" & txtStnID.Text & "&EMPID=" & .Cells(1).Text & "&SHSTDATE=" & .Cells(3).Text & "&SHSTTIME=" & .Cells(4).Text)
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
        If txtStnID.Text.Trim <> "" Then
            Session("QryField") = Me.ViewState("QryField")
            Session("iCurPage") = Me.ViewState("iCurPage")
            Response.Redirect("ATT_I_051.aspx?STNID=" & txtStnID.Text.Trim & "&FormMode=add")
        Else
            modUtil.showMsg(Me.Page, "錯誤訊息(新增)", "\n 需先選擇站別!")
        End If
        
    End Sub

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        Response.Redirect("ATT_Q_051.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        modUtil.GridView2Excel("ATTQ051.xls", Me.grdList)
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

            If ((Me.txtStnID.Text = "") Or (Me.txtStnName.Text.Trim() = "")) Then
                sValidOK = False
                sMsg = sMsg + "\n查無此加油站！"
            End If

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

        sFields.Add("txtStnID", txtStnID.Text)
        sFields.Add("txtStnNAME", txtStnName.Text)
        sFields.Add("txtDtFrom", txtDtFrom.Text)
        sFields.Add("txtDtTo", txtDtTo.Text)
        sFields.Add("txtEMPID", txtEmpID.Text)
        'sFields.Add("txtEMPNAME", txtEMPNAME.Text)
        'modUtil.showMsg(Me.Page, "Test", "txtStnID= " & txtStnID.Text & " txtDtFrom= " & txtDtFrom.Text & " txtDtTo.Text= " & txtDtTo.Text)
        sSql = "SELECT A.STNID,S.STNNAME,A.EMPID AS EMPID,M.EMPL_NAME AS EMPNAME" _
             & "      ,A.SHSTDATE AS SHSTDATE,A.SHSTTIME AS SHSTTIME,A.SHEDDATE AS SHEDDATE,A.SHEDTIME AS SHEDTIME " _
             & "      ,A.CLASSID AS CLASSID,A.WORKHOUR AS WORKHOUR " _
             & "      ,A.RtSTTIME AS RtSTTIME,A.RtEDTIME AS RtEDTIME " _
             & "  FROM SCHEDM A " _
             & "  LEFT JOIN [MP_HR].DBO.EMPLOYEE M ON A.EMPID=M.EMPL_ID " _
             & "  LEFT JOIN MECHSTNM S ON A.STNID=S.STNID "
        sSql = sSql & "WHERE A.SHSTDATE>='" & txtDtFrom.Text.Substring(0, 4) & txtDtFrom.Text.Substring(5, 2) & txtDtFrom.Text.Substring(8, 2) & "'" _
                    & "   AND A.SHSTDATE<='" & txtDtTo.Text.Substring(0, 4) & txtDtTo.Text.Substring(5, 2) & txtDtTo.Text.Substring(8, 2) & "'" _
                    & "   AND ((A.STNID='" & txtStnID.Text & "')" _
                    & "     OR (A.EMPID IN (SELECT EMPL_ID FROM [MP_HR].DBO.EMPLOYEE " _
                    & "                      WHERE EMPL_DEPTID='" & txtStnID.Text & "'))) "
        If txtEmpID.text <> "" Then
            sSql = sSql & " AND EMPID='" & txtEmpID.Text & "'"
        End If

        If Me.rdoType.SelectedValue = "1" Then
            sSql = sSql & " Order By Empid,Shstdate,Shsttime"
        Else
            sSql = sSql & " Order By Shstdate,Shsttime"
        End If

        Me.ViewState("QryField") = sFields
        Me.dscList.SelectCommand = sSql
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 8, Me.ViewState("EmptyTable"))
            'If IsPostBack Then modUtil.showMsg(Me.Page, "無資料", sSql)
            If IsPostBack Then modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            ViewState("Sql") = sSql '* 用於 Excel/換頁
        End If
    End Sub

#End Region

    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryStn.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
    End Sub

    Protected Sub txtStnID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'modUtil.showMsg(Me.Page, "test", "RolDeptID=" & Me.ViewState("RolDeptID"))
        'Me.txtStnID, Me.txtStnName, Me.ViewState("RolDeptID")
        modCharge.GetStnName(Me.txtStnID, Me.txtStnName, Me.ViewState("RolDeptID"))
    End Sub

    Protected Sub btnEmp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryEmp.Show(Me.txtStnID.Text)
    End Sub

    Protected Sub txtEmpID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call modCharge.GetEmpName(Me.txtEmpID, Me.txtEmpName, True)
    End Sub

End Class
