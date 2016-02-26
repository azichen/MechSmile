'******************************************************************************************************
'* 程式：ATT_Q_057 人事表單申請
'* 作成：陳盈志
'* 版次：2014/06/27(VER1.01)：新開發
'******************************************************************************************************

Partial Class ATT_Q_057
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

            Me.lblTitle.Text = "人事表單申請資料查詢"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList, 20)   '* 套用gridview樣式
            Me.grdList.Width = 900

            '--------------------------------------------------* 加入欄位並標題顯示中文

            'modDB.SetFields("STNID", "單位代號", grdList, HorizontalAlign.Center)
            modDB.SetFields("STNNAME", "單位名稱", grdList, HorizontalAlign.Center)
            modDB.SetFields("EMPID", "員工代號", grdList, HorizontalAlign.Center)
            modDB.SetFields("EMPNAME", "員工姓名", grdList, HorizontalAlign.Center)
            modDB.SetFields("FORMNAME", "申請表單", grdList, HorizontalAlign.Center)
            modDB.SetFields("APPLYDATE", "申請日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("STATUS", "狀態", grdList, HorizontalAlign.Center)
            modDB.SetFields("SERID", "申請序號", grdList, HorizontalAlign.Center)
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
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 7, Me.ViewState("EmptyTable"))
            Else
                'modUtil.showMsg(Me.Page, "(QryField)", "QryField IS NOT CLEAR")
                Call ReflashQryData()
            End If

        End If

        Me.dscList.SelectCommand = ViewState("Sql")
        'Me.ScriptManager1.RegisterPostBackControl(Me.btnExcel)
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
            'If .Cells(5).Text = "已處理" Then
            '    modUtil.showMsg(Me.Page, "注意", "此筆表單申請資料已處理，不可作異動!")
            'Else
            Session("SERID") = .Cells(0).Text
            'Response.Redirect("ATT_M_057.aspx?STNID=" & txtStnID.Text & "&EMPID=" & .Cells(1).Text & "&SHSTDATE=" & .Cells(3).Text & "&SHSTTIME=" & .Cells(4).Text)
            Response.Redirect("ATT_M_057.aspx?SERID=" & .Cells(6).Text & "&FORMID=" & .Cells(3).Text.Substring(0, 3)) ' & "&FINDATE=" & .Cells(3).Text & "&FINTIME=" & .Cells(4).Text
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
        Session("QryField") = Me.ViewState("QryField")
        Session("iCurPage") = Me.ViewState("iCurPage")
        Response.Redirect("ATT_M_057.aspx?FormMode=add")
    End Sub

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        Response.Redirect("ATT_Q_057.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    'Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
    '    modUtil.GridView2Excel("ATTQ057.xls", Me.grdList)
    'End Sub

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
        Dim DTFROM As String = txtDtFrom.Text.Substring(0, 4) & txtDtFrom.Text.Substring(5, 2) & txtDtFrom.Text.Substring(8, 2)
        Dim DTtO As String = txtDtTo.Text.Substring(0, 4) & txtDtTo.Text.Substring(5, 2) & txtDtTo.Text.Substring(8, 2)
        sFields.Add("txtStnID", txtStnID.Text)
        sFields.Add("txtStnNAME", txtStnName.Text)
        sFields.Add("txtDtFrom", txtDtFrom.Text)
        sFields.Add("txtDtTo", txtDtTo.Text)
        sFields.Add("txtEMPID", txtEmpID.Text)
        sFields.Add("txtEMPNAME", txtEmpName.Text)
        '
        sSql = "SELECT HAP_STNID AS STNID,STNNAME,HAP_EMPID AS EMPID,M.EMPL_NAME AS EMPNAME,HAP_FORMID AS FORMID,HAP_APPLYDATE AS APPLYDATE " _
             & "      ,CASE HAP_STATUS WHEN 'N' THEN '未處理' WHEN 'Y' THEN '已處理' WHEN 'I' THEN '不處理' END  AS STATUS" _
             & "      ,(HAP_FORMID+FORM_NAME) AS FORMNAME,HAP_SERID AS SERID " _
             & "  FROM HAFORMApply A " _
             & "  LEFT JOIN [MP_HR].DBO.EMPLOYEE M ON A.HAP_EMPID=M.EMPL_ID " _
             & "  LEFT JOIN MECHSTNM S ON A.HAP_STNID=S.STNID " _
             & "  LEFT JOIN HAFORM F ON A.HAP_FORMID=F.FORM_ID "

        sSql = sSql & "WHERE HAP_DELETOR='' AND HAP_STNID='" & txtStnID.Text & "'" _
                    & "  AND HAP_ApplyDATE>='" & DTFROM & "'" _
                    & "  AND HAP_ApplyDATE<='" & DTtO & "'"
        If txtEmpID.Text <> "" Then
            sSql = sSql & " AND HAP_EMPID='" & txtEmpID.Text & "'"
        End If
        sSql = sSql & " ORDER BY HAP_SERID"

        Me.ViewState("QryField") = sFields
        Me.dscList.SelectCommand = sSql
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 7, Me.ViewState("EmptyTable"))
            If IsPostBack Then modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            'Me.btnExcel.Enabled = True
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

    Protected Sub txtStnID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        modCharge.GetStnName(Me.txtStnID, Me.txtStnName, Me.ViewState("RolDeptID"))
    End Sub
End Class
