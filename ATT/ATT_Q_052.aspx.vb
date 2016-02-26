'******************************************************************************************************
'* 程式：ATT_Q_052 公出資料查詢 -> 2014.02.19 改請假資料
'* 作成：陳盈志
'* 版次：2013/05/21(VER1.01)：新開發
'******************************************************************************************************
Imports modUnset

Partial Class ATT_Q_052
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
                modUtil.SetObjReadOnly(form1, "TxtVANMNAME")
                'If modUtil.IsStnRol(Me.Request) Then '* 加油站權限
                If (modUtil.IsStnRol(Me.Request) Or modUnset.IsPAUnitRol(Me.Request)) Then
                    Me.txtStnID.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                    Me.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                    Me.btnStnID.Visible = False
                    modUtil.SetObjReadOnly(form1, "txtStnID")

                    'If modUtil.IsStnRol(Me.Request) Then '加油站暫不開放類別
                    '    Me.TxtVANMID.Text = "010"
                    '    Me.TxtVANMNAME.Text = "公出假"
                    '    modUtil.SetObjReadOnly(form1, "TxtVANMID")
                    '    modUtil.SetObjReadOnly(form1, "TxtVANMNAME")
                    '    Me.BTNVCM.Visible = False
                    'End If
                    'Me.txtDtFrom.Focus()
                Else
                    Me.txtStnID.Focus()
                End If
            End If

            Me.lblTitle.Text = "請假資料"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList, 20)   '* 套用gridview樣式
            Me.grdList.Width = 900

            '--------------------------------------------------* 加入欄位並標題顯示中文
            modDB.SetFields("VATM_EMPLID", "員工職號", grdList, HorizontalAlign.Center)
            modDB.SetFields("EMPL_NAME", "員工姓名", grdList, HorizontalAlign.Center)
            modDB.SetFields("VACA_NAME", "假別", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATM_DATE_ST", "起始日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATM_DATE_EN", "結束日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATM_TIME_ST", "起始時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATM_TIME_EN", "結束時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATM_HOURS", "請假時數", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATM_U_USER_ID", "最後維護者", grdList, HorizontalAlign.Center)
            modDB.SetFields("VATM_SEQ", "請假序號", grdList, HorizontalAlign.Center)
            modDB.SetFields("CHKMEMO", "檢核", grdList, HorizontalAlign.Center)

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

            'modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            'modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
            'modUtil.SetDateImgObj(Me.imgDtFrom, Me.txtDtFrom, True, False, Me.txtDtTo)
            'modUtil.SetDateImgObj(Me.imgDtTo, Me.txtDtTo, True, False)


            '----------------------------------------------* 重新顯示查詢結果
            Me.lblMsg.Text = Request("msg")

            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 10, Me.ViewState("EmptyTable"))
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
        Session("ATT052MODE") = "1"
        With Me.grdList.SelectedRow
            'modUtil.showMsg(Me.Page, "TEST", "ATT_M_052.aspx?VATM_EMPLID=" & .Cells(0).Text & "&VATM_SEQ=" & .Cells(7).Text)
            '紀錄資料日期 20130527
            Session("VATMDATE") = .Cells(2).Text
            Response.Redirect("ATT_M_052.aspx?ATT052MODE=1&VATM_EMPLID=" & .Cells(0).Text & "&VATM_SEQ=" & .Cells(9).Text)
            'Response.Redirect("ATT_M_052.aspx?FormMode=add&ATT052MODE=1")
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
        Response.Redirect("ATT_M_052.aspx?FormMode=add&ATT052MODE=1")
    End Sub

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        Response.Redirect("ATT_Q_052.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        modUtil.GridView2Excel("ATTQ052.xls", Me.grdList)
    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************
    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        Dim sMsg As String = ""
        Dim sValidOK As Boolean = True

        Try
            Me.lblMsg.Text = ""

            If ((Me.txtStnID.Text = "") Or (Me.txtStnName.Text.Trim() = "")) Then
                sValidOK = False
                sMsg = sMsg + "\n查無此加油站！"
            End If

            'If (Me.txtEmpID.Text = "") Then
            '    sValidOK = False
            '    sMsg = sMsg + "\n 需輸入員工職號！"
            'End If

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
        Dim ACCDATEST As String = ""

        sFields.Add("txtStnID", txtStnID.Text)
        sFields.Add("txtStnNAME", txtStnName.Text)
        sFields.Add("txtEMPID", txtEmpID.Text)
        sFields.Add("txtEMPNAME", txtEmpName.Text)
        sFields.Add("TxtVANMID", TxtVANMID.Text)
        sFields.Add("TxtVANMNAME", TxtVANMNAME.Text)

        Dim ACCYM As String = GET_LOCKYM("2", "000000") '取得目前計薪月份
        If ACCYM.Substring(4, 2) = "01" Then
            ACCDATEST = (Convert.ToInt32(ACCYM.Substring(0, 4)) - 1).ToString + "-01-21"
        Else
            ACCDATEST = ACCYM.Substring(0, 4) + "-" + Format(Convert.ToInt32(ACCYM.Substring(4, 2)) - 1, "00") + "-21"
        End If

        If ACCYM <> "" And ACCYM <> "20991231" Then
            Dim ACDT As String = (Convert.ToInt32(ACCYM.Substring(0, 4)) - 1).ToString & "1221"
            sSql = "SELECT VATM_EMPLID,CONVERT(char(10), EMPL_NAME) AS EMPL_NAME,VATM_VANMID,CONVERT(char(10), VACA_NAME) AS VACA_NAME,VATM_U_USER_ID" _
                 & "      ,CONVERT(char(8), VATM_DATE_ST,112) AS VATM_DATE_ST,CONVERT(char(8), VATM_DATE_EN,112) AS VATM_DATE_EN " _
                 & "      ,VATM_TIME_ST,VATM_TIME_EN,VATM_DAYS,VATM_HOURS,CONVERT(VARCHAR(30),VATM_REMARK) AS VATM_REMARK,VATM_SEQ" _
                 & "      ,CASE D.EFFECTIVE WHEN 'Y' THEN '已核' ELSE '未核' END AS CHKMEMO" _
                 & "  FROM MP_HR.DBO.VACATM A" _
                 & " INNER JOIN MP_HR.DBO.VACAMF B ON A.VATM_VANMID=B.VACA_ID" _
                 & " INNER JOIN MP_HR.DBO.EMPLOYEE C ON A.VATM_EMPLID=C.EMPL_ID" _
                 & "  LEFT JOIN VATMPRNLOG D ON A.VATM_EMPLID=D.EMPID AND A.VATM_SEQ=D.VATMSEQ AND D.VATMDATEU>=A.VATM_U_DATE AND D.EFFECTIVE='Y'" _
                 & " WHERE VATM_EMPLID IN (SELECT EMPL_ID FROM MP_HR.DBO.EMPLOYEE WHERE EMPL_DEPTID='" & txtStnID.Text & "'" _
                 & "                         AND (EMPL_LEV_DATE IS NULL OR EMPL_LEV_DATE='' OR EMPL_LEV_DATE>='" & ACCDATEST & "'))" _
                 & "   AND VATM_DATE_ST>='" & ACDT & "'"

            If txtEmpID.Text.Trim <> "" Then
                sSql = sSql + " AND VATM_EMPLID='" & txtEmpID.Text & "'"
            End If

            If TxtVANMID.Text.Trim <> "" Then
                sSql = sSql & " AND VATM_VANMID='" & TxtVANMID.Text & "'"
            End If
            sSql = sSql & " ORDER BY VATM_DATE_ST DESC,VATM_TIME_ST "
            '
            Me.ViewState("QryField") = sFields
            Me.dscList.SelectCommand = sSql
            Me.grdList.DataSourceID = Me.dscList.ID
            Me.grdList.DataBind()

            If Me.grdList.Rows.Count = 0 Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 10, Me.ViewState("EmptyTable"))
                'If IsPostBack Then modUtil.showMsg(Me.Page, "無資料", sSql)
                If IsPostBack Then modUtil.showMsg(Me.Page, "訊息", "查無資料！")
            Else
                Me.btnExcel.Enabled = True
                ViewState("Sql") = sSql '* 用於 Excel/換頁
            End If
        Else
            modUtil.showMsg(Me.Page, "錯誤", "取得目前計薪月份錯誤！")
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

    Protected Sub btnVCM_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Me.QryVcm.Show()
        If modUtil.IsStnRol(Me.Request) Then '加油站只能新增公出假
            Me.QryVcm.Show()
        Else
            Me.QryVcm2.Show()
        End If
    End Sub

    Protected Sub TxtVANMID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call modUnset.GetVcmName(Me.TxtVANMID, Me.TxtVANMNAME, True)
    End Sub

End Class
