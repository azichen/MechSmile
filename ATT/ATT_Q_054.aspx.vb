'******************************************************************************************************
'* 程式：ATT_Q_054 排班代碼資料維護
'* 作成：陳盈志
'* 版次：2013/06/03(VER1.01)：新開發
'******************************************************************************************************

Partial Class ATT_Q_054
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
                If modUtil.IsStnRol(Me.Request) Then '* 加油站權限
                    Me.txtStnID.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                    Me.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                    Me.btnStnID.Visible = False
                    modUtil.SetObjReadOnly(form1, "txtStnID")
                Else
                    Me.txtStnID.Focus()
                End If
            End If

            Me.lblTitle.Text = "排班代碼資料-查詢"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList, 20)   '* 套用gridview樣式
            Me.grdList.Width = 900

            '--------------------------------------------------* 加入欄位並標題顯示中文
            'modDB.SetFields("STNID", "排班單位", grdList, , HorizontalAlign.Center)
            modDB.SetFields("STNNAME", "打卡單位", grdList, HorizontalAlign.Center)
            modDB.SetFields("CLSID", "排班代碼", grdList, HorizontalAlign.Center)
            modDB.SetFields("STTIME", "起始時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("EDTIME", "結束時間", grdList, HorizontalAlign.Center)
            modDB.SetFields("NMEMO", "備註", grdList, HorizontalAlign.Center)
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

            '----------------------------------------------* 重新顯示查詢結果
            Me.lblMsg.Text = Request("msg")

            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 4, Me.ViewState("EmptyTable"))
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
            If .Cells(5).Text = "正常" Then
                modUtil.showMsg(Me.Page, "注意", "正常打卡資料不可作異動!")
            Else
                'Session("FINDATE") = .Cells(3).Text
                'Session("VSTNID") = txtStnID.Text
                'Session("VSTNNAME") = txtStnName.Text
                Response.Redirect("ATT_M_054.aspx?STNID=" & txtStnID.Text & "&CLSID=" & .Cells(1).Text)
            End If
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
        Response.Redirect("ATT_M_054.aspx?FormMode=add")
        'Response.Redirect("ATT_M_051.aspx?FormMode=add&DEPTID=" & Me.ViewState("RolDeptID") & "&STNID=" & Me.ViewState("STNID") & "&STNNAME=" & Me.ViewState("STNNAME"))
        'modUtil.showMsg(Me.Page, "TEST", "Request.Cookies(STNID).Value= " & HttpUtility.UrlDecode(Request.Cookies("STNID").Value))
    End Sub

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        Response.Redirect("ATT_Q_054.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        modUtil.GridView2Excel("ATTQ054.xls", Me.grdList)
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
        'modUtil.showMsg(Me.Page, "Test", "txtStnID= " & txtStnID.Text & " txtDtFrom= " & txtDtFrom.Text & " txtDtTo.Text= " & txtDtTo.Text)
        sSql = "SELECT A.STNID,STNNAME,CLSID,STTIME,EDTIME,CASE NFLAG WHEN 'Y' THEN '正常班' ELSE '' END AS NMEMO FROM CLSTIME A " _
             & "  LEFT JOIN MECHSTNM S ON A.STNID=S.STNID " _
             & " WHERE A.STNID='" & txtStnID.Text & "'" _
             & " ORDER BY CLSID"
        
        Me.ViewState("QryField") = sFields
        Me.dscList.SelectCommand = sSql
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 5, Me.ViewState("EmptyTable"))
            If IsPostBack Then modUtil.showMsg(Me.Page, "無資料", sSql)
            'If IsPostBack Then modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            ViewState("Sql") = sSql '* 用於 Excel/換頁
        End If
    End Sub

#End Region
    Protected Sub btnEmp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryEmp.Show(Me.txtStnID.Text)
    End Sub

    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryStn.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
    End Sub

    Protected Sub txtStnID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'modUtil.showMsg(Me.Page, "test", "RolDeptID=" & Me.ViewState("RolDeptID"))
        'Me.txtStnID, Me.txtStnName, Me.ViewState("RolDeptID")
        modCharge.GetStnName(Me.txtStnID, Me.txtStnName, Me.ViewState("RolDeptID"))
    End Sub
End Class
