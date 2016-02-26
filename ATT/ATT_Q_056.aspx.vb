'******************************************************************************************************
'* 程式：ATT_Q_056 站假日設定維護
'* 作成：陳盈志
'* 版次：2014/01/06(VER1.01)：新開發
'******************************************************************************************************

Partial Class ATT_Q_056
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
                Me.txtDtFrom.Focus()
            End If

            Me.lblTitle.Text = "站假日設定-查詢"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList, 20)   '* 套用gridview樣式
            Me.grdList.Width = 900

            '--------------------------------------------------* 加入欄位並標題顯示中文
            'modDB.SetFields("STNID", "排班單位", grdList, , HorizontalAlign.Center)
            modDB.SetFields("HDATE", "日期", grdList, HorizontalAlign.Center)
            modDB.SetFields("HOLIDAY", "是否為假日", grdList, HorizontalAlign.Center)
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

            '----------------------------------------------* 重新顯示查詢結果
            Me.lblMsg.Text = Request("msg")

            If Session("QryField") Is Nothing Then
                Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 1, Me.ViewState("EmptyTable"))
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
        Session("QryField") = Me.ViewState("QryField")
        Session("iCurPage") = Me.ViewState("iCurPage")
        With Me.grdList.SelectedRow
            Session("HDATE") = .Cells(0).Text
            Response.Redirect("ATT_M_056.aspx?HDATE=" & .Cells(0).Text)
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
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Session("iCurPage") = Nothing
        Response.Redirect("ATT_Q_056.aspx")
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        modUtil.GridView2Excel("ATTQ056.xls", Me.grdList)
    End Sub

    '******************************************************************************************************
    '* 新增一筆資料
    '******************************************************************************************************
    Protected Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Session("QryField") = Me.ViewState("QryField")
        Session("iCurPage") = Me.ViewState("iCurPage")
        Response.Redirect("ATT_M_056.aspx?FormMode=add")
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

        sSql = "SELECT HDATE,HOLIDAY FROM DUTY_HOLIDAY" _
             & " WHERE HDATE>='" & txtDtFrom.Text & "'  AND HDATE<='" & txtDtTo.Text & "'" _
             & " Order By HDATE"

        Me.ViewState("QryField") = sFields
        Me.dscList.SelectCommand = sSql
        Me.grdList.DataSourceID = Me.dscList.ID
        Me.grdList.DataBind()

        If Me.grdList.Rows.Count = 0 Then
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 1, Me.ViewState("EmptyTable"))
            If IsPostBack Then modUtil.showMsg(Me.Page, "訊息", "查無資料！")
        Else
            Me.btnExcel.Enabled = True
            ViewState("Sql") = sSql '* 用於 Excel/換頁
        End If
    End Sub

#End Region

End Class
