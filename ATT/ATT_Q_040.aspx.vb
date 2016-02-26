'******************************************************************************************************
'* 程式：ATT_Q_040 加油站天氣預報查詢
'* 作成：ERIC LIU
'* 版次：2014/01/24(VER1.00)：新開發
'******************************************************************************************************

Partial Class ATT_ATT_Q_040
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String


#Region "Page Event"

    '******************************************************************************************************
    '* Page_Load => 在此頁面初始化設定
    '******************************************************************************************************

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sRol As String = ""

        If Not IsPostBack Then


            '檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else
                Dim sFileName As String = Request.Path.Substring(Request.Path.LastIndexOf("/") + 1, Request.Path.Length - (Request.Path.LastIndexOf("/") + 1))
                sRol = HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
                ViewState("ROL") = modUtil.GetRolData(sFileName, sRol)
                '1.HttpRequest 2.回傳使用權限 3.回傳使用者層級 4.回傳使用者所屬部門ID
                modUtil.GetRolData(Me.Request, Me.ViewState("ROL"), Me.ViewState("RolType"), Me.ViewState("RolDeptID"))

                '檢查是否有讀取權限
                If ViewState("ROL").Substring(2, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                Me.MsgLabel.Text = ViewState("ROL").ToString
            End If
            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)

            Me.TitleLabel.Text = "各站天氣預報查詢"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList)          '* 套用gridview樣式


            '--------------------------------------------------* 加入欄位並標題顯示中文
            modDB.SetFields("STNID", "站代碼", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("STNNAME", "站名稱", grdList, HorizontalAlign.Left, "", False)
            modDB.SetFields("CutDate", "日期", grdList, HorizontalAlign.Center, "", False)
            modDB.SetFields("WetName", "天氣狀況", grdList, HorizontalAlign.Center, "", False)
            modDB.SetFields("AttName", "分類", grdList, HorizontalAlign.Center, "", False)
            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 4, Me.ViewState("EmptyTable"))

            Me.MsgLabel.Text = Request("msg")
            txtAreaFrom.Focus()

        Else
            '--------------------------------------------------* 光筆效果
            ViewState("Style") = "0"
            If (Not ClientScript.IsClientScriptIncludeRegistered(Me.GetType(), "HighLight")) Then
                Page.ClientScript.RegisterClientScriptInclude("HighLight", "/Script/HighLight.js")
            End If

            '--------------------------------------------------* 設定SqlDataSource連線及Select命令
            Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
            Me.item_defList.SelectCommand = ViewState("sSql")
            '--------------------------------------------------* 設定GridView資料來源ID
            Me.grdList.DataSourceID = Me.item_defList.ID
        End If

        '檢查是否為總部、業務處、區長、站長，列出所屬單位加油站，
        Select Case Left(sRol, 1)
            Case "2"
                '設定唯讀
                txtAreaFrom.Attributes("readonly") = "readonly"
                txtAreaFrom.CssClass = "readonly"
                txtAreaTo.Attributes("readonly") = "readonly"
                txtAreaTo.CssClass = "readonly"
                txtAreaFrom.Text = modDB.GetCodeName("ARE_ID", "SELECT ARE_ID FROM [MECHSTNM_View] where STNID='" & STNID & "'")
                txtAreaTo.Text = txtAreaFrom.Text
                AreaName.Text = modDB.GetCodeName("ARE_DESC", "SELECT ARE_DESC FROM [MECHSTNM_View] where STNID='" & STNID & "'")
                AreaName2.Text = AreaName.Text
                btnArea1.Enabled = False
                btnArea2.Enabled = False

            Case "3"
                txtAreaFrom.Attributes("readonly") = "readonly"
                txtAreaFrom.CssClass = "readonly"
                txtAreaTo.Attributes("readonly") = "readonly"
                txtAreaTo.CssClass = "readonly"
                Me.txtStnFrom.Text = STNID
                Me.txtStnTo.Text = STNID
                Me.txtStnName.Text = STNNAME
                Me.txtStnName2.Text = STNNAME


                Me.txtStnFrom.Attributes("readonly") = "readonly"
                Me.txtStnFrom.CssClass = "readonly"
                Me.txtStnTo.Attributes("readonly") = "readonly"
                Me.txtStnTo.CssClass = "readonly"

                Me.txtAreaFrom.Text = modDB.GetCodeName("ARE_ID", "SELECT ARE_ID FROM [MECHSTNM_View] where STNID='" & STNID & "'")
                Me.txtAreaTo.Text = Me.txtAreaFrom.Text
                Me.AreaName.Text = modDB.GetCodeName("ARE_DESC", "SELECT ARE_DESC FROM [MECHSTNM_View] where STNID='" & STNID & "'")
                Me.AreaName2.Text = AreaName.Text

                btnStnId.Enabled = False
                btnStnId2.Enabled = False
                btnArea1.Enabled = False
                btnArea2.Enabled = False

        End Select

        Me.txtAreaFrom.Attributes.Add("onkeypress", "return CheckKeyAZ09()")
        Me.txtAreaTo.Attributes.Add("onkeypress", "return CheckKeyAZ09()")
        Me.txtStnFrom.Attributes.Add("onkeypress", "return CheckKeyAZ09()")
        Me.txtStnTo.Attributes.Add("onkeypress", "return CheckKeyAZ09()")
        Me.txtStnFrom.Attributes.Add("onchange", "setValue('txtStnFrom','txtStnTo')")
        Me.txtAreaFrom.Attributes.Add("onchange", "setValue('txtAreaFrom','txtAreaTo')")
        'Me.codeS.Attributes.Add("onkeypress", "return CheckKeyAZ09()")
        'Me.codeE.Attributes.Add("onkeypress", "return CheckKeyAZ09()")
        'Me.codeS.Attributes.Add("onchange", "setValue('codeS','codeE')")

        '--------------------------------------------------* 設定欄位屬性
        'txtToDay.Value = Format(DateAdd(DateInterval.Day, -1, Today), "yyyy/MM/dd")
        Me.txtDtMax.Text = Format(DateAdd(DateInterval.Day, 6, Now), "yyyy/MM/dd")
        modUtil.SetDateObj(Me.DateSTextBox, False, Me.DateETextBox, False)
        modUtil.SetDateObj(Me.DateETextBox, False, Nothing, False)

        modUtil.SetDateImgObj(Me.imgDtFrom, Me.DateSTextBox, True, False, DateETextBox)
        modUtil.SetDateImgObj(Me.imgDtTo, DateETextBox, True, False, , , txtDtMax)




    End Sub

#End Region

#Region "GridView Event"
    '******************************************************************************************************
    '* GridView_DataBound => 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.DataBound
        modDB.SetGridPageNum(Me.grdList, PagerButtons.NumericFirstLast, HorizontalAlign.Left)
    End Sub

    '******************************************************************************************************
    '* GridView_RowDataBound => 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub grdList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdList.RowDataBound
        Dim sTotal As String
        Dim i As Integer
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                If Not IsDBNull(DataBinder.Eval(e.Row.DataItem, "STNID")) Then
                    sTotal = DataBinder.Eval(e.Row.DataItem, "STNID")
                    If (sTotal = "合計") Then
                        For i = 0 To 15
                            e.Row.Cells(i).BackColor = Drawing.Color.Yellow
                        Next
                        ' e.Row.BackColor = Drawing.Color.Yellow
                    Else
                        If (CType(ViewState("LineNo"), Int16) = 0) Then
                            e.Row.Attributes.Add("onmouseout", "onmouseoutColor1(this);")
                            e.Row.Attributes.Add("onmouseover", "onmouseoverColor1(this);")
                            ViewState("LineNo") = 1
                        Else
                            e.Row.Attributes.Add("onmouseout", "onmouseoutColor2(this);")
                            e.Row.Attributes.Add("onmouseover", "onmouseoverColor2(this);")
                            ViewState("LineNo") = 0
                        End If

                    End If

                End If
 

                ' e.Row.Cells(2).HorizontalAlign = HorizontalAlign.Left
        End Select
    End Sub
#End Region

#Region "Button Event"
    Public Sub txtStnFrom_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStnFrom.TextChanged
        'modCharge.GetStnName(txtStnFrom, Me.txtStnName, True)
        modCharge.GetStnName(txtStnFrom, Me.txtStnName, Me.ViewState("RolDeptID"))
        If Me.txtStnTo.Text.Trim = "" Then Me.txtStnTo.Text = Me.txtStnFrom.Text
        If txtStnName2.Text = "" And Me.txtStnTo.Text <> "" Then modCharge.GetStnName(txtStnTo, Me.txtStnName2, True)

    End Sub
    Protected Sub txtStnTo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStnTo.TextChanged
        'modCharge.GetStnName(txtStnTo, Me.txtStnName2, True)
        modCharge.GetStnName(txtStnTo, Me.txtStnName2, Me.ViewState("RolDeptID"))
    End Sub
    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStnId.Click
        ' Me.QryStn.Show()
        Me.QryStn.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"), True)

    End Sub
    Protected Sub btnStnID2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStnId2.Click
        'Me.QryStn2.Show()
        Me.QryStn2.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"), True)

    End Sub

    Public Sub txtAreaFrom_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAreaFrom.TextChanged

        Me.AreaName.Text = modDB.GetCodeName("ARE_DESC", "SELECT ARE_DESC FROM [STN_AREA] where ARE_ID='" & Me.txtAreaFrom.Text & "'")
        'modOrder.GetAreaDesc(txtAreaFrom, AreaName, Me.ViewState("RolDeptID"), True)

        If Me.txtAreaTo.Text = "" Then Me.txtAreaTo.Text = Me.txtAreaFrom.Text
        If Me.AreaName2.Text = "" And Me.txtAreaTo.Text <> "" Then Me.AreaName2.Text = modDB.GetCodeName("ARE_DESC", "SELECT ARE_DESC FROM [STN_AREA] where ARE_ID='" & Me.txtAreaTo.Text & "'")
    End Sub
    Public Sub txtAreaTo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAreaTo.TextChanged

        Me.AreaName2.Text = modDB.GetCodeName("ARE_DESC", "SELECT ARE_DESC FROM [STN_AREA] where ARE_ID='" & Me.txtAreaTo.Text & "'")
        'modOrder.GetAreaDesc(txtAreaTo, AreaName2, Me.ViewState("RolDeptID"), True)

    End Sub
    Protected Sub btnArea1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnArea1.Click
        'Me.QryArea1.Show()
        Me.QryArea1.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))

    End Sub
    Protected Sub btnArea2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnArea2.Click
        'Me.QryArea2.Show()
        Me.QryArea2.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))

    End Sub

    '******************************************************************************************************
    '* 查詢
    '******************************************************************************************************

    Protected Sub QueryButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QueryButton.Click
        Dim sTxt(8), sSql, sSqlc As String
        Dim i As Integer = 0

        Try
            '---------------------------------------------* 檢核欄位正確性
            sTxt(0) = Me.txtAreaFrom.Text.Trim
            sTxt(1) = Me.txtAreaTo.Text.Trim
            sTxt(2) = Me.txtStnFrom.Text.Trim
            sTxt(3) = Me.txtStnTo.Text.Trim
            sTxt(4) = Me.DateSTextBox.Text.Trim
            sTxt(5) = Me.DateETextBox.Text.Trim
            sTxt(6) = ""
            'sTxt(7) = Me.codeE.Text



            If (sTxt(0) > sTxt(1)) Or (sTxt(0) <> "" And sTxt(1) = "") Or (sTxt(1) <> "" And sTxt(0) = "") Then
                modUtil.showMsg(Me.Page, "訊息", "區間代號起訖值錯誤！")
                Exit Sub
            End If

            If (sTxt(2) > sTxt(3)) Or (sTxt(2) <> "" And sTxt(3) = "") Or (sTxt(3) <> "" And sTxt(2) = "") Then
                modUtil.showMsg(Me.Page, "訊息", "加油站代號起訖值錯誤！")
                Exit Sub
            End If

            If sTxt(0) <> "" And sTxt(1) <> "" And sTxt(2) <> "" And sTxt(3) <> "" Then
                sSql = "select STNID FROM MECHSTNM where STNID = '" & sTxt(2) & "' and ARE_ID between '" & sTxt(0) & "' and '" & sTxt(1) & "'"
                'Response.Write(sSql)
                If modDB.GetCodeName("STNID", sSql) = "" Then
                    modUtil.showMsg(Me.Page, "訊息", "所選取的【加油站起號】不屬於該區別!!")
                    Exit Sub
                End If
                sSql = "select STNID FROM MECHSTNM where STNID = '" & sTxt(3) & "' and ARE_ID between '" & sTxt(0) & "' and '" & sTxt(1) & "' "
                'Response.Write(sSql)
                If modDB.GetCodeName("STNID", sSql) = "" Then
                    modUtil.showMsg(Me.Page, "訊息", "所選取的【加油站迄號】不屬於該區別!!")
                    Exit Sub
                End If
            End If

            If (sTxt(4) > sTxt(5)) Or (sTxt(4) = "" And sTxt(5) = "") Then
                modUtil.showMsg(Me.Page, "訊息", "日期區間起訖值錯誤或空白！")
                Exit Sub
            End If

            sSql = ""
            sSqlc = ""
            '---------------------------------------------* 顯示查詢結果

            If sTxt(0) <> "" And sTxt(1) <> "" Then '---------------------------*區間代號有輸入的情況

                ' sSqlc = sSqlc & " And M.ARE_ID between '" & sTxt(0) & "' and '" & sTxt(1) & "'"
                sSqlc = sSqlc & " And STNID in (select STNID  from MECHSTNM where ARE_ID between '" & sTxt(0) & "' and '" & sTxt(1) & "')"

            End If

            If sTxt(2) <> "" And sTxt(3) <> "" Then '---------------------------*加油站代號有輸入的情況

                sSqlc = sSqlc & " And STNID between '" & sTxt(2) & "' and '" & sTxt(3) & "'"

            End If

            '取得氣象篩選條件CheckList
            For i = 0 To Me.WetCodeChk.Items.Count - 1
                If Me.WetCodeChk.Items(i).Selected Then
                    If sTxt(6) <> "" Then
                        sTxt(6) += ",'" & Me.WetCodeChk.Items(i).Value & "'"
                    Else
                        sTxt(6) = "'" & Me.WetCodeChk.Items(i).Value & "'"
                    End If
                End If
            Next

            If sTxt(6) <> "" Then sSqlc += " And AttCode in (" & sTxt(6) & ")"


            sSql = " select STNID,STNNAME,CutDate,WetName,AttName from StationWeather_View" _
            & " where CutDate between '" & sTxt(4) & "' and '" & sTxt(5) & "' " & sSqlc _
            & " order by STNID,CutDate"


            '--------------------------------------------------* 設定SqlDataSource連線及Select命令
            Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
            Me.item_defList.SelectCommand = sSql
            ViewState("sSql") = sSql '* 保留SQL
            'Response.Write(sSql)
            '--------------------------------------------------* 設定GridView資料來源ID
            Me.grdList.DataSourceID = Me.item_defList.ID
            Me.grdList.DataBind()

            If Me.grdList.Rows.Count = 0 Then
                modUtil.showMsg(Me.Page, "訊息", "查無資料！")
            Else
                Me.ExcelButton.Enabled = True
            End If


        Catch ex As Exception
            modUtil.showMsg(Me.Page, "訊息", "查詢時錯誤：" + ex.Message)
            'MsgBox("查詢時錯誤：" + ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.SystemModal, "錯誤訊息")
        End Try
    End Sub
    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub ClearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearButton.Click
        Server.Transfer("ATT_Q_040.aspx", False)
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************
    Protected Sub ExcelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExcelButton.Click
        Dim NowStr As String = String.Format("{0:yyyyMMddHHmmss}", DateTime.Now)
        grdList.AllowPaging = False
        grdList.AllowSorting = False
        grdList.DataBind()
        modUtil.GridView2Excel("StationWeather_" & NowStr & ".xls", Me.grdList)
        grdList.AllowPaging = True
        grdList.AllowSorting = True
        grdList.DataBind()
    End Sub


#End Region

End Class
