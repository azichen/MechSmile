'******************************************************************************************************
'* 程式：ATT_I_051 排班資料維護_新增資料用
'* 作成：MODIFY BY VOI_M_001
'* 版次：2013/11/13(VER1.00)：新開發
'******************************************************************************************************
Imports System.Data
Imports System.Data.SqlClient
Imports CartItem
Imports Cart
Imports modUnset

Partial Class ATT_I_051
    Inherits System.Web.UI.Page
    Private mycart As Cart
    'Dim INQ_ID As String
    Dim STNID As String
    Dim STNNAME As String
    Dim EMPID As String
    Dim EMPNAME As String
    Dim PERMITDATE As String


#Region "Page_Load"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sSql As String
        If Not IsPostBack Then

            '檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else
                Dim sFileName As String = Request.Path.Substring(Request.Path.LastIndexOf("/") + 1, Request.Path.Length - (Request.Path.LastIndexOf("/") + 1))
                Dim sRol As String = HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
                ViewState("ROL") = modUtil.GetRolData(sFileName, sRol)
            End If

            '1.HttpRequest 2.回傳使用權限 3.回傳使用者層級 4.回傳使用者所屬部門ID
            modUtil.GetRolData(Me.Request, Me.ViewState("ROL"), Me.ViewState("RolType"), Me.ViewState("RolDeptID"))

            Me.ViewState("QryField") = Session("QryField") '* 紀錄原查詢條件(供返回時更新)
            Session("QryField") = Nothing

            If Request("FormMode") = "add" Then
                Me.FormView1.ChangeMode(FormViewMode.Insert)
                Me.TitleLabel.Text = "排班資料維護 - 新增表單"
            Else
                Me.FormView1.ChangeMode(FormViewMode.ReadOnly)
                Me.TitleLabel.Text = "排班資料維護 - 瀏覽更新表單"
            End If

            'ViewState("TotalPrice") = 0         '訂購總金額
            ViewState("EditIndex") = Nothing     '修改明細索引值
            'Session("cart") = Nothing            '暫時明細資料 
            '
            modUtil.SetDateObj(Me.txtSchDt, False, Nothing, False)
            modUtil.SetDateImgObj(Me.imgSchDt, Me.txtSchDt, True, False)
            '
            Me.txtSTNID_i.Text = Request("STNID")
            modCharge.GetStnName(txtSTNID_i, Me.txtSTNNAME_i, Me.ViewState("RolDeptID"))

            Me.txtSchDt.Text = Format(CDate(Now()), "yyyy/MM/dd")
            '放合理工時抓取 20140205
            Me.DUTYHOUR.Text = "(" + GETDUTYHR(Me.txtSTNID_i.Text, Me.txtSchDt.Text) + ")"
        End If

        STNID = Request("STNID").Trim
        If STNID = "" Then STNID = Request.Cookies("STNID").Value
        sSql = "SELECT STNNAME  FROM [MECHSTNM] where Rtrim(STNID) = '" & STNID & "'"
        STNNAME = modDB.GetCodeName("STNNAME", sSql)
        '
        Me.txtEMPID_i.Focus()
        '
        If PERMITDATE = "" Then
            PERMITDATE = modUnset.GET_LOCKYM("", STNID) + "21"
        End If
    End Sub
#End Region

#Region "FormView Event"

    '顯示資料處理
    Protected Sub FormView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles FormView1.DataBound
        Dim row As FormViewRow = Me.FormView1.Row
        Dim rowView As Data.DataRowView = CType(FormView1.DataItem, Data.DataRowView)
        'Dim sSql As String
        Dim y As Integer = 0
        Try
            modUtil.SetObjReadOnly(Me, "txtSTNNAME_i")
            modUtil.SetObjReadOnly(Me, "txtEmpNAME_i")
            '
            'Me.txtSTNID_i.Text = Request("STNID")
            'modCharge.GetStnName(txtSTNID_i, Me.txtSTNNAME_i, Me.ViewState("RolDeptID"))

            Me.txtSTNID_i.Visible = False
            CLSTIME_Proc(Me.txtSTNID_i.Text)
            Me.txtEMPID_i.Focus()

            'Me.txtSHSTDATE_i.Text = Format(CDate(txtSchDt.Text), "yyyyMMdd")

            '明細格式設定
            modDB.SetGridViewStyle(CType(row.FindControl("GridView1"), GridView))        '套用gridview樣式
            CType(row.FindControl("GridView1"), GridView).AllowPaging = False
            modDB.SetFields("s1", "站代號", CType(row.FindControl("GridView1"), GridView), False)
            modDB.SetFields("s2", "站名稱", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s3", "員工代號", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s4", "員工姓名", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s5", "排班日期", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s6", "起始時間", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s7", "結束時間", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s8", "午休起始", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s9", "午休結束", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s10", "出勤時數", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            modDB.SetFields("s11", "結束日期", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)
            'modDB.SetFields("s12", "兩周累計時數", CType(row.FindControl("GridView1"), GridView), HorizontalAlign.Center, "", False)

            '將「結束日期」欄位隱藏
            CType(row.FindControl("GridView1"), GridView).Columns(10).Visible = False
            '增加刪除及修改鍵
            Dim delField As New CommandField()
            delField.ShowDeleteButton = True
            delField.DeleteText = "刪除"
            CType(row.FindControl("GridView1"), GridView).Columns.Add(delField)

            Dim editField As New CommandField()
            editField.EditText = "編輯"
            editField.ShowEditButton = True
            CType(row.FindControl("GridView1"), GridView).Columns.Add(editField)

            'End If

        Catch ex As Exception
            Me.msgLabel2.Text = "資料連結發生錯誤" & ex.Message
        End Try

    End Sub


#End Region

#Region "Button Event"
    '******************************************************************************************************
    '* 選擇站
    '******************************************************************************************************
    Public Sub txtSTNID_i_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        modCharge.GetStnName(txtSTNID_i, Me.txtSTNNAME_i, Me.ViewState("RolDeptID"))
        CLSTIME_Proc(Me.txtSTNID_i.Text)
    End Sub

    Protected Sub btnStn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        CType(Me.FindControl("QryStn"), Object).show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"), True)
    End Sub

    '清除重設表單
    Protected Sub ClearButton_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        mycart = New Cart
        Session("cart") = Nothing
        'Session("QryField") = Me.ViewState("QryField")
        Server.Transfer("ATT_I_051.aspx?STNID=" + Me.txtSTNID_i.Text.Trim + "&FormMode=add", False)
    End Sub

    '回主畫面
    Protected Sub Button_return_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        mycart = New Cart
        Session("cart") = Nothing
        Session("QryField") = Me.ViewState("QryField")
        Response.Redirect("ATT_Q_051.aspx")

    End Sub

    '新增送出
    Protected Sub SendButton_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sMsg As String = ""
        mycart = CType(Session("cart"), Cart)
        If mycart Is Nothing Then
            mycart = New Cart
        Else
            mycart = CType(Session("cart"), Cart)
        End If

        STNID = Me.txtSTNID_i.Text
        EMPID = Me.txtEMPID_i.Text
        'SHSTDATE = Me.txtSHSTDATE_i.Text

        If mycart.Count > 0 Then
            'If InsertSend() Then
            sMsg = InsertSend()
            'modDB.InsertSignRecord("azitest", "att_I_051: sMsg =" & sMsg, My.User.Name)
            'modDB.InsertSignRecord("azitest", "att_I_051: sMsg.sub =" & sMsg.Substring(3, sMsg.Length - 3).Trim(), My.User.Name)

            If sMsg.Substring(0, 1) = "Y" Then
                'modDB.InsertSignRecord("新增", "站代碼:" & STNID & " 員工:" & EMPID & " 排班日期:" & Me.txtSchDt.Text, My.User.Name)
                '
                Session("QryField") = Me.ViewState("QryField")
                '
                If sMsg.Substring(1, 2).Trim <> "00" Then
                    Response.Redirect("ATT_Q_051.aspx?msg=資料新增完成!!" & sMsg.Substring(3, sMsg.Length - 3).Trim & " 等共" & sMsg.Substring(1, 2) & "筆排班，超出次數不予儲存。")
                Else
                    Response.Redirect("ATT_Q_051.aspx?msg=資料新增完成!!")
                End If
                'Response.Redirect("ATT_Q_051.aspx?msg=資料新增完成!!")
            End If
        Else
            modUtil.showMsg(Me.Page, "訊息", "無任何排班資料，請建立")
        End If
    End Sub

    '新增一筆資料(新增)
    Protected Sub b1_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim sMsg As String = ""
        Dim VHOURS As Double = 0
        'modDB.InsertSignRecord("AziTest", "I_051:ok-A", My.User.Name)
        If Not CheckData_detail(sMsg) Then
            'modDB.InsertSignRecord("AziTest", "I_051:ok-B", My.User.Name)
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常：" & sMsg)
        Else
            'modDB.InsertSignRecord("AziTest", "I_051:ok-B-a", My.User.Name)
            If sMsg <> "" Then
                modUtil.showMsg(Me.Page, "訊息", "請注意：" & sMsg)
            End If

            Dim row As FormViewRow = FormView1.Row
            Dim s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11 As String
            mycart = CType(Session("cart"), Cart)
            If mycart Is Nothing Then
                mycart = New Cart
            Else
                mycart = CType(Session("cart"), Cart)
            End If
            '判斷是否為修改明細，若是則先刪除再做新增
            If Not (ViewState("EditIndex") Is Nothing) Then
                mycart.RemoveAt(ViewState("EditIndex"))
                ViewState("EditIndex") = Nothing
            End If
            '
            s1 = Me.txtSTNID_i.Text.Trim
            s2 = Me.txtSTNNAME_i.Text.Trim
            s3 = Me.txtEMPID_i.Text.Trim
            s4 = Me.txtEMPNAME_i.Text.Trim
            s5 = modUnset.PADATECHK(Me.txtSchDt.Text.Trim).Substring(1, 8)
            s6 = CType(row.FindControl("txtSHSTTIME_i"), TextBox).Text.Trim
            s7 = CType(row.FindControl("txtSHEDTIME_i"), TextBox).Text.Trim
            s8 = CType(row.FindControl("txtRTSTTIME_i"), TextBox).Text.Trim
            s9 = CType(row.FindControl("txtRTEDTIME_i"), TextBox).Text.Trim
            'modDB.InsertSignRecord("AziTest", "I_051:ok-B-b", My.User.Name)
            '20151130 新增跨中午班時，要扣一小時。(人資系統需求)
            If (s6 < "1200") And (s7 > "1300") And (CType(row.FindControl("txtRTSTTIME_i"), TextBox).Text.Trim = "") Then
                '
                Dim tt_b1 As DataTable
                Dim sSql_b1 As String
                'DEPT_ID='" & s1 & "' AND
                '20151207 麗敏:只要事站長身分就要扣1小時，支援別站也一樣
                sSql_b1 = "SELECT DEPT_ID FROM MP_HR.DBO.DEPARTMENT WHERE DEPT_MEMPLID='" & s3 & "' " _
                        & "   AND (DEPT_CLOSE_DATE IS NULL OR DEPT_CLOSE_DATE<'2000/01/01')"
                tt_b1 = get_DataTable(sSql_b1)
                'modDB.InsertSignRecord("AziTest", "I_051:s1=" & s1 & " s3=" & s3, My.User.Name)
                If tt_b1.Rows.Count > 0 Then
                    'modDB.InsertSignRecord("AziTest", "I_051: s4 IS STHDER", My.User.Name)
                    'txtRTSTTIME_i.Text = "1200"
                    'txtRTEDTIME_i.Text = "1300"
                    s8 = "1200"
                    s9 = "1300"
                End If
            End If

            '
            's8 = CType(row.FindControl("txtRTSTTIME_i"), TextBox).Text.Trim
            's9 = CType(row.FindControl("txtRTEDTIME_i"), TextBox).Text.Trim
            's10 = CType(row.FindControl("txtWORKHOUR_i"), TextBox).Text.Trim
            VHOURS = modUnset.COUNTHOURS(Me.txtSHSTTIME_i.Text, Me.txtSHEDTIME_i.Text)
            If s8 <> "" Then
                'VHOURS = VHOURS - modUnset.COUNTHOURS(Me.txtRTSTTIME_i.Text, Me.txtRTEDTIME_i.Text)
                VHOURS = VHOURS - modUnset.COUNTHOURS(s8, s9)
            End If
            s10 = VHOURS.ToString
            '
            s11 = s5
            If s6 > s7 Then
                s11 = Format(CDate(Format(CInt(s5), "0000/00/00")).AddDays(1), "yyyyMMdd")
            End If
            'modDB.InsertSignRecord("AziTest", "I_051:ok-3:s11=" & s11 & " vhours=" & VHOURS.ToString, My.User.Name)
            '
            Dim item As New CartItem(s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11) ', s12, s13, s14
            CType(mycart, IList).Add(item)
            CType(row.FindControl("GridView1"), GridView).DataSource = mycart
            CType(row.FindControl("GridView1"), GridView).DataBind()
            'modDB.InsertSignRecord("AziTest", "I_051:ok-4", My.User.Name)
            Session("cart") = mycart
            'CType(row.FindControl("txtNirw_i"), TextBox).Text = ""
            CType(row.FindControl("b1"), Button).Text = "新增一筆"
        End If
        'modDB.InsertSignRecord("AziTest", "I_051:ok-7", My.User.Name)
    End Sub

    '清除重設明細資料(新增)
    Public Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim row As FormViewRow = FormView1.Row
        'CType(row.FindControl("txtSHSTDATE_i"), TextBox).Text = ""
        CType(row.FindControl("txtSHSTTIME_i"), TextBox).Text = ""
        CType(row.FindControl("txtSHEDTIME_i"), TextBox).Text = ""
        CType(row.FindControl("txtRTSTTIME_i"), TextBox).Text = ""
        CType(row.FindControl("txtRTEDTIME_i"), TextBox).Text = ""
        ViewState("EditIndex") = Nothing
        CType(row.FindControl("b1"), Button).Text = "新增一筆"
        '
        Me.txtSTNID_i.Text = Request("STNID")
        modCharge.GetStnName(txtSTNID_i, Me.txtSTNNAME_i, Me.ViewState("RolDeptID"))
    End Sub

    '觸發商品小視窗
    Protected Sub Button_item_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        CType(FormView1.FindControl("QryItem"), Object).show()
    End Sub
    'txtEMPID_i_TextChanged
    Protected Sub txtEmpID_i_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call modCharge.GetEmpName(Me.txtEMPID_i, Me.txtEMPNAME_i, True)
    End Sub

    Protected Sub btnEmp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryEmp.Show(Me.txtSTNID_i.Text)
        'CType(Me.FindControl("QryEmp"), Object).show(Me.ViewState("RolType"), Me.txtSTNID_i.Text, True)
    End Sub

    Public Sub CLSTIME_i_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtSHSTTIME_i.Text = Me.CLSTIME.SelectedItem.Text.Substring(5, 4)
        txtSHEDTIME_i.Text = Me.CLSTIME.SelectedItem.Text.Substring(10, 4)
        If Me.CLSTIME.SelectedItem.Text.Substring(15, 1) = "Y" Then '正常班
            txtRTSTTIME_i.Text = "1200"
            txtRTEDTIME_i.Text = "1300"
        Else
            txtRTSTTIME_i.Text = ""
            txtRTEDTIME_i.Text = ""
        End If
    End Sub

    Protected Sub btnPreDate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.txtSchDt.Text = Format(CDate(txtSchDt.Text).AddDays(-1), "yyyy/MM/dd")
        Me.DUTYHOUR.Text = "(" + GETDUTYHR(Me.txtSTNID_i.Text, Me.txtSchDt.Text) + ")"
    End Sub

    Protected Sub btnPosDate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtSchDt.Text = Format(CDate(txtSchDt.Text).AddDays(1), "yyyy/MM/dd")
        Me.DUTYHOUR.Text = "(" + GETDUTYHR(Me.txtSTNID_i.Text, Me.txtSchDt.Text) + ")"
    End Sub

    Public Sub txtSchDt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'Me.txtSHSTDATE_i.Text = Format(CDate(txtSchDt.Text), "yyyyMMdd")
        Me.DUTYHOUR.Text = "(" + GETDUTYHR(Me.txtSTNID_i.Text, Me.txtSchDt.Text) + ")"
    End Sub

#End Region

#Region "add detailData"

    '刪除發票資料(新增)
    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs)
        Dim row As FormViewRow = FormView1.Row
        mycart = CType(Session("cart"), Cart)
        If mycart Is Nothing Then
            mycart = New Cart
        Else
            mycart = CType(Session("cart"), Cart)
        End If

        mycart.RemoveAt(e.RowIndex)
        CType(row.FindControl("GridView1"), GridView).DataSource = mycart
        CType(row.FindControl("GridView1"), GridView).DataBind()
        Session("cart") = mycart
        CallByName(Me.Page, "btnClear_Click", CallType.Method, Nothing, Nothing)

    End Sub

    '修改明細資料(新增)
    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs)
        Dim row As FormViewRow = FormView1.Row
        mycart = CType(Session("cart"), Cart)
        If mycart Is Nothing Then
            mycart = New Cart
        Else
            mycart = CType(Session("cart"), Cart)
        End If
        ViewState("EditIndex") = e.NewEditIndex
        'CType(row.FindControl("txtSHSTDATE_i"), TextBox).Text = mycart.GetItem(e.NewEditIndex).s5
        CType(row.FindControl("txtSHSTTIME_i"), TextBox).Text = mycart.GetItem(e.NewEditIndex).s6
        CType(row.FindControl("txtSHEDTIME_i"), TextBox).Text = mycart.GetItem(e.NewEditIndex).s7
        CType(row.FindControl("txtRTSTTIME_i"), TextBox).Text = mycart.GetItem(e.NewEditIndex).s8
        CType(row.FindControl("txtRTEDTIME_i"), TextBox).Text = mycart.GetItem(e.NewEditIndex).s9

        CType(row.FindControl("b1"), Button).Text = "送出修改"
        Session("cart") = mycart

    End Sub

#End Region

#Region "CheckData"

    Protected Sub CLSTIME_Proc(ByVal pStnid As String)
        Dim row As FormViewRow = Me.FormView1.Row
        Me.CLSTIME.Items.Clear()
        'CType(row.FindControl("CLSTIME"), DropDownList).Items.Clear()
        Dim sConn As SqlClient.SqlConnection = Nothing
        Dim sCmd As SqlClient.SqlCommand = Nothing
        Dim sSql As String

        sConn = New SqlClient.SqlConnection
        sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
        sConn.Open()

        sCmd = New SqlClient.SqlCommand()
        sCmd.Connection = sConn

        Dim sReader As SqlClient.SqlDataReader = Nothing

        '------------------------------------- 排班代碼資料 
        sSql = "SELECT CLSID,STTIME,EDTIME,NFLAG FROM CLSTIME WHERE STNID='" & pStnid & "'" _
             & " ORDER BY CLSID"

        sCmd.CommandText = sSql
        sReader = sCmd.ExecuteReader()
        Me.CLSTIME.Items.Add("")
        'CType(row.FindControl("CLSTIME"), DropDownList).Items.Add("")
        While sReader.Read
            'modDB.InsertSignRecord("AziTest", "att_i_051:ok-c:" & sReader.GetString(0) & " " & sReader.GetString(1) & " " & sReader.GetString(2) & " " & sReader.GetString(3), My.User.Name)
            If sReader.GetString(3) = "Y" Then
                Me.CLSTIME.Items.Add(FillStr(sReader.GetString(0), 4, "_") & " " & sReader.GetString(1) & "~" & sReader.GetString(2) & " " & sReader.GetString(3))
                'CType(row.FindControl("CLSTIME"), DropDownList).Items.Add(FillStr(sReader.GetString(0), 4, "_") & " " & sReader.GetString(1) & "~" & sReader.GetString(2) & " " & sReader.GetString(3))
            Else
                Me.CLSTIME.Items.Add(FillStr(sReader.GetString(0), 4, "_") & " " & sReader.GetString(1) & "~" & sReader.GetString(2) & " N")
                'CType(row.FindControl("CLSTIME"), DropDownList).Items.Add(FillStr(sReader.GetString(0), 4, "_") & " " & sReader.GetString(1) & "~" & sReader.GetString(2) & " N")
            End If
        End While
        sReader.Close()
        'CLSTIME.SelectedIndex = -1
        Me.CLSTIME.Text = ""
        'CType(row.FindControl("CLSTIME"), DropDownList).Text = ""
    End Sub


    '******************************************************************************************************
    '* 明細新增前置檢核處理
    '******************************************************************************************************
    Private Function CheckData_detail(ByRef pMsg As String) As Boolean
        Dim sValidOK As Boolean
        Dim sSql, Val, getEMPID, getSHSTDATE, getSHSTTIME, getSHEDDATE, getSHEDTIME As String ', getUseYear, getNirw, getNirs1, getNirs2 
        Dim sRow As Collection = New Collection
        Dim CHKRSTR As String = ""
        Dim tt As DataTable
        Dim EMPCARTQTY As Integer = 0
        pMsg = ""
        sValidOK = True
        Try
            'modDB.InsertSignRecord("AziTest", "I_051:ok-c", My.User.Name)
            'With FormView1
            '---------------------------------------------* 檢核欄位正確性
            '檢核站號
            Val = Me.txtSTNID_i.Text
            If Not modUtil.CheckNotEmpty(Me, "txtSTNID_i") Then
                sValidOK = False : pMsg = pMsg & "\n *必須指定站所"
                Me.txtSTNID_i.BackColor = Drawing.Color.Pink
                '----------------------------------------* 檢核格式
            ElseIf Not IsNumeric(Val) Then
                sValidOK = False : pMsg = pMsg & "\n *站代號需輸入數字"
                Me.txtSTNID_i.BackColor = Drawing.Color.Pink
            ElseIf Val.Length <> 6 Then
                sValidOK = False : pMsg = pMsg & "\n *站代號須為6碼"
                Me.txtSTNID_i.BackColor = Drawing.Color.Pink
            ElseIf txtSTNNAME_i.Text = "" Then
                sValidOK = False : pMsg = pMsg & "\n *站代碼錯誤"
                Me.txtSTNID_i.BackColor = Drawing.Color.Pink
            End If

            '檢核加油站是否已停用(以上都通過才檢測)
            'If sValidOK Then
            '    sSql = "Select STNNAME,Convert(char(10),CLOSEDATE,111) as CLOSEDATE from mechstnm where CLOSEDATE IS NOT NULL AND DATEDIFF(DAY,GETDATE(),CLOSEDATE)<0 " _
            '    & " AND STNID='" & Val & "' "
            '    sRow = modDB.GetRowData(sSql)
            '    If sRow.Count > 0 Then
            '        sValidOK = False : pMsg = pMsg & "\n * " & sRow("StnName").ToString & " 已設定關站( " _
            '        & sRow("CLOSEDATE").ToString & ")，不可編輯排班資料"
            '    End If
            'End If

            '檢核員工代號
            If sValidOK Then
                If txtEMPID_i.Text.Trim.Length <> 8 Then
                    sValidOK = False : pMsg = pMsg & "\n * 員工代號長度錯誤!"
                End If
            End If

            '檢核排班起始時間
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(CType(FormView1.FindControl("txtSHSTTIME_i"), TextBox).Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : pMsg = pMsg & "\n * 排班起始時間錯誤!請檢查!"
                Else
                    txtSHSTTIME_i.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '檢核排班終止時間
            If sValidOK Then
                'CHKRSTR = modUnset.PADATECHK(txtSHEDTIME_i.Text)
                CHKRSTR = modUnset.PATIMECHK(CType(FormView1.FindControl("txtSHEDTIME_i"), TextBox).Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : pMsg = pMsg & "\n * 排班迄止時間錯誤!請檢查!"
                Else
                    txtSHEDTIME_i.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            If sValidOK Then
                If txtSHSTTIME_i.Text = txtSHEDTIME_i.Text Then
                    sValidOK = False : pMsg = pMsg & "\n * 排班起迄時間錯誤!請檢查!"
                End If
            End If

            '檢核午休起始時間
            If sValidOK And (CType(FormView1.FindControl("txtRTSTTIME_i"), TextBox).Text <> "") Then
                CHKRSTR = modUnset.PATIMECHK(CType(FormView1.FindControl("txtRTSTTIME_i"), TextBox).Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : pMsg = pMsg & "\n * 午休起始時間錯誤!請檢查!"
                Else
                    txtRTSTTIME_i.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '檢核午休終止時間
            If sValidOK And (CType(FormView1.FindControl("txtRTEDTIME_i"), TextBox).Text <> "") Then
                'CHKRSTR = modUnset.PADATECHK(txtSHEDTIME_i.Text)
                CHKRSTR = modUnset.PATIMECHK(CType(FormView1.FindControl("txtRTEDTIME_i"), TextBox).Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : pMsg = pMsg & "\n * 午休迄止時間錯誤!請檢查!"
                Else
                    txtRTEDTIME_i.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            getEMPID = txtEMPID_i.Text.Trim
            '檢核日期

            'getSHSTDATE = CType(FormView1.FindControl("txtSHSTDATE_i"), TextBox).Text
            getSHSTDATE = modUnset.PADATECHK(Me.txtSchDt.Text).Substring(1, 8)
            getSHSTTIME = CType(FormView1.FindControl("txtSHSTTIME_i"), TextBox).Text
            getSHEDTIME = CType(FormView1.FindControl("txtSHEDTIME_i"), TextBox).Text
            getSHEDDATE = getSHSTDATE
            '
            If getSHSTTIME > getSHEDTIME Then
                getSHEDDATE = Format(CDate(Format(CInt(getSHSTDATE), "0000/00/00")).AddDays(1), "yyyyMMdd")
            End If

            '午休時間邏輯檢查
            If sValidOK And ((CType(FormView1.FindControl("txtRTSTTIME_i"), TextBox).Text <> "") Or (CType(FormView1.FindControl("txtRTEDTIME_i"), TextBox).Text <> "")) Then
                '20150506 增加午休時間齊全檢查
                If (txtRTSTTIME_i.Text.Trim = "") Or (txtRTEDTIME_i.Text.Trim = "") Then
                    sValidOK = False : pMsg = pMsg & "\n * 午休起迄時間錯誤!請檢查!"
                Else
                    Dim RTDATEST, RTDATEED As String
                    RTDATEST = getSHSTDATE
                    RTDATEED = getSHSTDATE
                    If ((txtRTEDTIME_i.Text < txtRTSTTIME_i.Text) And (txtRTEDTIME_i.Text <> "0000")) Then '休息跨日
                        sValidOK = False : pMsg = pMsg & "\n * 休息迄止時間錯誤!不可跨日!請檢查!"
                        'modDB.InsertSignRecord("AziTest", "ATT_I_051_OK-2", My.User.Name) 'AZITEST
                    Else
                        If txtSHSTTIME_i.Text > txtRTSTTIME_i.Text Then
                            RTDATEST = Format(CDate(Format(CInt(RTDATEST), "0000/00/00")).AddDays(1), "yyyyMMdd")
                            RTDATEED = RTDATEST
                        End If

                        If txtRTEDTIME_i.Text < txtRTSTTIME_i.Text Then
                            RTDATEED = Format(CDate(Format(CInt(RTDATEST), "0000/00/00")).AddDays(1), "yyyyMMdd")
                        End If

                        'modDB.InsertSignRecord("AziTest", "I_051:RTDTST=" & RTDATEST & " " & txtRTSTTIME_i.Text & " shstdt=" & getSHSTDATE + getSHSTTIME & " RTDTed=" & RTDATEST & " " & txtRTEDTIME_i.Text & " sheddt=" & getSHEDDATE + getSHEDTIME, My.User.Name)
                        '
                        If ((RTDATEST + txtRTSTTIME_i.Text) < (getSHSTDATE + getSHSTTIME)) Or ((RTDATEED + txtRTEDTIME_i.Text) > (getSHEDDATE + getSHEDTIME)) Then
                            sValidOK = False : pMsg = pMsg & "\n * 午休起迄時間錯誤!超出排班時間範圍!請檢查!"
                            'modDB.InsertSignRecord("AziTest", "ATT_I_051_SH:" + (getSHSTDATE + getSHSTTIME) + "~" + (getSHEDDATE + getSHEDTIME), My.User.Name) 'AZITEST
                            'modDB.InsertSignRecord("AziTest", "ATT_I_051_RT:" + (RTDATEST + txtRTSTTIME_i.Text) + "~" + (RTDATEED + txtRTEDTIME_i.Text), My.User.Name) 'AZITEST
                        ElseIf (RTDATEST + txtRTSTTIME_i.Text) >= (RTDATEED + txtRTEDTIME_i.Text) Then
                            sValidOK = False : pMsg = pMsg & "\n * 午休起迄時間錯誤!請檢查!"
                        End If
                    End If
                End If
                '
            End If

                '鎖檔日期檢核
                If sValidOK Then
                    'Dim LCKDT As String = GET_LOCKYM("1", Me.txtSTNID_i.Text) & "20"
                    Dim LCKDT As String = GET_LOCKDT(Me.txtSTNID_i.Text)
                    'If (getSHSTDATE + getSHSTTIME) <= (LCKDT + "2300") Then
                    If (getSHSTDATE <= LCKDT) Then
                        sValidOK = False : pMsg = pMsg & "\n * 資料週期已鎖或過帳，不可新增此日期之排班資料!"
                    End If
                End If
                Dim WDateSt As String = ""
                Dim WDateEn As String = ""
                Dim CHOURCHECK As Boolean = True
                Dim TWKCHK As Boolean = True
                Dim TWKCHKSTR As String = ""
                Dim EMPLCOLL As String = ""
                Dim EMPLARV, EMPLLEV As String
                Dim WORKDAY, WORKHOUR As Integer
                'Dim sConn As SqlClient.SqlConnection = Nothing
                'Dim sCmd As SqlClient.SqlCommand = Nothing
                'Dim sReader As SqlClient.SqlDataReader = Nothing
                Dim VHOUR As Double = modUnset.COUNTHOURS(getSHSTTIME, getSHEDTIME)
                '
                If (txtRTSTTIME_i.Text.Trim <> "") Then
                    VHOUR = VHOUR - modUnset.COUNTHOURS(txtRTSTTIME_i.Text.Trim, txtRTEDTIME_i.Text.Trim)
                End If
                '
                '此處增加資料庫排班資料是否重複
                If sValidOK Then
                    'modDB.InsertSignRecord("AziTest", "I_051:ok-1-1", My.User.Name)
                    'Dim sSql As String
                    'sConn = New SqlClient.SqlConnection
                    'sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
                    'sConn.Open()

                    'sCmd = New SqlClient.SqlCommand()
                    'sCmd.Connection = sConn

                    'sCmd = New SqlClient.SqlCommand()
                'sCmd.Connection = sConn

                sSql = "SELECT STNID,SHSTDATE,SHSTTIME,SHEDTIME FROM SCHEDM " _
                     & " WHERE EMPID='" & Me.txtEMPID_i.Text & "'" _
                     & "   AND ((SHEDDATE+SHEDTIME <='" & getSHEDDATE & getSHEDTIME & "'" _
                     & "     AND SHEDDATE+SHEDTIME > '" & getSHSTDATE & getSHSTTIME & "') " _
                     & "    OR  (SHSTDATE+SHSTTIME < '" & getSHEDDATE & getSHEDTIME & "'" _
                     & "     AND SHSTDATE+SHSTTIME >='" & getSHSTDATE & getSHSTTIME & "') " _
                     & "    OR  (SHSTDATE+SHSTTIME <='" & getSHEDDATE & getSHEDTIME & "'" _
                     & "     AND SHEDDATE+SHEDTIME > '" & getSHSTDATE & getSHSTTIME & "'))"

                'LOGSQL(sSql)

                    'sCmd.CommandText = sSql
                    'sReader = sCmd.ExecuteReader()
                    'sReader.Read()

                    'If (sReader.HasRows) Then
                    'modDB.InsertSignRecord("AziTest", "I_051:ok-1-2", My.User.Name)
                    tt = get_DataTable(sSql)
                    If tt.Rows.Count > 0 Then
                        'sValidOK = False
                        sValidOK = False : pMsg = pMsg & "\n *已有資料重複:(排班單位:" & tt.Rows(0).Item("STNID") & " 日期:" & tt.Rows(0).Item("SHSTDATE") & " 時間: " & tt.Rows(0).Item("SHSTTIME") & " ~ " & tt.Rows(0).Item("SHEDTIME") & " )"
                    End If
                    'modDB.InsertSignRecord("AziTest", "I_051:ok-1-3", My.User.Name)

                    'sReader.Close()
                    'sConn.Close()
                    'sCmd.Dispose() '手動釋放資源
                    'sConn.Dispose()
                    'sConn = Nothing '移除指標
                    ''
                    'sConn = PA_GetSqlConnection()
                    'sCmd = New SqlClient.SqlCommand()
                    'sCmd.Connection = sConn

                    '20140403 增加建教生與工作時數檢查
                    'getSHEDDATE 取兩周日期區間
                    Dim SHEDDATE As String = getSHEDDATE.Substring(0, 4) + "/" + getSHEDDATE.Substring(4, 2) + "/" + getSHEDDATE.Substring(6, 2)
                    sSql = "SELECT CONVERT(VARCHAR(8),Convert(datetime,FWDS_DATEST, 111),112) AS FWDS_DATEST" _
                         & "      ,CONVERT(VARCHAR(8),Convert(datetime,FWDS_DATEEN, 111),112) AS FWDS_DATEEN FROM FWDUTYSH " _
                         & " WHERE '" & SHEDDATE & "' BETWEEN FWDS_DATEST AND FWDS_DATEEN "
                    tt = get_DataTable(sSql)
                    'sCmd.CommandText = sSql
                    'modDB.InsertSignRecord("AziTest", "I_051:SHEDDATE=" & SHEDDATE, My.User.Name)

                    'sReader = sCmd.ExecuteReader()
                    'If sReader.Read() Then
                    If tt.Rows.Count > 0 Then
                        WDateSt = tt.Rows(0).Item("FWDS_DATEST")
                        WDateEn = tt.Rows(0).Item("FWDS_DATEEN")
                    Else
                        sValidOK = False : pMsg = pMsg & "\n 日期:" & getSHEDDATE & "之雙週週期未設定!"
                    End If
                '
                    If sValidOK Then
                        'sReader.Close()
                        '20150616 修改以SHEDDATE 作判斷
                        sSql = "SELECT EMPL_COLL_RELATION,ISNULL(EMPL_ARV_DATE,'') AS EMPLARVDATE,ISNULL(EMPL_LEV_DATE,'') AS EMPLLEVDATE" _
                             & "      ,ISNULL(SUM(WORKHOUR),0) AS WORKHOUR FROM MP_HR.DBO.EMPLOYEE A " _
                             & "  LEFT JOIN SCHEDM B ON B.EMPID='" & Me.txtEMPID_i.Text & "' AND SHEDDATE>='" & WDateSt & "' AND SHEDDATE<='" & WDateEn & "'" _
                             & " WHERE EMPL_ID='" & Me.txtEMPID_i.Text & "'" _
                             & " GROUP BY EMPL_COLL_RELATION,EMPL_ARV_DATE,EMPL_LEV_DATE "
                    '
                        tt = get_DataTable(sSql)
                        EMPLCOLL = tt.Rows(0).Item("EMPL_COLL_RELATION")
                        EMPLARV = Format(tt.Rows(0).Item("EMPLARVDATE"), "yyyyMMdd")  'Format(sReader.GetDateTime(1), "yyyyMMdd")
                        EMPLLEV = Format(tt.Rows(0).Item("EMPLLEVDATE"), "yyyyMMdd")  'Format(sReader.GetDateTime(2), "yyyyMMdd")
                        WORKHOUR = tt.Rows(0).Item("WORKHOUR") 'sReader.GetDouble(3) 
                    '
                        If (getSHEDDATE < EMPLARV) Or ((EMPLLEV > "19800101") And (getSHEDDATE > EMPLLEV)) Then '檢查離職日
                            sValidOK = False : pMsg = pMsg & "\n 該員本日期不在職!"
                        End If
                        '
                        If EMPLCOLL = "1" Then '建教生
                            '
                            If sValidOK Then
                                If (getSHSTTIME < "0600") Or (getSHEDTIME > "2200") Or (getSHSTDATE <> getSHEDDATE) Then
                                    sValidOK = False : pMsg = pMsg & "\n 建教生排班時間限制在 06:00 ~ 22:00 之間"
                                End If
                            End If
                        End If
                    End If
                End If

                'modDB.InsertSignRecord("AziTest", " WDateSt=" + WDateSt + " WDateEn=" + WDateEn, My.User.Name) 'AZITEST
                '檢核是否與目前表單Cart中的資料重疊(以上都通過才檢測)
            If sValidOK And ViewState("EditIndex") Is Nothing Then

                mycart = CType(Session("cart"), Cart)
                If mycart Is Nothing Then
                    mycart = New Cart
                Else
                    mycart = CType(Session("cart"), Cart)
                End If

                Dim item As CartItem
                EMPCARTQTY = 0
                For Each item In mycart
                    If item.s3 = getEMPID Then
                        WORKDAY = WORKDAY + 1
                        WORKHOUR = WORKHOUR + item.s10
                    End If
                    '檢測本次新增字碼起迄是否與表單Cart中字碼起迄重疊
                    If ((item.s3 = getEMPID) And (item.s11 + item.s7) <= (getSHEDDATE & getSHEDTIME) And (item.s11 + item.s7) > (getSHSTDATE & getSHSTTIME)) Or _
                       ((item.s3 = getEMPID) And (item.s5 + item.s6) < (getSHEDDATE & getSHEDTIME) And (item.s5 + item.s6) >= (getSHSTDATE & getSHSTTIME)) Or _
                       ((item.s3 = getEMPID) And (item.s5 + item.s6) <= (getSHEDDATE & getSHEDTIME) And (item.s11 + item.s7) > (getSHSTDATE & getSHSTTIME)) Then
                        sValidOK = False : pMsg = pMsg & "\n 資料重疊(日期:" & item.s5 & " 起始時間:" & item.s6 & ")"
                        Exit For
                    End If
                    '20151221 增加CART數量檢查
                    If (item.s3 = getEMPID) Then
                        EMPCARTQTY = EMPCARTQTY + 1
                    End If
                Next
            End If
            '
            If sValidOK Then

                sSql = "EXEC TWKSCHK '" & Me.txtEMPID_i.Text & "','" & getSHEDDATE & "','" & getSHSTDATE & "'"

                TWKCHKSTR = get_DataTable(sSql).Rows(0).Item(0)

                'modDB.InsertSignRecord("ATT_I_051", Me.txtEMPID_i.Text & " " & getSHEDDATE & " " & getSHSTDATE, My.User.Name)

                If sValidOK Then
                    'If getSHEDDATE >= "2015/07/05" Then '20150702 7/4之後的排班才檢查
                    If TWKCHKSTR.Substring(0, 1) = "N" Then
                        sValidOK = False
                        pMsg = pMsg & "\n " & TWKCHKSTR.Substring(1, 8) & " ~ " & TWKCHKSTR.Substring(9, 8) & " 期間已排班" & TWKCHKSTR.Substring(17, 2) & "天"
                    Else 'ADD IN 20151221
                        'MARK IN 20151222 同天多筆排班的話就會有問題
                        'If Convert.ToInt16(TWKCHKSTR.Substring(17, 2)) + EMPCARTQTY >= 12 Then
                        '    sValidOK = False
                        '    pMsg = pMsg & "\n 該員 " & " 期間排班已12筆以上，請先存檔再繼續新增資料。"
                        'End If
                    End If
                    'End If
                End If
                '
                If sValidOK Then
                    If VHOUR > 12 Then
                        sValidOK = False
                        pMsg = pMsg & "\n 本次排班超過 12 小時!"
                    End If
                End If
                '
                If sValidOK Then
                    If (WORKHOUR + VHOUR) > 84 Then
                        'sValidOK = False : pMsg = pMsg & "\n 建教生 " & WDateSt & " ~ " & WDateEn & " 期間不可排班超過 84小時"
                        pMsg = pMsg & "\n " & WDateSt & " ~ " & WDateEn & " 期間不該排班超過84小時,目前已累計" & (WORKHOUR + VHOUR).ToString & "小時"
                    End If
                End If

                'If EMPLCOLL = "1" Then '建教生
                '    If CHOURCHECK Then '前28天才作天數與總時數的檢查
                '        If sValidOK Then
                '            If (WORKDAY + 1) > 12 Then
                '                sValidOK = False : pMsg = pMsg & "\n 建教生 " & WDateSt & " ~ " & WDateEn & " 期間不可排班超過12天"
                '            End If
                '        End If
                '        '
                '        If sValidOK Then
                '            If (WORKHOUR + VHOUR) > 84 Then
                '                'sValidOK = False : pMsg = pMsg & "\n 建教生 " & WDateSt & " ~ " & WDateEn & " 期間不可排班超過 84小時"
                '                pMsg = pMsg & "\n 建教生 " & WDateSt & " ~ " & WDateEn & " 期間不該排班超過84小時,目前已累計" & (WORKHOUR + VHOUR).ToString & "小時"
                '            End If
                '        End If
                '    End If
                '    '
                '    If sValidOK Then
                '        If VHOUR > 12 Then
                '            sValidOK = False : pMsg = pMsg & "\n 建教生 不可排班超過 12 小時"
                '        End If
                '    End If
                'Else

                '    If CHOURCHECK Then '前28天才作天數與總時數的檢查
                '        If sValidOK Then
                '            'If (WORKDAY + 1) > 12 Then
                '            If TWKCHKSTR.Substring(0, 1) = "N" Then
                '                sValidOK = False
                '                pMsg = pMsg & "\n " & TWKCHKSTR.Substring(1, 8) & " ~ " & TWKCHKSTR.Substring(9, 8) & " 期間已排班" & TWKCHKSTR.Substring(17, 2) & "天"
                '            End If
                '        End If
                '        '
                '        If sValidOK Then
                '            If (WORKHOUR + VHOUR) > 84 Then
                '                'sValidOK = False
                '                pMsg = pMsg & "\n " & WDateSt & " ~ " & WDateEn & " 期間已排班超過84小時,目前已累計" & (WORKHOUR + VHOUR).ToString & "小時"
                '            End If
                '        End If
                '    End If
                '    '
                '    If sValidOK Then
                '        If VHOUR > 12 Then
                '            sValidOK = False
                '            pMsg = pMsg & "\n 本次排班超過 12 小時!"
                '        End If
                '    End If
                'End If
            End If
            '
        Catch ex As Exception
            sValidOK = False
        End Try
        Return sValidOK
    End Function

    '******************************************************************************************************
    '* 明細新增前置檢核處理
    '******************************************************************************************************
    Private Function GETDUTYHR(ByVal pStnid As String, ByVal pBDATE As String) As String
        Dim strDutyhour As String = ""
        Dim sSql As String = ""
        Dim sConn As SqlClient.SqlConnection = Nothing
        Dim sCmd As SqlClient.SqlCommand = Nothing
        'Dim sSql As String

        sConn = New SqlClient.SqlConnection
        sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("Smile_HQConnectionString").ToString
        sConn.Open()

        sCmd = New SqlClient.SqlCommand()
        sCmd.Connection = sConn

        Dim sReader As SqlClient.SqlDataReader = Nothing
        sCmd = New SqlClient.SqlCommand()
        sCmd.Connection = sConn

        sSql = "SELECT A.STNID,CASE ATTCODE WHEN '1' THEN '晴' WHEN '2' THEN '陰' WHEN '3' THEN '雨' ELSE '無' END AS ATTNAME " _
             & "      ,CASE WHEN ATTCODE='1' THEN CASE WHEN HOLIDAY='Y' THEN HR12 ELSE HR11 END " _
             & "            WHEN ATTCODE='2' THEN CASE WHEN HOLIDAY='Y' THEN HR22 ELSE HR21 END " _
             & "            WHEN ATTCODE='3' THEN CASE WHEN HOLIDAY='Y' THEN HR32 ELSE HR31 END " _
             & "            ELSE 0 END AS STNDAYHR " _
             & "  FROM Duty_STNREAL A " _
             & "  LEFT JOIN StationWeather_View B ON B.STNID='" & pStnid & "' AND Cutdate='" & pBDATE & "'" _
             & "  LEFT JOIN Duty_HOLIDAY C ON C.HDATE='" & pBDATE & "'" _
             & " WHERE A.STNID='" & pStnid & "'"

        sCmd.CommandText = sSql
        sReader = sCmd.ExecuteReader()
        sReader.Read()

        'modDB.InsertSignRecord("AziTest", "pstnid=" + pStnid + " pbdate=" + pBDATE + " get(2)=" + sReader.GetInt32(2).ToString, My.User.Name)
        If (sReader.HasRows) Then
            If sReader.GetInt32(2) > 0 Then
                strDutyhour = "天氣:" + sReader.GetString(1) + ",合理工時:" + sReader.GetInt32(2).ToString + " 小時"
            Else
                strDutyhour = "天氣工時無資料"
            End If
        Else
            strDutyhour = "天氣工時無資料"
        End If
        'modDB.InsertSignRecord("AziTest", "strDutyhour=" + strDutyhour, My.User.Name) ' azitest

        Return strDutyhour
    End Function

#End Region

#Region "InsertSend"

    '新增資料記錄
    'Protected Function InsertSend() As Boolean
    Protected Function InsertSend() As String
        Dim sSql, TWKCHKSTR As String
        Dim vINDATE, vINTIME As String 'VSHEDDATE
        'Dim connection As New SqlConnection(Web.Configuration.WebConfigurationManager.ConnectionStrings("MechPAConnectionString").ToString)
        Dim ErrMsg As String = ""
        Dim Errqty As Integer = 0
        'connection.Open()
        'Dim command As SqlCommand = connection.CreateCommand()
        'Dim transaction As SqlTransaction

        'transaction = connection.BeginTransaction("SampleTransaction")
        'command.Connection = connection
        'command.Transaction = transaction

        'getSTNID = Me.txtSTNID_i.Text.Trim
        'getEMPID = Me.txtEMPID_i.Text.Trim

        Try
            mycart = CType(Session("cart"), Cart)
            If mycart Is Nothing Then
                mycart = New Cart
            Else
                mycart = CType(Session("cart"), Cart)
            End If
            'modDB.InsertSignRecord("AziTest", "I_051-7", My.User.Name)

            Dim num As Integer
            Dim item As CartItem
            For Each item In mycart
                'AZITEMP
                num += 1
                sSql = "EXEC TWKSCHK '" & Me.txtEMPID_i.Text & "','" & item.s11 & "','" & item.s5 & "'"
                'modDB.InsertSignRecord("AziTest", "att_i_051:" + sSql, My.User.Name)
                TWKCHKSTR = get_DataTable(sSql).Rows(0).Item(0)
                '
                If TWKCHKSTR.Substring(0, 1) = "Y" Then
                    'modDB.InsertSignRecord("AziTest", "att_i_051: insert data" & "站:" & item.s1 & " 員工:" & item.s3 & " 日時:" & item.s5 & " " & item.s6, My.User.Name)
                    vINDATE = Format(Now, "yyyyMMdd")
                    vINTIME = Format(Now, "HHmmss")
                    'command.CommandText = _
                    '" INSERT INTO SCHEDM (STNID,EMPID,SHSTDATE,SHSTTIME,SHEDDATE,SHEDTIME,CLASSID,WORKHOUR,RtSTTIME,RtEDTIME,FACFLAG,FACWH,NITEWH,inuser,indate,intime) VALUES " _
                    '& "('" & item.s1 & "','" & item.s3 & "','" & item.s5 & "','" & item.s6 & "','" & item.s11 & "','" & item.s7 & "',''," & item.s10 & ",'" & item.s8 & "','" & item.s9 & "','N',0,0,'" & My.User.Name & "','" & vINDATE & "','" & vINTIME & "')"
                    'command.ExecuteNonQuery()
                    sSql = " INSERT INTO SCHEDM (STNID,EMPID,SHSTDATE,SHSTTIME,SHEDDATE,SHEDTIME,CLASSID,WORKHOUR,RtSTTIME,RtEDTIME,FACFLAG,FACWH,NITEWH,inuser,indate,intime) VALUES " _
                         & "('" & item.s1 & "','" & item.s3 & "','" & item.s5 & "','" & item.s6 & "','" & item.s11 & "','" & item.s7 & "',''," & item.s10 & ",'" & item.s8 & "','" & item.s9 & "','N',0,0,'" & My.User.Name & "','" & vINDATE & "','" & vINTIME & "')"
                    EXE_SQL(sSql)
                    'modDB.InsertSignRecord("ATT_I_051", "排班新增 站:" & item.s1 & " 員工:" & item.s3 & " 日時:" & item.s5 & " " & item.s6, My.User.Name)
                Else
                    Errqty = Errqty + 1
                    'modDB.InsertSignRecord("AziTest", "att_i_051: err:" & "站:" & item.s1 & " 員工:" & item.s3 & " 日時:" & item.s5 & " " & item.s6, My.User.Name)
                    If ErrMsg = "" Then '先只存一筆資料作訊息提醒
                        ErrMsg = ErrMsg + " 員工:" & item.s3 & " 日時:" & item.s5 & " " & item.s6
                    End If
                End If
                'Response.Write(command.CommandText & "<br>")
            Next
            'transaction.Commit()
            Session("cart") = Nothing
            Return "Y" + Errqty.ToString("00") + ErrMsg ' + "TEST123456"

        Catch ex As Exception
            'transaction.Rollback()
            Me.msgLabel2.Text = "新增失敗，訊息：" & ex.Message
            Return "N"
        Finally
            'connection.Close()
            'command.Dispose()
            'transaction.Dispose()
            'connection.Dispose()
        End Try

    End Function

    Protected Sub LOGSQL(ByVal datastr As String)
        Dim TMPSTR As String = datastr
        While (TMPSTR.Length) > 0
            If TMPSTR.Length > 95 Then
                modDB.InsertSignRecord("AziTest", TMPSTR.Substring(0, 95), My.User.Name) 'AZITEST
                TMPSTR = TMPSTR.Substring(95, TMPSTR.Length - 95)
            Else
                modDB.InsertSignRecord("AziTest", TMPSTR.Substring(0, TMPSTR.Length - 1), My.User.Name) 'AZITEST
                TMPSTR = ""
            End If
        End While
    End Sub

#End Region

    '執行quary不回傳值
    Public Function EXE_SQL(ByVal TMPSQL1 As String) As Boolean
        'Dim AA As Integer
        Dim CONNSTR1 As String = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
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

    '連接資料庫取得datatable
    Public Function get_DataTable(ByVal SQL1 As String) As System.Data.DataTable
        Dim CONNSTR1 As String = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
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
End Class
