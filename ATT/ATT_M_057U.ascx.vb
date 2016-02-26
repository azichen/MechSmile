'******************************************************************************************************
'* 程式：ATT_M_057U 資料維護
'* //版次：2013/11/01(VER1.01)：新開發
'******************************************************************************************************

Imports modUtil
Imports modUnset
Imports System.Data


Partial Class ATT_M_057U
    Inherits System.Web.UI.UserControl

    'Private cWriteLog As Boolean = False '* 上線時須設成 False
    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Public Sub InitScreen(ByVal pMode As FormViewMode)
        Me.btnStnID.Visible = False '暫不開放
        'Me.btnCLSID.Visible = False '暫不開放
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label9.Visible = False
        txtRefStrA.Visible = False
        txtRefStrB.Visible = False
        txtRefStrC.Visible = False
        txtRefStrD.Visible = False
        txtRefSTRE.Visible = False
        TxtRefSTRA1.Visible = False
        '
        Label11.Visible = False
        txtDtStart.Visible = False
        imgDtStart.Visible = False
        Label10.Visible = False
        txtDtEnd.Visible = False
        imgDtEnd.Visible = False
        '
        'UpdatePanel1.Visible = False
        If pMode = FormViewMode.ReadOnly Then
            SetAllReadOnly(Me)
            Me.TxtSERID.Visible = True
            Me.btnStnID.Enabled = False
            Me.btnEmp.Visible = False
            Me.txtFORMid.Visible = False
            Me.TxtFORMNAME.Visible = True
            '
            If (Request("FORMID").Trim = "068") Or (Request("FORMID").Trim = "069") Then
                Label1.Visible = True
                Label2.Visible = True
                Label4.Visible = True
                Label5.Visible = True
                Label7.Visible = True
                txtRefStrA.Visible = True
                txtRefStrD.Visible = True
            ElseIf (Request("FORMID").Trim = "065") Then '薪資單
                Label11.Visible = True
                txtDtStart.Visible = True
                imgDtStart.Visible = True
                Label10.Visible = True
                txtDtEnd.Visible = True
                imgDtEnd.Visible = True
                '
                Label5.Text = "用途:"
                Label1.Visible = True
                Label5.Visible = True
                TxtRefSTRA1.Visible = True
                'UpdatePanel1.Visible = True
            ElseIf (Request("FORMID").Trim = "070") Or (Request("FORMID").Trim = "071") Then '扣繳憑單
                Label5.Text = "年度"
                Label1.Visible = True
                Label5.Visible = True
                txtRefStrA.Visible = True
                If (Request("FORMID").Trim = "070") Then
                    Label6.Text = "用途:"
                    Label6.Visible = True
                    TxtRefSTRA1.Visible = True
                End If
            ElseIf (Request("FORMID").Trim = "066") Or (Request("FORMID").Trim = "067") Then '服務、離職證明名單
                Label5.Text = "用途"
                Label1.Visible = True
                Label5.Visible = True
                TxtRefSTRA1.Visible = True
            End If
            '
        Else
            Dim sConn As SqlClient.SqlConnection = Nothing
            Dim sCmd As SqlClient.SqlCommand = Nothing
            Dim sSql As String

            sConn = New SqlClient.SqlConnection
            sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
            sConn.Open()

            sCmd = New SqlClient.SqlCommand()
            sCmd.Connection = sConn

            Dim sReader As SqlClient.SqlDataReader = Nothing
            '
            SetObjReadOnly(Me, "TxtSERID")
            SetObjReadOnly(Me, "txtStnID")
            SetObjReadOnly(Me, "txtStnName")
            SetObjReadOnly(Me, "txtEmpName")
            SetObjReadOnly(Me, "txtContacts")

            modUtil.SetDateObj(Me.txtDtStart, False, Nothing, False)
            modUtil.SetDateObj(Me.txtDtEnd, False, Nothing, False)
            modUtil.SetDateImgObj(Me.imgDtStart, Me.txtDtStart, True, True, Me.txtDtEnd)
            modUtil.SetDateImgObj(Me.imgDtEnd, Me.txtDtEnd, True, True)

            Me.txtFORMid.Visible = True
            Me.TxtFORMNAME.Visible = False
            'SetObjReadOnly(Me, "txtHAPFORM")
            Me.TxtSERID.Visible = False
            Me.txtEmpID.Attributes.Add("onkeypress", "return CheckKeyAZ09()")
            If pMode = FormViewMode.Insert Then
                '只考慮新增模式()
                Me.txtSTNID.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                Me.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                modUtil.SetObjFocus(Me.Page, Me.txtEmpID)
                '
                Lbl_STATUS.Visible = False
                TxtStatus.Visible = False
                Label8.Visible = False
            End If
            'modDB.InsertSignRecord("azitest", "hapformname-2=" + TxtFORMNAME.Text, My.User.Name)
            '
            sSql = "SELECT EMPL_NAME FROM [MP_HR].DBO.DEPARTMENT A INNER JOIN [MP_HR].DBO.EMPLOYEE B ON A.DEPT_MEMPLID=B.EMPL_ID" _
                 + " Where Dept_id='" + HttpUtility.UrlDecode(Request.Cookies("STNID").Value) + "'"

            'modDB.InsertSignRecord("Azitest", "SQL=" + sSql, My.User.Name)

            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()

            'modDB.InsertSignRecord("Azitest", "SQL=" + sSql, My.User.Name)
            sReader.Read()
            TxtContacts.Text = sReader.GetString(0)

            sReader.Close()
            txtFORMid.Items.Clear()

            'sCmd = New SqlClient.SqlCommand()
            'sCmd.Connection = sConn
            '------------------------------------- 表單代碼資料 
            sSql = "SELECT FORM_ID,FORM_NAME FROM HAFORM " _
                 & " ORDER BY FORM_ID"

            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()
            txtFORMid.Items.Add("")
            While sReader.Read
                txtFORMid.Items.Add(Format(Convert.ToInt32(sReader.GetString(0)), "000") & " " & sReader.GetString(1))
            End While
            sReader.Close()
            txtFORMid.Text = ""
            '
            txtRefSTRE.Items.Add("")
            txtRefSTRE.Items.Add("0 本人")
            txtRefSTRE.Items.Add("1 配偶")
            txtRefSTRE.Items.Add("2 父母")
            txtRefSTRE.Items.Add("3 子女")
            txtRefSTRE.Items.Add("4 祖父母")
            txtRefSTRE.Items.Add("5 孫子女")
            txtRefSTRE.Items.Add("6 外祖父母")
            txtRefSTRE.Items.Add("7 外孫子女")
            txtRefSTRE.Items.Add("8 曾祖父母")
            txtRefSTRE.Items.Add("9 外曾祖父母")
            txtRefSTRE.Text = ""
            '
            'modDB.InsertSignRecord("azitest", "hapformname-3=" + TxtFORMNAME.Text, My.User.Name)
        End If
        'modDB.InsertSignRecord("azitest", "hapformname-4=" + TxtFORMNAME.Text, My.User.Name)
    End Sub

    '******************************************************************************************************
    '* DB 更新前置檢核處理
    '******************************************************************************************************
    Private Function CheckData(ByVal pMode As FormViewMode) As Boolean
        Dim sValidOK As Boolean = True
        Dim sMsg As String = ""
        Dim CHKRSTR As String = ""

        Try
            
            If (pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit) Then '* 新增模式
                sValidOK = sValidOK And CheckLength(Me, "txtEMPID", 8, sMsg, "員工編號")

                sValidOK = sValidOK And CheckNotEmpty(Me, "txtStnid", sMsg, "申請站別")
                sValidOK = sValidOK And CheckNotEmpty(Me, "txtEMPID", sMsg, "申請員工")
                sValidOK = sValidOK And CheckNotEmpty(Me, "txtFORMid", sMsg, "表單類別")
                sValidOK = sValidOK And CheckNotEmpty(Me, "txtContacts", sMsg, "連絡人")
                '
                If (txtFORMid.Text.Substring(0, 3) = "066") Or (txtFORMid.Text.Substring(0, 3) = "067") Then '服務離職證明
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrA1", sMsg, "用途")
                    'modDB.InsertSignRecord("AziTest", TxtRefSTRA1.Text + "=" + System.Text.Encoding.Default.GetBytes(TxtRefSTRA1.Text).Length.ToString, My.User.Name) 'AZITEST
                ElseIf (txtFORMid.Text.Substring(0, 3) = "068") Or (txtFORMid.Text.Substring(0, 3) = "069") Then '健保轉出入
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrA", sMsg, "轉出入姓名")
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrB", sMsg, "身分字號")
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrC", sMsg, "生日")
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrD", sMsg, "生效日")
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrE", sMsg, "稱謂")
                ElseIf (txtFORMid.Text.Substring(0, 3) = "065") Then '薪資單
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtDtStart", sMsg, "起始年月")
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtDtEnd", sMsg, "迄止年月")
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrA1", sMsg, "用途")
                ElseIf (txtFORMid.Text.Substring(0, 3) = "070") Or (txtFORMid.Text.Substring(0, 3) = "071") Then '扣繳憑單/繳費證明
                    sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrA", sMsg, "年度")
                    If (txtFORMid.Text.Substring(0, 3) = "070") Then
                        sValidOK = sValidOK And CheckNotEmpty(Me, "txtRefStrA1", sMsg, "用途")
                    End If
                End If
                '
                If System.Text.Encoding.Default.GetBytes(TxtRefSTRA1.Text).Length > 20 Then
                    sValidOK = False
                    sMsg = "用途欄位長度限制10中文字或20個字元!請重新輸入。"
                End If
                'modDB.InsertSignRecord("azitest", "ok-1", My.User.Name)
                '檢查資料重複
                If sValidOK And (pMode = FormViewMode.Insert) Then
                    Dim sConn As SqlClient.SqlConnection = Nothing
                    Dim sCmd As SqlClient.SqlCommand = Nothing
                    Dim sSql As String

                    sConn = New SqlClient.SqlConnection
                    sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
                    sConn.Open()

                    sCmd = New SqlClient.SqlCommand()
                    sCmd.Connection = sConn

                    Dim sReader As SqlClient.SqlDataReader = Nothing
                    sCmd = New SqlClient.SqlCommand()
                    sCmd.Connection = sConn
                    '-------------------------------------* 找出是否有重複資料
                    sSql = "SELECT HAP_STNID FROM HAFORMAPPLY " _
                         & " WHERE HAP_STNID='" & txtSTNID.Text & "'" _
                         & "   AND HAP_EMPID='" & txtEmpID.Text & "'" _
                         & "   AND HAP_FORMID='" & txtFORMid.Text.Substring(0, 3) & "'" _
                         & "   AND HAP_STATUS='N' AND HAP_DELETOR=''"
                    If (txtFORMid.Text.Substring(0, 3) = "066") Then '檢查資料重複
                        sSql = sSql + " AND HAP_REFSTRA='" + txtRefStrA.Text + "'"
                    End If
                    sCmd.CommandText = sSql
                    'modDB.InsertSignRecord("azitest", "sSql=" & sSql.Trim, My.User.Name)
                    sReader = sCmd.ExecuteReader()
                    sReader.Read()
                    '
                    If (sReader.HasRows) Then
                        sValidOK = False
                        sMsg = "已有資料重複!"
                    End If
                    '
                    sReader.Close()
                    '
                End If
                'modDB.InsertSignRecord("azitest", "ok-2", My.User.Name)
                '
                If Not sValidOK Then
                    showMsg(Me.Page, "訊息(請修正錯誤欄位)", sMsg)
                Else
                    If sMsg <> "" Then modUtil.showMsg(Me.Page, "訊息", "請注意：" & sMsg)
                End If
            Else
                '檢查是否已處理
                'sSql = "SELECT HAP_STATUS FROM HAFORMAPPLY " _
                '     & " WHERE HAP_SERID='" & TxtSERID.Text & "'"
                'sCmd.CommandText = sSql
                ''modDB.InsertSignRecord("azitest", "sSql=" & sSql.Trim, My.User.Name)
                'sReader = sCmd.ExecuteReader()
                'sReader.Read()
                ''
                'If (sReader.GetString(0) = "Y") Or (sReader.GetString(0) = "I") Then
                '    sValidOK = False
                '    'sMsg = "此筆資料總部已處理，不可刪除!"
                '    showMsg(Me.Page, "此筆資料總部已處理，不可刪除!", sMsg)

                'End If
                ''
                'sReader.Close()
            End If
        Catch ex As Exception
            sValidOK = False
            modUtil.showMsg(Me.Page, "錯誤訊息(CheckData)", ex.Message)
        End Try
        Return sValidOK
    End Function

    '******************************************************************************************************
    '* DB 更新處理(主檔)
    '******************************************************************************************************
    Public Sub UpdateDB(ByVal pMode As FormViewMode, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs)
        Dim VSHEDDATE As String = ""
        Dim LCKDT As String = GET_LOCKYM("1", txtSTNID.Text) & "20"
        Dim VHOURS As Double = 0

        'modUtil.showMsg(Me.Page, "鎖檔日期", LCKDT)
        'If (pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit) Then
        '    txtSHEDDATE.Text = Me.txtSHSTDATE.Text
        '    If Me.txtSHSTTIME.Text > Me.txtSHEDTIME.Text Then
        '        txtSHEDDATE.Text = Format(CDate(Format(CInt(Me.txtSHSTDATE.Text), "0000/00/00")).AddDays(1), "yyyyMMdd")
        '    End If
        'End If

        Try
            If Not ((pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit)) Then
                If Not (Me.TxtStatus.Text = "未處理") Then
                    'sValidOK = False
                    'sMsg = "此筆資料總部已處理，不可刪除!"
                    'showMsg(Me.Page, "此筆資料總部已處理，不可刪除!", sMsg)
                    modUtil.showMsg(Me.Page, "訊息", "此筆資料總部已處理，不可刪除!")
                    e.Cancel = True
                    Exit Sub
                End If
            End If

            If Not CheckData(pMode) Then '* 取消新增/修改
                e.Cancel = True
                Exit Sub
            End If
            'VHOURS = modUnset.COUNTHOURS(Me.txtSHSTTIME.Text, Me.txtSHEDTIME.Text)

            'If Me.txtRTSTTIME.Text.Trim <> "" Then
            '    VHOURS = VHOURS - modUnset.COUNTHOURS(Me.txtRTSTTIME.Text, Me.txtRTEDTIME.Text)
            'End If

            'Me.txtWORKHOUR.Text = VHOURS.ToString
            With e.Command
                'modDB.InsertSignRecord("azitest", "OK-d", My.User.Name)
                If pMode = FormViewMode.Insert Then
                    .Parameters("@STNID").Value = Me.txtSTNID.Text
                    .Parameters("@EMPID").Value = Me.txtEmpID.Text
                    .Parameters("@FORMID").Value = Me.txtFORMid.Text.Substring(0, 3)
                    .Parameters("@Contacts").Value = Me.TxtContacts.Text
                    .Parameters("@REFSTRA").Value = ""
                    .Parameters("@REFSTRB").Value = ""
                    .Parameters("@REFSTRC").Value = ""
                    .Parameters("@REFSTRD").Value = ""
                    .Parameters("@REFSTRE").Value = ""
                    If (txtFORMid.Text.Substring(0, 3) = "066") Or (txtFORMid.Text.Substring(0, 3) = "067") Then
                        .Parameters("@REFSTRA").Value = Me.TxtRefSTRA1.Text.Trim
                    ElseIf (txtFORMid.Text.Substring(0, 3) = "068") Or (txtFORMid.Text.Substring(0, 3) = "069") Then
                        .Parameters("@REFSTRA").Value = Me.txtRefStrA.Text
                        .Parameters("@REFSTRB").Value = Me.txtRefStrB.Text
                        .Parameters("@REFSTRC").Value = Me.txtRefStrC.Text
                        .Parameters("@REFSTRD").Value = Me.txtRefStrD.Text
                        .Parameters("@REFSTRE").Value = Me.txtRefSTRE.Text
                    ElseIf (txtFORMid.Text.Substring(0, 3) = "065") Then '新資單
                        .Parameters("@REFSTRA").Value = Me.txtDtStart.Text
                        .Parameters("@REFSTRB").Value = Me.txtDtEnd.Text
                        .Parameters("@REFSTRC").Value = Me.TxtRefSTRA1.Text.Trim
                    ElseIf (txtFORMid.Text.Substring(0, 3) = "070") Or (txtFORMid.Text.Substring(0, 3) = "071") Then '扣繳憑單
                        .Parameters("@REFSTRA").Value = Me.txtRefStrA.Text
                        If (txtFORMid.Text.Substring(0, 3) = "070") Then
                            .Parameters("@REFSTRC").Value = Me.TxtRefSTRA1.Text.Trim
                        End If
                    Else
                        '.Parameters("@REFSTRA").Value = ""
                        '.Parameters("@REFSTRB").Value = ""
                        '.Parameters("@REFSTRC").Value = ""
                        '.Parameters("@REFSTRD").Value = ""
                        '.Parameters("@REFSTRE").Value = ""
                    End If
                    .Parameters("@Creator").Value = Me.Page.User.Identity.Name
                    '-------------------------------------* 刪除舊資料
                    .Parameters("@ApplyDate").Value = DateTime.Now.ToString("yyyyMMdd")
                    .Parameters("@ApplyTime").Value = DateTime.Now.ToString("HHmmss")
                Else
                    'modDB.InsertSignRecord("azitest", "OK-A=" + Me.TxtSERID.Text, My.User.Name)
                    'modDB.InsertSignRecord("azitest", "OK-A=" + Me.Page.User.Identity.Name, My.User.Name)
                    .Parameters("@SERID").Value = Convert.ToInt32(Me.TxtSERID.Text)
                    .Parameters("@DELETOR").Value = Me.Page.User.Identity.Name
                    'modDB.InsertSignRecord("azitest", "OK-B", My.User.Name)
                End If
            End With
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(UpdateDB)", ex.Message)
        End Try

    End Sub

    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryStn.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
    End Sub

    Protected Sub txtStnID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'modCharge.GetStnName(Me.txtSTNID, Me.txtStnName, Me.ViewState("RolDeptID"))
    End Sub

    Protected Sub txtEmpID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call modCharge.GetEmpName(Me.txtEmpID, Me.txtEmpName, True)
    End Sub

    Protected Sub btnEmp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryEmp.Show(Me.txtSTNID.Text)
    End Sub

    Protected Sub btnEmp2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryEmp.Show(Me.txtSTNID.Text)
    End Sub

    Protected Sub txtFORMid_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFORMid.SelectedIndexChanged
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label9.Visible = False
        txtRefStrA.Visible = False
        txtRefStrB.Visible = False
        txtRefStrC.Visible = False
        txtRefStrD.Visible = False
        txtRefSTRE.Visible = False
        TxtRefSTRA1.Visible = False
        '
        Label11.Visible = False
        txtDtStart.Visible = False
        imgDtStart.Visible = False
        Label10.Visible = False
        txtDtEnd.Visible = False
        imgDtEnd.Visible = False
        '
        If (txtFORMid.SelectedItem.Text.Substring(0, 3) = "066") Or (txtFORMid.SelectedItem.Text.Substring(0, 3) = "067") Then  '服務、離職證明
            Label5.Text = "用途:"
            Label1.Visible = True
            Label5.Visible = True
            TxtRefSTRA1.Visible = True
        ElseIf (txtFORMid.SelectedItem.Text.Substring(0, 3) = "068") Or (txtFORMid.SelectedItem.Text.Substring(0, 3) = "069") Then  '健保轉入 '健保轉出
            Label5.Text = "姓名:"
            Label6.Text = "身分字號:"
            Label1.Visible = True
            Label6.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Label4.Visible = True
            Label5.Visible = True
            Label7.Visible = True
            Label9.Visible = True
            txtRefStrA.Visible = True
            txtRefStrB.Visible = True
            txtRefStrC.Visible = True
            txtRefStrD.Visible = True
            txtRefSTRE.Visible = True
        ElseIf (txtFORMid.SelectedItem.Text.Substring(0, 3) = "065") Then  '薪資單
            Label11.Visible = True
            txtDtStart.Visible = True
            imgDtStart.Visible = True
            Label10.Visible = True
            txtDtEnd.Visible = True
            imgDtEnd.Visible = True
            '
            Label5.Text = "用途:"
            Label1.Visible = True
            Label5.Visible = True
            TxtRefSTRA1.Visible = True
        ElseIf (txtFORMid.SelectedItem.Text.Substring(0, 3) = "070") Or (txtFORMid.SelectedItem.Text.Substring(0, 3) = "071") Then  '扣繳憑單/繳費證明
            Label5.Text = "年度:"
            Label1.Visible = True
            Label5.Visible = True
            txtRefStrA.Visible = True
            If (txtFORMid.SelectedItem.Text.Substring(0, 3) = "070") Then
                Label6.Text = "用途:"
                Label6.Visible = True
                TxtRefSTRA1.Visible = True
            End If

        Else
            'Label1.Visible = False
            'Label2.Visible = False
            'Label3.Visible = False
            'Label4.Visible = False
            'Label5.Visible = False
            'Label6.Visible = False
            'Label7.Visible = False
            'Label9.Visible = False
            'txtRefStrA.Visible = False
            'txtRefStrB.Visible = False
            'txtRefStrC.Visible = False
            'txtRefStrD.Visible = False
            'txtRefSTRE.Visible = False
        End If
    End Sub

    Protected Sub txtFORMid_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFORMid.TextChanged
        'modDB.InsertSignRecord("azitest", "0606_OK-1:" & CLSTIME.Text.Trim, My.User.Name)
        'If CLSTIME.Text.Trim <> "" Then
        '    txtSHSTTIME.Text = CLSTIME.Text.Substring(5, 4)
        '    txtSHEDTIME.Text = CLSTIME.Text.Substring(10, 4)
        '    If CLSTIME.Text.Substring(15, 1) = "Y" Then '正常班
        '        chkNormalCls.Checked = True
        '    Else
        '        chkNormalCls.Checked = False
        '    End If
        'End If
    End Sub

    Protected Sub txtRefStrB_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRefStrB.TextChanged
        txtRefStrB.Text = UCase(txtRefStrB.Text)
    End Sub

    Protected Sub UsrCalendar_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles UsrCalendar.Load

    End Sub
End Class
