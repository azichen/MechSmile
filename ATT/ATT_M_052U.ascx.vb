'******************************************************************************************************
'* 程式：ATT_M_052U 公出資料維護
'* 版次：2013/05/24(VER1.01)：新開發
'******************************************************************************************************

Imports modUtil
Imports modUnset
Imports System.Data


Partial Class ATT_M_052U
    Inherits System.Web.UI.UserControl

    'Private cWriteLog As Boolean = False '* 上線時須設成 False
    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Public Sub InitScreen(ByVal pMode As FormViewMode)
        LEditmode.Text = ""

        If pMode = FormViewMode.ReadOnly Then
            SetAllReadOnly(Me)
            Me.btnEmp.Visible = False
            Me.BTNVCM.Visible = False
        Else
            SetObjReadOnly(Me, "txtEmpName")
            'SetObjReadOnly(Me, "txtVATMHOUR")
            Me.btnEmp.Visible = False
            Me.txtEmpID.Attributes.Add("onkeypress", "return CheckKeyAZ09()")
            Me.txtVATMHOUR.Attributes.Add("onkeypress", "return CheckKey09()")
            If pMode = FormViewMode.Insert Then
                Me.btnEmp.Visible = True
                Me.BTNVCM.Visible = True
                'If modUtil.IsStnRol(Me.Request) Then '加油站只能新增公出假
                '    TxtVANMID.Text = "010"
                '    TxtVANMNAME.Text = "公出假"
                '    SetObjReadOnly(Me, "TxtVANMID")
                '    SetObjReadOnly(Me, "TxtVANMNAME")
                '    Me.BTNVCM.Visible = False
                '    Me.txtVATMHOUR.Visible = False
                'Else
                'Me.BTNVCM.Visible = True
                'End If
                '只考慮新增模式()
                'Me.txtEmpID.Text = HttpUtility.UrlDecode(Request.Cookies("VATM_EMPLID").Value)
                'Me.txtEmpName.Text = HttpUtility.UrlDecode(Request.Cookies("EMPNAME").Value)
                If Session("ATT052MODE") = "2" Then '出勤查詢模式
                    txtEmpID.Text = Request("EmpID")
                    txtEmpName.Text = Request("EmpName")
                    txtVATMDATEST.Text = Request("SHSTDATE")
                    txtVATMDATEEN.Text = Request("SHSTDATE")
                    SetObjReadOnly(Me, "TxtEmpID")
                    SetObjReadOnly(Me, "TxtEmpName")
                End If
                modUtil.SetObjFocus(Me.Page, Me.txtEmpID)
                LEditmode.Text = "Y"
            Else
                '20140225 應該要限制可以維護的假別，再與人資TEAM確認。
                SetObjReadOnly(Me, "txtEmpID")
                '
                Me.BTNVCM.Visible = False
                SetObjReadOnly(Me, "TxtVANMID")
                SetObjReadOnly(Me, "TxtVANMNAME")
                '20140916_2014.10月起加油站可以自行維護所有假別
                'If modUtil.IsStnRol(Me.Request) Then '加油站只能維護公出假
                '    Me.BTNVCM.Visible = False
                '    SetObjReadOnly(Me, "TxtVANMID")
                '    SetObjReadOnly(Me, "TxtVANMNAME")
                'End If
                '
                txtSTDT_tmp.Text = txtVATMDATEST.Text
                txtSTTM_tmp.Text = txtVATMTIMEST.Text
                txtENTM_tmp.Text = txtVATMTIMEEN.Text
                LEditmode.Text = "Y"
            End If
        End If
    End Sub

    '******************************************************************************************************
    '* DB 更新前置檢核處理
    '******************************************************************************************************
    Private Function CheckData(ByVal pMode As FormViewMode) As Boolean
        Dim sValidOK As Boolean = True
        Dim LOOPCHK As Boolean = True
        Dim sMsg As String = ""
        Dim CHKRSTR As String = ""
        Dim I As Integer = 0

        Try
            If pMode = FormViewMode.Insert Then '* 新增模式
                sValidOK = sValidOK And CheckLength(Me, "txtEMPID", 8, sMsg, "員工編號")
            End If
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtVATMDATEST", sMsg, "請假起始日期")
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtVATMDATEEN", sMsg, "請假迄止日期")
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtVATMTIMEST", sMsg, "起始時間")
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtVATMTIMEEN", sMsg, "結束時間")
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtVATMHOUR", sMsg, "請假時數")
            '假別檢查 20140225 增加
            If sValidOK Then
                If TxtVANMNAME.Text = "" Then
                    sValidOK = False : sMsg = "請假別錯誤!請檢查!"
                End If
            End If

            '請假日期檢查
            If sValidOK Then
                CHKRSTR = modUnset.PADATECHK(txtVATMDATEST.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "請假起始日期錯誤!請檢查!"
                Else
                    txtVATMDATEST.Text = CHKRSTR.Substring(1, 8)
                End If
            End If

            '請假迄止日期檢查
            If sValidOK Then
                CHKRSTR = modUnset.PADATECHK(txtVATMDATEEN.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "請假迄止日期錯誤!請檢查!"
                Else
                    txtVATMDATEEN.Text = CHKRSTR.Substring(1, 8)
                End If
            End If

            '起始時間檢查
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(txtVATMTIMEST.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "起始時間錯誤!請檢查!"
                Else
                    txtVATMTIMEST.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '結束時間檢查
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(txtVATMTIMEEN.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "結束時間錯誤!請檢查!"
                Else
                    txtVATMTIMEEN.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '日期合理性檢查
            If sValidOK Then
                If (txtVATMDATEST.Text + txtVATMTIMEST.Text) >= (txtVATMDATEEN.Text + txtVATMTIMEEN.Text) Then
                    sValidOK = False : sMsg = sMsg & "\n * 請假起迄時間錯誤!請檢查!"
                End If
            End If
            '
            If sValidOK Then '再作鎖檔日期檢核
                Dim LCKDT As String = GET_LOCKDT(Request.Cookies("STNID").Value)
                'Dim LCKDT As String = GET_LOCKYM("1", Request.Cookies("STNID").Value) & "20"
                If Me.txtVATMDATEST.Text <= LCKDT Then
                    sValidOK = False
                    sMsg = "資料週期已鎖或過帳，不可異動!"
                End If
            End If
            '
            '請假日與員工資料作檢核
            If sValidOK Then
                Dim sConn As SqlClient.SqlConnection = Nothing
                Dim sCmd As SqlClient.SqlCommand = Nothing
                Dim sSql As String
                Dim sReader As SqlClient.SqlDataReader = Nothing
                '
                If (pMode = FormViewMode.Insert) Then
                    Dim VATMDATEEN As String = txtVATMDATEST.Text
                    If VATMDATEEN < txtVATMTIMEST.Text Then
                        VATMDATEEN = Format(CDate(Format(CInt(txtVATMDATEST.Text), "0000/00/00")).AddDays(1), "yyyyMMdd")
                    End If

                    sConn = New SqlClient.SqlConnection
                    sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
                    sConn.Open()

                    sCmd = New SqlClient.SqlCommand()
                    sCmd.Connection = sConn

                    sSql = " SELECT ISNULL(A.EMPL_SVN_DATE,'') AS SVNDATE,ISNULL(B.EMCG_DATE,'') AS STOPDATE "
                    sSql += "      ,ISNULL(C.EMPL_ARV_DATE,'') AS ARVDATE,ISNULL(D.EMPL_LEV_DATE,'') AS LEVDATE"
                    sSql += "      ,ISNULL(A.EMPL_DEPTID,'') AS DEPTID"
                    sSql += "  FROM MP_HR.DBO.EMPLOYEE A"
                    sSql += "  LEFT JOIN MP_HR.DBO.EMPLOYEE_CHANGE B ON B.EMCG_EMPLID='" & txtEmpID.Text & "' AND B.EMCG_TYPE ='6' "
                    sSql += "                                 AND '" & txtVATMDATEST.Text & "' BETWEEN B.EMCG_DATE AND B.EMCG_STAY_DATE"
                    sSql += "  LEFT JOIN MP_HR.DBO.EMPLOYEE C ON C.EMPL_ID='" & txtEmpID.Text & "' AND C.EMPL_ARV_DATE >'" & txtVATMDATEST.Text & "'"
                    sSql += "  LEFT JOIN MP_HR.DBO.EMPLOYEE D ON D.EMPL_ID='" & txtEmpID.Text & "' AND A.EMPL_LEV_DATE IS NOT NULL AND A.EMPL_LEV_DATE <='" & txtVATMDATEEN.Text & "'"
                    sSql += "  WHERE A.EMPL_ID='" & txtEmpID.Text & "'  "

                    sCmd.CommandText = sSql
                    sReader = sCmd.ExecuteReader()
                    sReader.Read()
                    If Not IsDBNull(sReader.GetValue(0)) Then
                        '員工所屬單位與輸入人員單位檢查
                        'If (modUtil.IsStnRol(Me.Request) Or modUnset.IsPAUnitRol(Me.Request)) Then 
                        '目前一律檢查:員工所屬單位與輸入人員單位一定要一致
                        If sReader.GetValue(4) <> HttpUtility.UrlDecode(Request.Cookies("STNID").Value) Then
                            sValidOK = False
                            sMsg = "非本單位員工不可維護!"
                        End If
                        'End If
                        If sReader.GetValue(2) >= "1980-01-01" Then
                            sValidOK = False
                            sMsg = "請假日期不能比到職日早!"
                        End If
                        '
                        If sValidOK And sReader.GetValue(3) >= "1980-01-01" Then
                            sValidOK = False
                            sMsg = "該員工於此日期時為已離職狀態!"
                        End If
                        '
                        If sValidOK And sReader.GetValue(1) >= "1980-01-01" Then
                            sValidOK = False
                            sMsg = "該員工於此日期時為留職停薪狀態!"
                        End If
                        '
                        If sValidOK And (TxtVANMID.Text = "000") Then
                            If sReader.GetValue(0) > (txtVATMDATEST.Text.Substring(0, 4) & "-" & txtVATMDATEST.Text.Substring(4, 2) & "-" & txtVATMDATEST.Text.Substring(6, 2)) Then
                                sValidOK = False
                                sMsg = "特休假必須到職滿1年!!"
                            End If
                        End If
                    Else
                        sValidOK = False
                        sMsg = "查無此人基本資料!不允許維護資料。"
                    End If
                    sReader.Close()
                    '
                    If sValidOK Then
                        '20150817 增加檢查前後3小時內的重覆資料
                        If TxtVANMID.Text <> "010" Then
                            'modDB.InsertSignRecord("AziTest", "VMCHKDTST=" + txtVATMDATEST.Text.Substring(0, 4) + "/" + txtVATMDATEST.Text.Substring(4, 2) + "/" + txtVATMDATEST.Text.Substring(6, 2) _
                            '                      + " " + txtVATMTIMEST.Text.Substring(0, 2) + ":" + txtVATMTIMEST.Text.Substring(2, 2), My.User.Name) 'AZITEST
                            Dim VMCHKDTST As String = DateTime.Parse(txtVATMDATEST.Text.Substring(0, 4) + "/" + txtVATMDATEST.Text.Substring(4, 2) + "/" + txtVATMDATEST.Text.Substring(6, 2) _
                                                    + " " + txtVATMTIMEST.Text.Substring(0, 2) + ":" + txtVATMTIMEST.Text.Substring(2, 2)).AddHours(-3).ToString("yyyyMMddHHmm")
                            Dim VMCHKDTEN As String = DateTime.Parse(txtVATMDATEEN.Text.Substring(0, 4) + "/" + txtVATMDATEEN.Text.Substring(4, 2) + "/" + txtVATMDATEEN.Text.Substring(6, 2) _
                                                    + " " + txtVATMTIMEEN.Text.Substring(0, 2) + ":" + txtVATMTIMEEN.Text.Substring(2, 2)).AddHours(3).ToString("yyyyMMddHHmm")
                            'modDB.InsertSignRecord("AziTest", "VMCHKDTST=" + VMCHKDTST + " VMCHKDTEN=" + VMCHKDTEN, My.User.Name) 'AZITEST
                            '-------------------------------------* 找出是否有重複資料
                            sSql = "SELECT VATM_VANMID,VATM_DATE_ST,VATM_TIME_ST FROM MP_HR.DBO.VACATM " _
                                 & " WHERE VATM_EMPLID='" & txtEmpID.Text & "'" _
                                 & "   AND ((CONVERT(char(8), VATM_DATE_EN,112)+VATM_TIME_EN)>'" & VMCHKDTST & "')" _
                                 & "   AND ((CONVERT(char(8), VATM_DATE_ST,112)+VATM_TIME_ST)<'" & VMCHKDTEN & "')" _
                                 & "   AND VATM_VANMID<>'010'"
                            '& "   AND ((CONVERT(char(8), VATM_DATE_EN,112)+VATM_TIME_EN)>'" & txtVATMDATEST.Text & txtVATMTIMEST.Text & "')" _
                            '& "   AND ((CONVERT(char(8), VATM_DATE_ST,112)+VATM_TIME_ST)<'" & VATMDATEEN & txtVATMTIMEEN.Text & "')"
                            '
                            If (pMode = FormViewMode.Edit) Then '修改時不要檢查到自己
                                sSql += " AND VATM_SEQ<>" & Me.txtVATMSEQ.Text
                            End If
                            '
                            sCmd.CommandText = sSql
                            sReader = sCmd.ExecuteReader()
                            sReader.Read()
                            If (sReader.HasRows) Then
                                sValidOK = False
                                sMsg = "起迄前後3小時內，有請假資料:(假別" & sReader.GetString(0) & " 日期: " & sReader.GetDateTime(1).ToShortDateString & " 時間 " & sReader.GetString(2) & " )"
                            End If
                            sReader.Close()
                        End If
                    End If
                    '
                    If sValidOK Then
                        '-------------------------------------* 找出是否有重複資料
                        sSql = "SELECT VATM_VANMID,VATM_DATE_ST,VATM_TIME_ST FROM MP_HR.DBO.VACATM " _
                             & " WHERE VATM_EMPLID='" & txtEmpID.Text & "'" _
                             & "   AND ((CONVERT(char(8), VATM_DATE_EN,112)+VATM_TIME_EN)>'" & txtVATMDATEST.Text & txtVATMTIMEST.Text & "')" _
                             & "   AND ((CONVERT(char(8), VATM_DATE_ST,112)+VATM_TIME_ST)<'" & VATMDATEEN & txtVATMTIMEEN.Text & "')"
                        '
                        If (pMode = FormViewMode.Edit) Then '修改時不要檢查到自己
                            sSql += " AND VATM_SEQ<>" & Me.txtVATMSEQ.Text
                        End If
                        '
                        sCmd.CommandText = sSql
                        sReader = sCmd.ExecuteReader()
                        sReader.Read()
                        If (sReader.HasRows) Then
                            sValidOK = False
                            sMsg = "請假資料重複:(假別" & sReader.GetString(0) & " 日期: " & sReader.GetDateTime(1).ToShortDateString & " 時間 " & sReader.GetString(2) & " )"
                        End If
                        sReader.Close()
                    End If
                    '
                    If sValidOK Then
                        '新增時抓新序號
                        If (pMode = FormViewMode.Insert) Then
                            sSql = "SELECT MAX(VATM_SEQ) AS VATM_SEQ FROM MP_HR.DBO.VACATM " _
                                 & " WHERE VATM_EMPLID='" & txtEmpID.Text & "'"
                            sCmd.CommandText = sSql
                            sReader = sCmd.ExecuteReader()
                            sReader.Read()

                            If IsDBNull(sReader.GetValue(0)) Then
                                txtVATMSEQ.Text = "1"
                            Else
                                txtVATMSEQ.Text = (sReader.GetInt16(0) + 1).ToString
                            End If
                            '
                        End If
                    End If
                    sConn.Close()
                    sCmd.Dispose() '手動釋放資源
                    sConn.Dispose()
                    sConn = Nothing '移除指標
                End If
                '
                '作時數檢查
                If sValidOK Then
                    If (pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit) Then
                        '20150817 
                        Dim CHOURS As Double = COUNTHOURS(Me.txtVATMDATEST.Text, Me.txtVATMTIMEST.Text, Me.txtVATMDATEEN.Text, Me.txtVATMTIMEEN.Text)
                        '
                        Dim VHOURS As Double = Convert.ToDouble(Me.txtVATMHOUR.Text)
                        '
                        If VHOURS - CHOURS > 0 Then
                            'azitemp
                            'modDB.InsertSignRecord("AziTest", "attm052u_CHOUR=" + CHOURS.ToString + " VHOURS=" + VHOURS.ToString, My.User.Name) 'AZITEST
                            sValidOK = False
                            sMsg = "請假時數錯誤請檢查!注意請假時間若不連續，需分開輸入。"
                        End If
                        'add in 20151119
                        If Not (VHOURS > 0) Then
                            sValidOK = False
                            sMsg = "時數錯誤,不可為0 請檢查!"
                        End If
                        '
                        If sValidOK Then
                            Dim HHR As Double = 8
                            If TxtVANMID.Text = "010" Then '公出
                                HHR = 12
                            End If
                            '
                            If VHOURS > HHR Then
                                sValidOK = False
                                sMsg = "請假時數超出 " + HHR.ToString + " 小時，不可自行輸入。"
                            End If
                        End If
                        '
                        If sValidOK Then
                            Dim CHECKSTR As String = HODHOURCHECK(txtEmpID.Text, Me.txtVATMDATEST.Text, Me.txtVATMDATEEN.Text, TxtVANMID.Text)
                            If CHECKSTR.Substring(14, 1) = "Y" Then '須與可請假核對
                                If Convert.ToDouble(CHECKSTR.Substring(0, 7)) > 0 Then
                                    If (pMode = FormViewMode.Edit) And (txtVATMHOUR.Text <> "") Then
                                        VHOURS = VHOURS - Convert.ToDouble(txtVATMHOUR.Text)
                                    End If
                                    '
                                    If (Convert.ToDouble(CHECKSTR.Substring(0, 7)) - Convert.ToDouble(CHECKSTR.Substring(7, 7)) - VHOURS) < 0 Then
                                        sValidOK = False
                                        sMsg = "請假時數超過年度限制!請檢查!(可請時數:" & Convert.ToDouble(CHECKSTR.Substring(0, 7)).ToString & ",已請時數:" & Convert.ToDouble(CHECKSTR.Substring(7, 7)).ToString & ",本次時數:" & VHOURS.ToString & ")"
                                    Else
                                        '014生理假再作月檢查，每月只能請一次。azitemp 20140915
                                    End If
                                Else
                                    sValidOK = False
                                    sMsg = "查無該假別可請假之時數!請檢查!(可請時數:" & Convert.ToDouble(CHECKSTR.Substring(0, 7)).ToString & ",已請時數:" & Convert.ToDouble(CHECKSTR.Substring(7, 7)).ToString & ",本次時數:" & VHOURS.ToString & ")"""
                                End If
                            End If
                        End If
                        '20140930 增加病假 002
                        '20150121 改:只能申請2天
                        If sValidOK Then
                            'If (TxtVANMID.Text = "002") And (VHOURS > 24) Then
                            '    sValidOK = False
                            '    sMsg = "病假申請時數!已超出24小時(3天),不可自行輸入!"
                            'End If
                            If (Me.TxtVANMID.Text <> "010") And modUtil.IsStnRol(Me.Request) Then '20150331
                                Dim VATMDAYS As Integer = 0
                                Dim VATMHOURS As Double = 0
                                Dim DATE_S As Date
                                Dim DATE_E As Date

                                DATE_S = CDate(Format(CInt(txtVATMDATEST.Text), "0000/00/00"))
                                DATE_E = CDate(Format(CInt(txtVATMDATEEN.Text), "0000/00/00"))
                                VATMDAYS = 0
                                Do While DATE_S.Date <= DATE_E.Date
                                    VATMDAYS = VATMDAYS + 1
                                    DATE_S = DATE_S.AddDays(1)
                                Loop
                                VATMHOURS = VHOURS
                                'If (Format(CDate(Format(CInt(txtVATMDATEST.Text), "0000/00/00")).AddDays(1), "yyyyMMdd") < txtVATMDATEEN.Text.Trim) Or (VHOURS > 24) Then
                                If (VATMDAYS > 2) Or (VATMHOURS > 16) Then
                                    sValidOK = False
                                    'modDB.InsertSignRecord("AziTest", "P1:VATMDAYS=" + VATMDAYS.ToString + " VHOURS=" + VHOURS.ToString, My.User.Name) 'AZITEST
                                    sMsg = "已超出2天16小時,需送假單不可自行輸入!"
                                Else
                                    '20150121 還要檢查前一天與後一天是否有請假資料
                                    sConn = New SqlClient.SqlConnection
                                    sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
                                    sConn.Open()
                                    sCmd = New SqlClient.SqlCommand()
                                    sCmd.Connection = sConn
                                    'for i =1 to 2 loop 
                                    I = 1
                                    LOOPCHK = True
                                    CHKRSTR = Format(CDate(Format(CInt(txtVATMDATEST.Text), "0000/00/00")).AddDays(-1), "yyyy/MM/dd")
                                    Do While (I <= 2) And sValidOK And LOOPCHK '連續往前檢查兩天
                                        sReader.Close()
                                        '-------------------------------------* 前一天是否有重複資料
                                        sSql = "SELECT VATM_VANMID,VATM_DATE_ST,VATM_DATE_EN,VATM_HOURS FROM MP_HR.DBO.VACATM " _
                                             & " WHERE VATM_EMPLID='" & txtEmpID.Text & "'" _
                                             & "   AND VATM_DATE_EN='" & CHKRSTR & "'" _
                                             & "   AND VATM_VANMID<>'010'"
                                        '
                                        sCmd.CommandText = sSql
                                        sReader = sCmd.ExecuteReader()
                                        sReader.Read()
                                        If (sReader.HasRows) Then
                                            DATE_S = sReader.GetDateTime(1)
                                            DATE_E = sReader.GetDateTime(2)
                                            VATMHOURS = VATMHOURS + Convert.ToDouble(sReader.GetValue(3))
                                            Do While DATE_S.Date <= DATE_E.Date
                                                VATMDAYS = VATMDAYS + 1
                                                DATE_S = DATE_S.AddDays(1)
                                            Loop
                                        Else
                                            LOOPCHK = False
                                        End If
                                        '
                                        If (VATMDAYS > 2) And (VATMHOURS > 16) Then
                                            sValidOK = False
                                            'modDB.InsertSignRecord("AziTest", "P2:VATMDAYS=" + VATMDAYS.ToString + " VATMHOURS=" + VATMHOURS.ToString + " CHKRSTR=" + CHKRSTR, My.User.Name) 'AZITEST
                                            sMsg = "包含 " + CHKRSTR + " 請假，連續請假超出2天16小時,需送假單不可自行輸入!"
                                        End If
                                        I = I + 1
                                        CHKRSTR = Format(CDate(CHKRSTR).AddDays(-1), "yyyy/MM/dd")
                                    Loop
                                    '
                                    '再找後一天的請假資料
                                    If sValidOK Then
                                        I = 1
                                        LOOPCHK = True
                                        CHKRSTR = Format(CDate(Format(CInt(txtVATMDATEEN.Text), "0000/00/00")).AddDays(1), "yyyy/MM/dd")
                                        Do While (I <= 2) And sValidOK And LOOPCHK '連續往前檢查兩天
                                            sReader.Close()
                                            '-------------------------------------* 前一天是否有重複資料
                                            sSql = "SELECT VATM_VANMID,VATM_DATE_ST,VATM_DATE_EN,VATM_HOURS FROM MP_HR.DBO.VACATM " _
                                                 & " WHERE VATM_EMPLID='" & txtEmpID.Text & "'" _
                                                 & "   AND VATM_DATE_EN='" & CHKRSTR & "'" _
                                                 & "   AND VATM_VANMID<>'010'"
                                            '
                                            sCmd.CommandText = sSql
                                            sReader = sCmd.ExecuteReader()
                                            sReader.Read()
                                            If (sReader.HasRows) Then
                                                DATE_S = sReader.GetDateTime(1)
                                                DATE_E = sReader.GetDateTime(2)
                                                VATMHOURS = VATMHOURS + Convert.ToDouble(sReader.GetValue(3))
                                                Do While DATE_S.Date <= DATE_E.Date
                                                    VATMDAYS = VATMDAYS + 1
                                                    DATE_S = DATE_S.AddDays(1)
                                                Loop
                                                If (VATMDAYS > 2) And (VATMHOURS > 16) Then
                                                    sValidOK = False
                                                    sMsg = "包含 " + CHKRSTR + " 請假，連續請假已超出2天假,需送假單不可自行輸入!"
                                                    'modDB.InsertSignRecord("AziTest", "P3:VATMDAYS=" + VATMDAYS.ToString + " VATMHOURS=" + VATMHOURS.ToString + " CHKRSTR=" + CHKRSTR, My.User.Name) 'AZITEST
                                                End If
                                            Else
                                                LOOPCHK = False
                                            End If
                                            I = I + 1
                                            CHKRSTR = Format(CDate(CHKRSTR).AddDays(1), "yyyy/MM/dd")
                                        Loop
                                    End If
                                End If
                                sReader.Close()
                                'modDB.InsertSignRecord("AziTest", "PF:VATMDAYS=" + VATMDAYS.ToString + " VATMHOURS=" + VATMHOURS.ToString + " CHKRSTR=" + CHKRSTR, My.User.Name) 'AZITEST
                            End If 'End of If (Me.TxtVANMID.Text <> "010") Then
                        End If
                    End If
                End If
            End If
            If Not sValidOK Then showMsg(Me.Page, "訊息(請修正錯誤欄位)", sMsg)
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
        'Dim LCKDT As String = GET_LOCKYM("1", "000000") & "20"
        Dim LCKDT As String = GET_LOCKDT(Request.Cookies("STNID").Value)
        'modUtil.showMsg(Me.Page, "鎖檔日期", LCKDT)

        If pMode = FormViewMode.Insert Then
            txtSTDT_tmp.Text = txtVATMDATEST.Text
            txtSTTM_tmp.Text = txtVATMTIMEST.Text
            txtENTM_tmp.Text = txtVATMTIMEEN.Text
        ElseIf pMode = FormViewMode.Edit Then
            'txtSTDT_tmp.Text = txtVATMDATEST.Text
            'txtSTTM_tmp.Text = txtVATMTIMEST.Text
            'txtENTM_tmp.Text = txtVATMTIMEEN.Text
        End If

        Try
            If Not CheckData(pMode) Then '* 取消新增/修改
                e.Cancel = True
                Exit Sub
            Else
                '增加排班資料處理 -> N
                Dim sConn As SqlClient.SqlConnection = Nothing
                Dim sCmd As SqlClient.SqlCommand = Nothing
                Dim sSql As String
                Dim VATMDATEEN As String = txtVATMDATEST.Text
                If VATMDATEEN < txtVATMTIMEST.Text Then
                    VATMDATEEN = Format(CDate(Format(CInt(txtVATMDATEST.Text), "0000/00/00")).AddDays(1), "yyyyMMdd")
                End If

                sConn = New SqlClient.SqlConnection
                sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
                sConn.Open()
                sCmd = New SqlClient.SqlCommand()
                sCmd.Connection = sConn
                '
                sSql = "UPDATE SCHEDM SET FACFLAG='N' FROM " _
                     + " (SELECT EMPID,SHSTDATE,SHSTTIME,SHEDDATE,SHEDTIME FROM SCHEDM " _
                     + "   WHERE EMPID='" + Me.txtEmpID.Text + "'" _
                     + "     AND (SHSTDATE+SHSTTIME)<='" + VATMDATEEN + txtENTM_tmp.Text + "' AND (SHEDDATE+SHEDTIME)>='" + txtSTDT_tmp.Text + txtSTTM_tmp.Text + "') A" _
                     + " WHERE SCHEDM.EMPID=A.EMPID AND SCHEDM.SHSTDATE=A.SHSTDATE AND SCHEDM.SHSTTIME=A.SHSTTIME "
                '
                sCmd.CommandText = sSql
                sCmd.ExecuteNonQuery()
            End If


            With e.Command
                .Parameters("@VATM_EMPLID").Value = Me.txtEmpID.Text
                .Parameters("@VATM_SEQ").Value = Convert.ToInt32(Me.txtVATMSEQ.Text)

                '-------------------------------------* 刪除舊資料
                '<asp:Parameter Name="VATM_SEQ"     Type="Int32"  />
                If (pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit) Then

                    'Dim VHOURS As Double = COUNTHOURS(Me.txtVATMDATEST.Text, Me.txtVATMTIMEST.Text, Me.txtVATMDATEEN.Text, Me.txtVATMTIMEEN.Text)
                    ''azitemp
                    'If Me.txtVATMTIMEST.Text.Substring(0, 2) <= "12" And Me.txtVATMTIMEEN.Text.Substring(0, 2) >= "13" Then
                    '    If Not modUtil.IsStnRol(Me.Request) Then
                    '        VHOURS = VHOURS - 1
                    '    End If
                    'End If
                    'Me.txtVATMHOUR.Text = VHOURS.ToString
                    '
                    .Parameters("@VATM_HOURS").Value = Convert.ToDouble(Me.txtVATMHOUR.Text)
                    .Parameters("@VATM_DATE_ST").Value = Me.txtVATMDATEST.Text
                    .Parameters("@VATM_DATE_EN").Value = Me.txtVATMDATEEN.Text
                    .Parameters("@VATM_TIME_ST").Value = Me.txtVATMTIMEST.Text
                    .Parameters("@VATM_TIME_EN").Value = Me.txtVATMTIMEEN.Text
                    .Parameters("@VATM_REMARK").Value = Me.txtREMARK.Text
                    .Parameters("@VATM_U_USER_ID").Value = Me.Page.User.Identity.Name
                    .Parameters("@VATM_U_USER_NM").Value = Me.Page.User.Identity.Name
                    .Parameters("@VATM_U_DATE").Value = DateTime.Now.ToString("yyyy/MM/dd") & " " & DateTime.Now.ToString("HH:mm:ss")
                    If (pMode = FormViewMode.Insert) Then
                        Dim VINYM As String = Me.txtVATMDATEST.Text.Substring(0, 6)
                        If Me.txtVATMDATEST.Text.Substring(6, 2) > "20" Then
                            If Me.txtVATMDATEST.Text.Substring(4, 2) = "12" Then
                                VINYM = (Convert.ToInt16(Me.txtVATMDATEST.Text.Substring(0, 4)) + 1).ToString("0000") & "01"
                            Else
                                VINYM = Me.txtVATMDATEST.Text.Substring(0, 4) & (Convert.ToInt16(Me.txtVATMDATEST.Text.Substring(4, 2)) + 1).ToString("00")
                            End If
                        End If
                        .Parameters("@VATM_VANMID").Value = Me.TxtVANMID.Text
                        .Parameters("@VATM_INYM").Value = ""
                        .Parameters("@VATM_A_USER_ID").Value = Me.Page.User.Identity.Name
                        .Parameters("@VATM_A_USER_NM").Value = Me.Page.User.Identity.Name
                        .Parameters("@VATM_A_DATE").Value = DateTime.Now.ToString("yyyy/MM/dd") & " " & DateTime.Now.ToString("HH:mm:ss")
                    End If
                End If 'END OF INSERT
            End With
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(UpdateDB)", ex.Message)
        End Try

    End Sub

    Protected Sub txtEmpID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call modCharge.GetEmpName(Me.txtEmpID, Me.txtEmpName, True)
    End Sub

    Protected Sub btnEmp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryEmp.Show(Request.Cookies("STNID").Value) 'azitemp
    End Sub

    'Protected Sub txtVATMTIMEEN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVATMTIMEEN.TextChanged
    '    If (txtVATMTIMEEN.Text.Length = 4) And (txtVATMTIMEST.Text.Length = 4) Then
    '        Me.txtVATMHOUR.Text = COUNTHOURS(Me.txtVATMTIMEST.Text, Me.txtVATMTIMEEN.Text).ToString
    '    End If
    'End Sub

    Protected Sub TxtVANMID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If modUtil.IsStnRol(Me.Request) Then '加油站只能新增公出假
            Call modUnset.GetVcmName(Me.TxtVANMID, Me.TxtVANMNAME, "1", True) '加油站
        Else
            Call modUnset.GetVcmName(Me.TxtVANMID, Me.TxtVANMNAME, "2", True) '其他單位
        End If

    End Sub

    Protected Sub btnVCM_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If modUtil.IsStnRol(Me.Request) Then '加油站只能新增公出假
            Me.QryVcm.Show()
        Else
            Me.QryVcm2.Show()
        End If
    End Sub

    Private Function HODHOURCHECK(ByVal pEmpid As String, ByVal pDTST As String, ByVal pDTEN As String, ByVal pVANMID As String) As String ', ByVal pHours As Double
        '人事單位的核對，只有該單位才能維護資料
        Dim CHKFLAG As String = "N"
        Dim LSql As String
        Dim Date_St As Date
        Dim SDate As String
        Dim EDate As String
        '抓出 年度起迄日期
        Dim TmpD As Date = CDate(pDTST.Substring(0, 4) + "/" + pDTST.Substring(4, 2) + "/" + pDTST.Substring(6, 2))
        If TmpD.Month = 12 And TmpD.Day >= 21 Then
            SDate = (TmpD.Year).ToString
            SDate += "/12/21"
            EDate = (TmpD.Year + 1).ToString
            EDate += "/12/20"
        Else
            SDate = (TmpD.Year - 1).ToString
            SDate += "/12/21"
            EDate = (TmpD.Year).ToString
            EDate += "/12/20"
        End If

        If TmpD.Day >= 21 Then
            TmpD = TmpD.AddMonths(1)
            Date_St = CDate(TmpD.Year.ToString + "/" + Microsoft.VisualBasic.Right("00" + TmpD.Month.ToString, 2) + "/01")
        Else
            Date_St = TmpD '?
        End If

        Dim sConn As SqlClient.SqlConnection = Nothing
        Dim sCmd As SqlClient.SqlCommand = Nothing
        Dim sReader As SqlClient.SqlDataReader = Nothing
        'Dim sSql As String

        sConn = PA_GetSqlConnection()
        sCmd = New SqlClient.SqlCommand()
        sCmd.Connection = sConn

        Dim LTime1 As Double = 0 '可休時數
        Dim LTime2 As Double = 0 '已休時數
        Dim LTime3 As Double = 0 '?

        If pVANMID = "013" Then 'Case "013"  '補代休
            CHKFLAG = "Y"
            '查本年度的加班轉待休時數...->可休時數
            LSql = "SELECT SUM(LTIME1) AS LTIME1,SUM(LTIME2) AS LTIME2 FROM ("
            LSql += "SELECT ISNULL(SUM(OVTM_V_HOURS),0) AS LTIME1,0 AS LTIME2 FROM MP_HR.DBO.OVERTIME WITH (NOLOCK)"
            LSql += " WHERE OVTM_COMPID='01'"
            LSql += " AND OVTM_DATE BETWEEN '" & SDate & "' AND '" & EDate & "'"
            LSql += " AND OVTM_EMPLID='" & pEmpid & "'"
            LSql += " UNION ALL "
            LSql += "SELECT 0 AS LTIME1,ISNULL(SUM(VATM_HOURS),0) AS LTIME2  FROM MP_HR.DBO.VACATM WITH (NOLOCK)"
            'LSql += " AND VATM_COMPID='01'"
            LSql += " WHERE VATM_DATE_ST BETWEEN '" & SDate & "' AND '" & EDate & "'"
            LSql += " AND VATM_EMPLID='" & pEmpid & "'"
            LSql += " AND VATM_VANMID='013') A "

            sCmd.CommandText = LSql
            sReader = sCmd.ExecuteReader()
            sReader.Read()
            LTime1 = sReader.GetValue(0)
            LTime2 = sReader.GetValue(1)
            sReader.Close()
        ElseIf pVANMID = "000" Then '特休
            CHKFLAG = "Y"

            LSql = "SELECT SUM(LTIME1) AS LTIME1,SUM(LTIME2) AS LTIME2 FROM ("
            LSql += "SELECT ISNULL(VABL_ADD_HOURS,0) AS LTIME1,0 AS LTIME2 FROM MP_HR.DBO.VACA_BALANCE"
            LSql += " WHERE VABL_COMPID='01'"
            LSql += " AND VABL_EMPLID='" & pEmpid & "'"
            LSql += " AND VABL_YEAR='" & EDate.Substring(0, 4) & "'"
            LSql += " UNION ALL "
            LSql += "SELECT 0 AS LTIME1,ISNULL(SUM(VATM_HOURS),0) AS LTIME2"
            LSql += " FROM MP_HR.DBO.VACATM WITH (NOLOCK) "
            LSql += " WHERE VATM_DATE_ST BETWEEN '" & SDate & "' AND '" & EDate & "'"
            LSql += " AND VATM_COMPID='01'"
            LSql += " AND VATM_EMPLID='" & pEmpid & "'"
            LSql += " AND VATM_VANMID='000') A"
            sCmd.CommandText = LSql
            sReader = sCmd.ExecuteReader()
            sReader.Read()
            LTime1 = sReader.GetValue(0)
            LTime2 = sReader.GetValue(1)
            sReader.Close()
        ElseIf (pVANMID = "004") Or (pVANMID = "005") Or (pVANMID = "006") Then '婚假,喪假,產假
            CHKFLAG = "Y"
            'LSql = "SELECT TOP 1 FUNM_MAX_DAYS FROM MP_HR.DBO.FUNERAL_M WITH (NOLOCK) "
            LSql = "SELECT TOP 1 FUNM_MAX_DAYS*8 ,CONVERT(VARCHAR(10),FUNM_ST_DATE,111) FUNM_ST_DATE,CONVERT(VARCHAR(10),FUNM_EN_DATE,111) FUNM_EN_DATE"
            LSql += " FROM MP_HR.DBO.FUNERAL_M WITH (NOLOCK) "
            LSql += " WHERE FUNM_EMPLID = '" & pEmpid & "'"
            LSql += " AND FUNM_VANMID   = '" & pVANMID & "'"
            LSql += " ORDER BY FUNM_ST_DATE DESC "
            sCmd.CommandText = LSql
            sReader = sCmd.ExecuteReader()
            If sReader.Read() Then
                Dim FuneralSdate As String = ""
                Dim FuneralEdate As String = ""
                LTime1 = sReader.GetValue(0)
                FuneralSdate = sReader.GetString(1)
                FuneralEdate = sReader.GetString(2)
                sReader.Close()
                If FuneralSdate = "" Then
                    FuneralSdate = SDate
                End If
                If FuneralEdate = "" Then
                    FuneralSdate = EDate
                End If
                '
                LSql = "SELECT ISNULL(SUM(VATM_HOURS),0) VATM_HOURS"
                LSql += " FROM MP_HR.DBO.VACATM WITH (NOLOCK)"
                LSql += " WHERE VATM_DATE_ST BETWEEN '" & FuneralSdate & "' AND '" & FuneralEdate & "'"
                LSql += " AND VATM_COMPID='01'"
                LSql += " AND VATM_EMPLID='" & pEmpid & "'"
                LSql += " AND VATM_VANMID= '" & pVANMID & "'"
                sCmd.CommandText = LSql
                sReader = sCmd.ExecuteReader()
                sReader.Read()
                LTime2 = sReader.GetValue(0)
            End If
            sReader.Close()
        ElseIf (pVANMID = "002") Or (pVANMID = "003") Or (pVANMID = "007") Or (pVANMID = "014") Or (pVANMID = "015") Then 'Or (pVANMID = "006")
            CHKFLAG = "Y"
            '002=病假 003=事假 007=陪產假 014=生理假 015=家庭照顧假
            LSql = "SELECT SUM(LTIME1) AS LTIME1,SUM(LTIME2) AS LTIME2 FROM ("
            LSql += "SELECT ISNULL(SUM(VACA_MAX_DAYS),0)*8 AS LTIME1,0 AS LTIME2 FROM MP_HR.DBO.VACAMF WITH (NOLOCK)"
            LSql += " WHERE VACA_ID='" & pVANMID & "'"
            LSql += " UNION ALL "
            LSql += "SELECT 0 AS LTIME1,ISNULL(SUM(VATM_HOURS),0) AS LTIME2 FROM MP_HR.DBO.VACATM WITH (NOLOCK) "
            LSql += " WHERE VATM_DATE_ST BETWEEN '" & SDate & "' AND '" & EDate & "'"
            LSql += " AND VATM_COMPID='01'"
            LSql += " AND VATM_EMPLID='" & pEmpid & "'"
            If (pVANMID = "003") Or (pVANMID = "015") Then '事假要合併事假與家庭照顧假
                LSql += " AND VATM_VANMID IN ('003','015')) A"
            ElseIf (pVANMID = "002") Or (pVANMID = "014") Then '病假假要合併病假與生理假
                LSql += " AND VATM_VANMID IN ('002','014')) A"
            Else
                LSql += " AND VATM_VANMID='" & pVANMID & "') A"
            End If
            '
            sCmd.CommandText = LSql
            sReader = sCmd.ExecuteReader()
            sReader.Read()
            LTime1 = sReader.GetValue(0)
            LTime2 = sReader.GetValue(1)
            sReader.Close()
            '
            'modDB.InsertSignRecord("AziTest", "ok-H-5", My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "CLOSE_ACCYM= " + VACCYM, My.User.Name) 'AZITEST
            'ElseIf (pVANMID = "005") Then '喪假 '須對照親屬名稱...
            '    If Val(Me.txtVatm_FunmSeq.Text) = 0 Then
            '        Call Me.DialogProvider.GetControlValue(Me.txtVatm_FunmName)
            '    End If

            'LSql = "SELECT FUNM_MAX_DAYS,CONVERT(VARCHAR(10),FUNM_ST_DATE,111) FUNM_ST_DATE,CONVERT(VARCHAR(10),FUNM_EN_DATE,111) FUNM_EN_DATE"
            'LSql += " FROM FUNERAL_M"
            'LSql += " WITH (NOLOCK)"
            'LSql += " WHERE FUNM_EMPLID = '" & pEmpid & "'"
            'LSql += " AND FUNM_NAME     = '" & txtVatm_FunmName.Text.Trim & "'"
            'LSql += " AND FUNM_VANMID   = '005' "
            'LTable = giDB.GetDataTable(LSql)
            'Dim LTime1 As Double = 0
            'Dim FuneralSdate As String = ""
            'Dim FuneralEdate As String = ""
            'If LTable.Rows.Count <> 0 Then
            '    LTime1 = GI_Common.CSys_NullToZero(LTable.Rows(0).Item("FUNM_MAX_DAYS")) * 8
            '    FuneralSdate = GI_Common.CSys_NullToEmpty(LTable.Rows(0).Item("FUNM_ST_DATE"))
            '    FuneralEdate = GI_Common.CSys_NullToEmpty(LTable.Rows(0).Item("FUNM_EN_DATE"))
            'End If
            'If FuneralSdate = "" Then
            '    FuneralSdate = SDate
            'End If
            'If FuneralEdate = "" Then
            '    FuneralSdate = EDate
            'End If
            'LSql = "SELECT ISNULL(SUM(VATM_HOURS),0) VATM_HOURS"
            'LSql += " FROM VACATM,FUNERAL_M "
            'LSql += " WITH (NOLOCK)"
            'LSql += " WHERE VATM_DATE_ST BETWEEN '" & FuneralSdate & "' AND '" & FuneralEdate & "'"
            'LSql += " AND VATM_COMPID  = '" & MSys_Member.MSys_CompanyID & "'"
            'LSql += " AND VATM_EMPLID  = '" & txtVatm_EmplId.Text.Trim & "'"
            'LSql += " AND VATM_VANMID  = '005'"
            'LSql += " AND VATM_EMPLID  = FUNM_EMPLID"
            'LSql += " AND FUNM_NAME    = '" & txtVatm_FunmName.Text & "'"
            'LSql += " AND FUNM_VANMID  = '005'"
            'LSql += " AND VATM_FUNMSEQ = FUNM_SEQ "

            '    LTable = giDB.GetDataTable(LSql)
            '    Dim LTime2 As Double = 0
            '    If Not LTable.Rows.Count = 0 Then
            '        LTime2 = GI_Common.CSys_NullToZero(LTable.Rows(0).Item("VATM_HOURS"))
            '    End If
            '    txtTime1.Text = LTime1
            '    txtTime2.Text = LTime2
            '    txtTime3.Text = LTime1 - LTime2
        Else '其他假別
            LTime1 = 0
            LSql = "SELECT ISNULL(SUM(VATM_HOURS),0) VATM_HOURS"
            LSql += " FROM MP_HR.DBO.VACATM"
            LSql += " WITH (NOLOCK)"
            LSql += " WHERE VATM_DATE_ST BETWEEN '" & SDate & "' AND '" & EDate & "'"
            LSql += " AND VATM_COMPID='01'"
            LSql += " AND VATM_EMPLID='" & pEmpid & "'"
            LSql += " AND VATM_VANMID='" & pVANMID & "'"
            sCmd.CommandText = LSql
            sReader = sCmd.ExecuteReader()
            sReader.Read()
            LTime2 = sReader.GetValue(0)
            sReader.Close()
        End If
        '
        sConn.Close()
        sCmd.Dispose() '手動釋放資源
        sConn.Dispose()
        sConn = Nothing '移除指標
        '
        Return Format(LTime1, "0000.00") & Format(LTime2, "0000.00") & CHKFLAG
    End Function

    Public Sub VATMHOUR_Calculate(ByVal VMode As String)
        '20150304 add Vmode '1' = 強制重新計算
        Dim sValidOK As Boolean = True
        Dim CHKRSTR As String = ""
        Dim sMsg As String = ""
        '
        'modDB.InsertSignRecord("AziTest", "VATMHOUR_Calculate", My.User.Name) 'AZITEST
        'If LEditmode.Text = "Y" Then
        '    modDB.InsertSignRecord("AziTest", "Editmode is true", My.User.Name) 'AZITEST
        'Else
        '    modDB.InsertSignRecord("AziTest", "editmode is false", My.User.Name) 'AZITEST
        'End If

        '
        If (VMode = "1") Or (LEditmode.Text = "Y") And ((Me.txtVATMHOUR.Text.Trim = "") Or ((Me.txtVATMHOUR.Text.Trim = "0"))) Then
            ' modDB.InsertSignRecord("AziTest", "VATMHOUR_Calculate_editmode", My.User.Name) 'AZITEST
            '請假日期檢查
            If sValidOK Then
                CHKRSTR = modUnset.PADATECHK(txtVATMDATEST.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "請假起始日期錯誤!請檢查!"
                Else
                    txtVATMDATEST.Text = CHKRSTR.Substring(1, 8)
                End If
            End If

            '請假迄止日期檢查
            If sValidOK Then
                CHKRSTR = modUnset.PADATECHK(txtVATMDATEEN.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "請假迄止日期錯誤!請檢查!"
                Else
                    txtVATMDATEEN.Text = CHKRSTR.Substring(1, 8)
                End If
            End If

            '起始時間檢查
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(txtVATMTIMEST.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "起始時間錯誤!請檢查!"
                Else
                    txtVATMTIMEST.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '結束時間檢查
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(txtVATMTIMEEN.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "結束時間錯誤!請檢查!"
                Else
                    txtVATMTIMEEN.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '日期合理性檢查
            If sValidOK Then
                If (txtVATMDATEST.Text + txtVATMTIMEST.Text) >= (txtVATMDATEEN.Text + txtVATMTIMEEN.Text) Then
                    sValidOK = False : sMsg = sMsg & "\n * 請假起迄時間錯誤!請檢查!"
                End If
            End If
            '
            If sValidOK Then
                'modDB.InsertSignRecord("AziTest", "VATMHOUR_Calculate_editmode_ok", My.User.Name) 'AZITEST
                Dim VHOURS As Double = COUNTHOURS(Me.txtVATMDATEST.Text, Me.txtVATMTIMEST.Text, Me.txtVATMDATEEN.Text, Me.txtVATMTIMEEN.Text)
                If Me.txtVATMTIMEST.Text.Substring(0, 2) <= "12" And Me.txtVATMTIMEEN.Text.Substring(0, 2) >= "13" Then
                    If Not modUtil.IsStnRol(Me.Request) Then '非加油站資料，過午休時間的話一律扣除1小時。
                        VHOURS = VHOURS - 1
                    End If
                End If
                '
                'modDB.InsertSignRecord("AziTest", "VATMHOUR_Calculate_editmode_print", My.User.Name) 'AZITEST
                Me.txtVATMHOUR.Text = VHOURS.ToString
            Else
                showMsg(Me.Page, "訊息(請修正錯誤欄位)", sMsg)
            End If
            '
        End If
    End Sub

    '20140227 在此模式下沒有作用
    'Protected Sub txtVATMDATEST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVATMDATEST.TextChanged
    'modDB.InsertSignRecord("AziTest", "txtVATMDATEST_TextChanged", My.User.Name)
    'txtVATMTIMEST.Text = "test"
    'modDB.InsertSignRecord("AziTest", "txtVATMDATEST_TextChanged", My.User.Name)
    'txtVATMDATEEN.Text = txtVATMDATEST.Text
    'modUtil.showMsg(Me.Page, "訊息", "txtVATMDATEEN=" & txtVATMDATEEN.Text)

    'If txtVATMDATEEN.Text.Trim = "" Then
    '    modDB.InsertSignRecord("AziTest", "Me.txtVATMDATEEN.Text is empty!", My.User.Name)
    '    Me.txtVATMDATEEN.Text = Me.txtVATMDATEST.Text.Trim
    '    modDB.InsertSignRecord("AziTest", "Me.txtVATMDATEEN.Text = " & Me.txtVATMDATEST.Text.Trim, My.User.Name)
    'Else
    '    modDB.InsertSignRecord("AziTest", "Me.txtVATMDATEEN.Text is not empty!", My.User.Name)
    '    modDB.InsertSignRecord("AziTest", "Me.txtVATMDATEEN.Text = " & Me.txtVATMDATEEN.Text, My.User.Name)
    'End If
    'End Sub

    'Protected Sub txtVATMTIMEEN_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVATMTIMEEN.TextChanged
    '    modDB.InsertSignRecord("AziTest", "txtVATMTIMEEN_TextChanged", My.User.Name)
    '    Dim TIME_S As Date
    '    Dim TIME_E As Date

    '    TIME_S = CDate(txtVATMDATEST.Text.Substring(0, 4) + "/" + txtVATMDATEST.Text.Substring(4, 2) + "/" + txtVATMDATEST.Text.Substring(6, 2) + " " + txtVATMTIMEST.Text.Substring(0, 2) + ":" + txtVATMTIMEST.Text.Substring(2, 2))
    '    TIME_E = CDate(txtVATMDATEEN.Text.Substring(0, 4) + "/" + txtVATMDATEEN.Text.Substring(4, 2) + "/" + txtVATMDATEEN.Text.Substring(6, 2) + " " + txtVATMTIMEEN.Text.Substring(0, 2) + ":" + txtVATMTIMEEN.Text.Substring(2, 2))
    '    Dim VHours As Double = 0
    '    If txtVATMDATEST.Text <> txtVATMDATEEN.Text Then
    '        If TIME_E < TIME_S Then Exit Sub
    '        Dim TDate As Date
    '        TDate = CDate(txtVATMDATEST.Text.Substring(0, 4) + "/" + txtVATMDATEST.Text.Substring(4, 2) + "/" + txtVATMDATEST.Text.Substring(6, 2))
    '        Do While TDate.Date < TIME_E.Date
    '            VHours += 8
    '            TDate = TDate.AddDays(1)
    '        Loop
    '    Else
    '        VHours = 0
    '    End If

    '    If (txtVATMHOUR.Text.Trim = "0") Or (txtVATMHOUR.Text.Trim = "") Then
    '        If TIME_S.Hour <= 12 And TIME_E.Hour >= 13 Then
    '            VHours += Math.Round((TIME_E - TIME_S).TotalMinutes / 60, 2)
    '            Me.txtVATMHOUR.Text = VHours
    '        Else
    '            VHours += Math.Round((TIME_E - TIME_S).TotalMinutes / 60, 2)
    '            Me.txtVATMHOUR.Text = VHours
    '        End If
    '    End If

    '    txtVATMHOUR.Text = VHours.ToString

    'End Sub

    Protected Sub txtVATMDATEST_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVATMDATEST.TextChanged
        'modDB.InsertSignRecord("AziTest", "txtVATMDATEST_TextChanged", My.User.Name) 'AZITEST
        'VATMHOUR_Calculate()
    End Sub


    Protected Sub BtnHourCal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnHourCal.Click
        VATMHOUR_Calculate("1")
    End Sub
End Class
