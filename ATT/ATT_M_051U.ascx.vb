'******************************************************************************************************
'* 程式：ATT_M_051U 排班資料維護
'* //版次：2013/11/01(VER1.01)：新開發
'******************************************************************************************************

Imports modUtil
Imports modUnset
Imports System.Data


Partial Class ATT_M_051U
    Inherits System.Web.UI.UserControl

    'Private cWriteLog As Boolean = False '* 上線時須設成 False
    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Public Sub InitScreen(ByVal pMode As FormViewMode)
        Me.btnStnID.Visible = False '暫不開放
        'Me.btnCLSID.Visible = False '暫不開放

        If pMode = FormViewMode.ReadOnly Then
            SetAllReadOnly(Me)
            Me.btnStnID.Visible = False
            Me.btnEmp.Visible = False
            'Me.btnCLSID.Visible = False
        Else
            SetObjReadOnly(Me, "txtStnID")
            SetObjReadOnly(Me, "txtStnName")
            SetObjReadOnly(Me, "txtEmpName")
            'SetObjReadOnly(Me, "txtRTSTTIME")
            'SetObjReadOnly(Me, "txtRTEDTIME")
            SetObjReadOnly(Me, "txtWORKHOUR")
            Me.txtEmpID.Attributes.Add("onkeypress", "return CheckKeyAZ09()")

            If pMode = FormViewMode.Insert Then
                '只考慮新增模式()
                Me.txtSTNID.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                Me.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                Me.txtSHSTDATE.Text = Now.AddDays(-1).ToString("yyyyMMdd") '以前一天日期為預設值
                Me.txtWORKHOUR.Visible = False '20131105_新增時時數不顯示，因為時數是存檔才做計算
                modUtil.SetObjFocus(Me.Page, Me.txtEmpID)
            ElseIf pMode = FormViewMode.Edit Then '修改模式()
                SetObjReadOnly(Me, "txtEmpID")
                SetObjReadOnly(Me, "txtSHSTDATE")
                Me.txtSTNID.Text = Request("STNID").Trim
                If Me.txtSTNID.Text = "" Then
                    Me.txtSTNID.Text = Request.Cookies("STNID").Value
                End If
                'HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                Me.vSHSTDATE.Text = Me.txtSHSTDATE.Text
                Me.vSHSTTIME.Text = Me.txtSHSTTIME.Text
                Me.vSHEDTIME.Text = Me.txtSHEDTIME.Text
                modUtil.SetObjFocus(Me.Page, Me.txtSHSTTIME)
            End If
            '
            CLSTIME.Items.Clear()
            Dim sConn As SqlClient.SqlConnection = Nothing
            Dim sCmd As SqlClient.SqlCommand = Nothing
            Dim sSql As String

            sConn = New SqlClient.SqlConnection
            sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
            sConn.Open()

            sCmd = New SqlClient.SqlCommand()
            sCmd.Connection = sConn

            Dim sReader As SqlClient.SqlDataReader = Nothing
            'sCmd = New SqlClient.SqlCommand()
            'sCmd.Connection = sConn
            '------------------------------------- 排班代碼資料 
            sSql = "SELECT CLSID,STTIME,EDTIME,NFLAG FROM CLSTIME WHERE STNID='" & Me.txtSTNID.Text & "'" _
                 & " ORDER BY CLSID"

            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()
            'sReader.Read()
            CLSTIME.Items.Add("")
            While sReader.Read
                If sReader.GetString(3) = "Y" Then
                    CLSTIME.Items.Add(FillStr(sReader.GetString(0), 4, "_") & " " & sReader.GetString(1) & "~" & sReader.GetString(2) & " " & sReader.GetString(3))
                Else
                    CLSTIME.Items.Add(FillStr(sReader.GetString(0), 4, "_") & " " & sReader.GetString(1) & "~" & sReader.GetString(2) & " N")
                End If
            End While
            sReader.Close()
            'CLSTIME.SelectedIndex = -1
            CLSTIME.Text = ""
            '
        End If
    End Sub

    '******************************************************************************************************
    '* DB 更新前置檢核處理
    '******************************************************************************************************
    Private Function CheckData(ByVal pMode As FormViewMode) As Boolean
        Dim sValidOK As Boolean = True
        Dim sMsg As String = ""
        Dim CHKRSTR As String = ""
        'modDB.InsertSignRecord("AziTest", "ATT_M_051U_CheckData", My.User.Name) 'AZITEST

        Try
            If pMode = FormViewMode.Insert Then '* 新增模式
                sValidOK = sValidOK And CheckLength(Me, "txtEMPID", 8, sMsg, "員工編號")
            End If
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtSHSTDATE", sMsg, "起始日期")
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtSHSTTIME", sMsg, "起始時間")
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtSHEDTIME", sMsg, "結束時間")

            '排班日期檢查
            If sValidOK Then
                CHKRSTR = modUnset.PADATECHK(txtSHSTDATE.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "排班日期錯誤!請檢查!"
                Else
                    txtSHSTDATE.Text = CHKRSTR.Substring(1, 8)
                End If
            End If

            '起始時間檢查
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(txtSHSTTIME.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "起始時間錯誤!請檢查!"
                Else
                    txtSHSTTIME.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '結束時間檢查
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(txtSHEDTIME.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "結束時間錯誤!請檢查!"
                Else
                    txtSHEDTIME.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            If sValidOK Then
                If txtSHSTTIME.Text = txtSHEDTIME.Text Then
                    sValidOK = False : sMsg = sMsg & "\n * 排班起迄時間錯誤!請檢查!"
                End If
            End If

            If sValidOK Then '再作鎖檔日期檢核
                'Dim LCKDT As String = GET_LOCKYM("1", txtSTNID.Text) & "20"
                Dim LCKDT As String = GET_LOCKDT(Me.txtSTNID.Text)
                'modDB.InsertSignRecord("AziTest", "LCKDT= " + LCKDT + " vSHSTDATE.Text+vSHSTTIME.Text=" + vSHSTDATE.Text + vSHSTTIME.Text + " txtSHSTDATE.Text+txtSHSTTIME.Text=" + txtSHSTDATE.Text + txtSHSTTIME.Text, My.User.Name) 'AZITEST
                'If (Me.txtSHSTDATE.Text & txtSHSTTIME.Text) < (LCKDT & "2300") Then
                If Me.txtSHSTDATE.Text <= LCKDT Then
                    sValidOK = False
                    sMsg = "資料週期已鎖或過帳，不可異動!"
                    'ElseIf (pMode = FormViewMode.Edit) And ((vSHSTDATE.Text & vSHSTTIME.Text) < (LCKDT & "2300")) Then
                ElseIf (pMode = FormViewMode.Edit) And (vSHSTDATE.Text <= LCKDT) Then
                    sValidOK = False
                    sMsg = "原資料月份已鎖或過帳，不可異動!"
                End If
            End If

            '檢核午休起始時間
            If sValidOK And (txtRTSTTIME.Text <> "") Then

                CHKRSTR = modUnset.PATIMECHK(txtRTSTTIME.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = sMsg & "\n * 午休起始時間錯誤!請檢查!"
                Else
                    txtRTSTTIME.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '檢核午休終止時間
            If sValidOK And (txtRTEDTIME.Text <> "") Then

                CHKRSTR = modUnset.PATIMECHK(txtRTEDTIME.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = sMsg & "\n * 午休迄止時間錯誤!請檢查!"
                Else
                    txtRTEDTIME.Text = CHKRSTR.Substring(1, 4)
                End If
            End If
            '
            '午休時間邏輯檢查
            If sValidOK And ((txtRTSTTIME.Text <> "") Or (txtRTEDTIME.Text <> "")) Then
                If (txtRTSTTIME.Text = "") Or (txtRTEDTIME.Text = "") Then
                    sValidOK = False : sMsg = sMsg & "\n * 午休起迄時間錯誤!請檢查!"
                Else
                    Dim RTDATEST, RTDATEED As String
                    RTDATEST = txtSHSTDATE.Text
                    RTDATEED = txtSHSTDATE.Text
                    'modDB.InsertSignRecord("AziTest", "ATT_M_051U_OK-1", My.User.Name) 'AZITEST
                    If ((txtRTEDTIME.Text < txtRTSTTIME.Text) And (txtRTEDTIME.Text <> "0000")) Then '休息跨日
                        sValidOK = False : sMsg = sMsg & "\n * 休息迄止時間錯誤!不可跨日!請檢查!"
                        'modDB.InsertSignRecord("AziTest", "ATT_M_051U_OK-2", My.User.Name) 'AZITEST
                    Else
                        'modDB.InsertSignRecord("AziTest", "ATT_M_051U_OK-3", My.User.Name) 'AZITEST
                        If txtSHSTTIME.Text > txtRTSTTIME.Text Then
                            RTDATEST = Format(CDate(Format(CInt(RTDATEST), "0000/00/00")).AddDays(1), "yyyyMMdd")
                            RTDATEED = RTDATEST
                        End If

                        If txtRTEDTIME.Text < txtRTSTTIME.Text Then
                            RTDATEED = Format(CDate(Format(CInt(RTDATEST), "0000/00/00")).AddDays(1), "yyyyMMdd")
                        End If
                        '
                        'modDB.InsertSignRecord("AziTest", "ATT_M_051U_SH:" + (txtSHSTDATE.Text + txtSHSTTIME.Text) + "~" + (RTDATEED + txtRTEDTIME.Text), My.User.Name) 'AZITEST
                        'modDB.InsertSignRecord("AziTest", "ATT_M_051U_RT:" + (RTDATEST + txtRTSTTIME.Text) + "~" + (RTDATEED + txtRTEDTIME.Text), My.User.Name) 'AZITEST
                        If (RTDATEST + txtRTSTTIME.Text) < (txtSHSTDATE.Text + txtSHSTTIME.Text) Or (RTDATEED + txtRTEDTIME.Text) > (txtSHEDDATE.Text + txtSHEDTIME.Text) Then
                            sValidOK = False : sMsg = sMsg & "\n * 午休起迄時間錯誤!超出排班時間範圍!請檢查!"
                        ElseIf (RTDATEST + txtRTSTTIME.Text) >= (RTDATEED + txtRTEDTIME.Text) Then
                            sValidOK = False : sMsg = sMsg & "\n * 午休起迄時間錯誤!請檢查!"
                        End If
                    End If
                End If
            End If
            '
            If sValidOK And ((pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit)) Then

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
                '
                '-------------------------------------* 找出是否有重複資料
                sSql = "SELECT EMPID,SHSTDATE,SHSTTIME,SHEDDATE,SHEDTIME FROM SCHEDM " _
                     & " WHERE EMPID='" & txtEmpID.Text & "'" _
                     & "   AND ((SHEDDATE+SHEDTIME <='" & txtSHEDDATE.Text & txtSHEDTIME.Text & "'" _
                     & "     AND SHEDDATE+SHEDTIME > '" & txtSHSTDATE.Text & txtSHSTTIME.Text & "') " _
                     & "    OR  (SHSTDATE+SHSTTIME < '" & txtSHEDDATE.Text & txtSHEDTIME.Text & "'" _
                     & "     AND SHSTDATE+SHSTTIME >='" & txtSHSTDATE.Text & txtSHSTTIME.Text & "') " _
                     & "    OR  (SHSTDATE+SHSTTIME <='" & txtSHEDDATE.Text & txtSHEDTIME.Text & "'" _
                     & "     AND SHEDDATE+SHEDTIME > '" & txtSHSTDATE.Text & txtSHSTTIME.Text & "'))"

                If pMode = FormViewMode.Edit Then
                    sSql = sSql & " AND (SHSTDATE+SHSTTIME+SHEDTIME)<>'" + vSHSTDATE.Text + vSHSTTIME.Text + vSHEDTIME.Text + "'"
                End If

                sCmd.CommandText = sSql
                sReader = sCmd.ExecuteReader()
                sReader.Read()
                '
                If (sReader.HasRows) Then
                    sValidOK = False
                    sMsg = "已有資料重複:(排班單位:" & txtSTNID.Text & " 日期:" & sReader.GetString(1) & " 時間: " & sReader.GetString(2) & " ~ " & sReader.GetString(4) & " )"
                    'modUtil.showMsg(Me.Page, "錯誤訊息", "已有資料1重複:(排班單位:" & txtSTNID.Text & " 日期:" & sReader.GetString(1) & " 時間:" & sReader.GetString(2) & "~" & sReader.GetString(4))
                End If
                '20140407 新增建教生檢查
                sReader.Close()
                '20140403 增加建教生與工作時數檢查
                'getSHEDDATE 取兩周日期區間
                Dim WDateSt As String = ""
                Dim WDateEn As String = ""
                Dim CHOURCHECK As Boolean = True
                Dim EMPLCOLL As String = ""
                Dim WORKDAY, WORKHOUR As Integer
                Dim VWORKHR As Double = modUnset.COUNTHOURS(Me.txtSHSTTIME.Text, Me.txtSHEDTIME.Text)

                If txtSHEDDATE.Text.Substring(6, 2) >= "21" Then
                    WDateSt = txtSHEDDATE.Text.Substring(0, 6) + "21"
                    WDateEn = Format(CDate(Format(CInt(WDateSt), "0000/00/00")).AddDays(13), "yyyyMMdd")
                Else
                    WDateSt = Format(CDate(Format(CInt(txtSHEDDATE.Text), "0000/00/00")).AddMonths(-1), "yyyyMM") + "21"
                    WDateEn = Format(CDate(Format(CInt(WDateSt), "0000/00/00")).AddDays(13), "yyyyMMdd")
                    If txtSHEDDATE.Text > WDateEn Then
                        WDateSt = Format(CDate(Format(CInt(WDateEn), "0000/00/00")).AddDays(1), "yyyyMMdd")
                        WDateEn = Format(CDate(Format(CInt(WDateSt), "0000/00/00")).AddDays(13), "yyyyMMdd")
                        If txtSHEDDATE.Text > WDateEn Then
                            CHOURCHECK = False
                        End If
                    End If
                End If

                If sValidOK Then
                    sSql = "SELECT EMPL_COLL_RELATION,SUM(WORKDAY) AS WORKDAY,SUM(WORKHOUR) AS WORKHOUR" _
                                             & "  FROM (SELECT COUNT(*) AS WORKDAY,0 AS WORKHOUR  " _
                                             & "      FROM (SELECT DISTINCT SHEDDATE FROM SCHEDM " _
                                             & "             WHERE EMPID='" & txtEmpID.Text & "' AND SHSTDATE>='" & WDateSt & "' AND SHSTDATE<='" & WDateEn & "'" _
                                             & "               AND (SHSTDATE+SHSTTIME+SHEDTIME)<>'" + vSHSTDATE.Text + vSHSTTIME.Text + vSHEDTIME.Text + "') A" _
                                             & " UNION ALL " _
                                             & " SELECT 0 AS WORKDAY,SUM(WORKHOUR) AS WORKHOUR FROM SCHEDM " _
                                             & "            WHERE EMPID='" & txtEmpID.Text & "' AND SHSTDATE>='" & WDateSt & "' AND SHSTDATE<='" & WDateEn & "'" _
                                             & "              AND (SHSTDATE+SHSTTIME+SHEDTIME)<>'" + vSHSTDATE.Text + vSHSTTIME.Text + vSHEDTIME.Text + "') A" _
                                             & " INNER JOIN MP_HR.DBO.EMPLOYEE B ON B.EMPL_ID='" & txtEmpID.Text & "'" _
                                             & " GROUP BY EMPL_COLL_RELATION,EMPL_ARV_DATE,EMPL_LEV_DATE "
                    sCmd.CommandText = sSql
                    sReader = sCmd.ExecuteReader()
                    sReader.Read()
                    EMPLCOLL = sReader.GetString(0)
                    WORKDAY = sReader.GetInt32(1)
                    WORKHOUR = sReader.GetDouble(2)
                    sReader.Close()
                    'modDB.InsertSignRecord("AziTest", "EMPLCOLL= " + EMPLCOLL + " WORKDAY=" + WORKDAY.ToString + " WORKHOUR=" + WORKHOUR.ToString + "  VWORKHR=" + VWORKHR.ToString, My.User.Name) 'AZITEST
                    '
                    If EMPLCOLL = "1" Then '建教生
                        If (txtSHSTTIME.Text < "0600") Or (txtSHEDTIME.Text > "2200") Or (txtSHSTDATE.Text <> txtSHEDDATE.Text) Then
                            sValidOK = False : sMsg = sMsg & "\n 建教生排班時間限制在 06:00 ~ 22:00 之間"
                        End If
                        '
                        If CHOURCHECK Then '前28天才作天數與總時數的檢查
                            If sValidOK Then
                                If (WORKDAY + 1) > 12 Then
                                    sValidOK = False : sMsg = sMsg & "\n 建教生 " & WDateSt & " ~ " & WDateEn & " 期間不可排班超過12天"
                                End If
                            End If
                            '
                            If sValidOK Then
                                If (WORKHOUR + VWORKHR) > 84 Then
                                    sValidOK = False : sMsg = sMsg & "\n 建教生 " & WDateSt & " ~ " & WDateEn & " 期間不可排班超過 84小時"
                                End If
                            End If
                        End If
                        '
                        If sValidOK Then
                            If VWORKHR > 12 Then
                                sValidOK = False : sMsg = sMsg & "\n 建教生 不可排班超過 12 小時"
                            End If
                        End If
                    Else
                        If CHOURCHECK Then '前28天才作天數與總時數的檢查
                            If sValidOK Then
                                If (WORKDAY + 1) > 12 Then
                                    sValidOK = False : sMsg = sMsg & "\n " & WDateSt & " ~ " & WDateEn & " 期間不可排班超過12天"
                                End If
                            End If
                            '
                            If sValidOK Then
                                If (WORKHOUR + VWORKHR) > 84 Then
                                    sMsg = sMsg & "\n " & WDateSt & " ~ " & WDateEn & " 期間不應排班超過 84小時"
                                End If
                            End If
                        End If
                        '
                        If sValidOK Then
                            If VWORKHR > 12 Then
                                sValidOK = False : sMsg = sMsg & "\n 不可排班超過 12 小時"
                            End If
                        End If
                    End If
                End If
            End If
            '
            If Not sValidOK Then
                showMsg(Me.Page, "訊息(請修正錯誤欄位)", sMsg)
            Else
                If sMsg <> "" Then modUtil.showMsg(Me.Page, "訊息", "請注意：" & sMsg)
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
        'Dim LCKDT As String = GET_LOCKYM("1", txtSTNID.Text) & "20"
        Dim LCKDT As String = GET_LOCKDT(Me.txtSTNID.Text)
        Dim VHOURS As Double = 0

        'modDB.InsertSignRecord("AziTest", "ATT_M_051U in", My.User.Name) 'AZITEST

        'modUtil.showMsg(Me.Page, "鎖檔日期", LCKDT)
        If (pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit) Then
            'modDB.InsertSignRecord("AziTest", "ATT_M_051U: edit mode", My.User.Name) 'AZITEST
            txtSHEDDATE.Text = Me.txtSHSTDATE.Text
            If Me.txtSHSTTIME.Text > Me.txtSHEDTIME.Text Then
                txtSHEDDATE.Text = Format(CDate(Format(CInt(Me.txtSHSTDATE.Text), "0000/00/00")).AddDays(1), "yyyyMMdd")
            End If
            '
            If Not CheckData(pMode) Then '* 取消新增/修改
                e.Cancel = True
                Exit Sub
            End If
            '
            VHOURS = modUnset.COUNTHOURS(Me.txtSHSTTIME.Text, Me.txtSHEDTIME.Text)

            If Me.txtRTSTTIME.Text.Trim <> "" Then
                VHOURS = VHOURS - modUnset.COUNTHOURS(Me.txtRTSTTIME.Text, Me.txtRTEDTIME.Text)
            End If

            Me.txtWORKHOUR.Text = VHOURS.ToString
        Else '20151109 增加刪除的檢查
            If Me.txtSHSTDATE.Text <= LCKDT Then
                showMsg(Me.Page, "訊息", "資料週期已鎖或過帳，不可異動!")
                e.Cancel = True
                Exit Sub
            Else
                modDB.InsertSignRecord("AziTest", "ATT_M_051U: DELETE SCHEDM:" & Me.txtEmpID.Text & "-" & Me.txtSHSTDATE.Text & "-" & Me.txtSHSTTIME.Text, My.User.Name) 'AZITEST
                '20151201 增加刪除資料的update時間STAMP
                Dim RESETLDTST, RESETLDTEN As String
                Dim CSQL As String
                CSQL = "SELECT ISNULL(MIN(STDATE),'') AS STDATE,ISNULL(MAX(EDDATE),'') AS EDDATE FROM MECHSYA" _
                     & " WHERE STNID='000000' AND ACCYM>='201510' AND LOCKCODE='N'" _
                     & "   AND '" & Me.txtSHSTDATE.Text & "' BETWEEN STDATE AND EDDATE"
                Dim tt As DataTable
                'LOGSQL(CSQL) 'AZITEST
                tt = get_DataTable(CSQL)
                RESETLDTST = tt.Rows(0).Item("STDATE")
                RESETLDTEN = tt.Rows(0).Item("EDDATE")
                'modDB.InsertSignRecord("AziTest", "ATT_M_051U: LCKDT=" & LCKDT & " RESETLDTST=" & RESETLDTST & " RESETLDTEN=" & RESETLDTEN, My.User.Name) 'AZITEST
                '
                If (RESETLDTST <> "") And (RESETLDTEN <> "") And (RESETLDTEN > LCKDT) Then
                    If (LCKDT >= RESETLDTST) Then
                        RESETLDTST = Convert.ToDateTime(LCKDT.Substring(0, 4) + "/" + LCKDT.Substring(4, 2) + "/" + LCKDT.Substring(6, 2)).AddDays(1).ToString("yyyy/MM/dd")
                    End If
                    '
                    CSQL = "UPDATE SCHEDM SET INDATE='" & DateTime.Now.ToString("yyyyMMdd") & "',INTIME='" & DateTime.Now.ToString("HHmmss") & "'" _
                         & " WHERE EMPID='" & Me.txtEmpID.Text & "'" _
                         & "   AND SHSTDATE>='" & RESETLDTST & "' AND SHSTDATE<='" & RESETLDTEN & "'"
                    'LOGSQL(CSQL) 'AZITEST
                    EXE_SQL(CSQL)
                End If
            End If
        End If
        '
        Try

            With e.Command
                .Parameters("@STNID").Value = Me.txtSTNID.Text
                .Parameters("@EMPID").Value = Me.txtEmpID.Text
                .Parameters("@SHSTDATE").Value = Me.txtSHSTDATE.Text
                .Parameters("@SHSTTIME").Value = Me.txtSHSTTIME.Text
                If (pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit) Then
                    .Parameters("@SHEDDATE").Value = Me.txtSHEDDATE.Text
                    .Parameters("@SHEDTIME").Value = Me.txtSHEDTIME.Text
                    .Parameters("@WORKHOUR").Value = Me.txtWORKHOUR.Text
                    .Parameters("@RTSTTIME").Value = Me.txtRTSTTIME.Text
                    .Parameters("@RTEDTIME").Value = Me.txtRTEDTIME.Text
                    .Parameters("@inuser").Value = Me.Page.User.Identity.Name
                    .Parameters("@indate").Value = DateTime.Now.ToString("yyyyMMdd")
                    .Parameters("@intime").Value = DateTime.Now.ToString("HHmmss")
                    .Parameters("@CLASSID").Value = ""
                    '-------------------------------------* 刪除舊資料
                    If pMode = FormViewMode.Edit Then
                        .Parameters("@VSTDATE").Value = Me.vSHSTDATE.Text
                        .Parameters("@VSTTIME").Value = Me.vSHSTTIME.Text
                    End If
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

    Protected Sub chkNormalCls_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkNormalCls.CheckedChanged
        'azitemp 勾正常班 -> 午休時間=1200~1300
        If chkNormalCls.Checked Then
            txtRTSTTIME.Text = "1200"
            txtRTEDTIME.Text = "1300"
        Else
            txtRTSTTIME.Text = ""
            txtRTEDTIME.Text = ""
        End If
    End Sub

    Protected Sub txtEmpID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call modCharge.GetEmpName(Me.txtEmpID, Me.txtEmpName, True)
    End Sub

    Protected Sub btnEmp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryEmp.Show(Me.txtSTNID.Text)
    End Sub

    Protected Sub CLSTIME_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CLSTIME.SelectedIndexChanged

        txtSHSTTIME.Text = CLSTIME.SelectedItem.Text.Substring(5, 4)
        txtSHEDTIME.Text = CLSTIME.SelectedItem.Text.Substring(10, 4)
        If CLSTIME.SelectedItem.Text.Substring(15, 1) = "Y" Then '正常班
            chkNormalCls.Checked = True
            txtRTSTTIME.Text = "1200"
            txtRTEDTIME.Text = "1300"
        Else
            chkNormalCls.Checked = False
            txtRTSTTIME.Text = ""
            txtRTEDTIME.Text = ""
        End If
    End Sub

    Protected Sub CLSTIME_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CLSTIME.TextChanged
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

    '連接資料庫取得datatable
    Public Function get_DataTable(ByVal SQL1 As String) As System.Data.DataTable
        Dim CONNSTR1 As String = modUnset.PA_GetDataSource.ConnectionString
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

    '執行quary不回傳值
    Public Function EXE_SQL(ByVal TMPSQL1 As String) As Boolean
        'Dim AA As Integer
        Dim CONNSTR1 As String = modUnset.PA_GetDataSource.ConnectionString
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
End Class
