'******************************************************************************************************
'* 程式：ATT_M_053U 打卡資料維護
'* //版次：2013/05/31(VER1.01)：新開發
'******************************************************************************************************

Imports modUtil
Imports modUnset
Imports System.Data


Partial Class ATT_M_053U
    Inherits System.Web.UI.UserControl

    'Private cWriteLog As Boolean = False '* 上線時須設成 False
    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Public Sub InitScreen(ByVal pMode As FormViewMode)
        Me.btnStnID.Visible = False '暫不開放
        'modDB.InsertSignRecord("AziTest", "ATT_053U-1", My.User.Name)
        If (pMode = FormViewMode.ReadOnly) Or (pMode = FormViewMode.Edit) Then
            SetAllReadOnly(Me)
            Me.btnStnID.Visible = False
            Me.btnEmp.Visible = False
            If Not (pMode = FormViewMode.Edit) Then
                'modDB.InsertSignRecord("AziTest", "ATT_053U-2", My.User.Name)
                Label1.Visible = False
                txtEmpID2.Visible = False
                txtEmpName2.Visible = False
            Else
                'modDB.InsertSignRecord("AziTest", "ATT_053U-3", My.User.Name)
                'txtEmpID2.Enabled = True
                SetObjEnabled(txtEmpID2)
                'LCHGEMPID.Text = txtEmpID.Text '先儲存員工編號
                'modDB.InsertSignRecord("AziTest", "LCHGEMPID = " + LCHGEMPID.Text, My.User.Name)
            End If
        Else
            'modDB.InsertSignRecord("AziTest", "ATT_053U-B", My.User.Name)
            Label1.Visible = False
            txtEmpID2.Visible = False
            txtEmpName2.Visible = False
            '
            SetObjReadOnly(Me, "txtStnID")
            SetObjReadOnly(Me, "txtStnName")
            SetObjReadOnly(Me, "txtEmpName")
            Me.txtEmpID.Attributes.Add("onkeypress", "return CheckKeyAZ09()")

            If pMode = FormViewMode.Insert Then
                '只考慮新增模式()
                'Me.txtSTNID.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                'Me.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                Me.txtSTNID.Text = Request("STNID")
                Me.txtStnName.Text = Request("STNNAME")
                modUtil.SetObjFocus(Me.Page, Me.txtEmpID)

            End If
        End If
    End Sub

    '******************************************************************************************************
    '* DB 更新前置檢核處理
    '******************************************************************************************************
    Private Function CheckData(ByVal pMode As FormViewMode) As Boolean
        Dim sValidOK As Boolean = True
        Dim sMsg As String = ""
        Dim CHKRSTR As String = ""

        Try
            If pMode = FormViewMode.Insert Then '* 新增模式
                sValidOK = sValidOK And CheckLength(Me, "txtEMPID", 8, sMsg, "員工編號")
                txtFINTIME.Text = txtFINTIME.Text.Substring(0, 4) '20150820 新增打卡時間只取前4碼

                sValidOK = sValidOK And CheckNotEmpty(Me, "txtFINDATE", sMsg, "補登打卡日期")
                sValidOK = sValidOK And CheckNotEmpty(Me, "txtFINTIME", sMsg, "補登打卡時間")

                '補登日期檢查
                If sValidOK Then
                    CHKRSTR = modUnset.PADATECHK(txtFINDATE.Text)
                    If CHKRSTR.Substring(0, 1) <> "Y" Then
                        sValidOK = False : sMsg = "補登打卡日期錯誤!請檢查!"
                    Else
                        txtFINDATE.Text = CHKRSTR.Substring(1, 8)
                    End If
                End If

                '補登時間檢查
                If sValidOK Then
                    CHKRSTR = modUnset.PATIMECHK(txtFINTIME.Text)
                    If CHKRSTR.Substring(0, 1) <> "Y" Then
                        sValidOK = False : sMsg = "補登打卡時間錯誤!請檢查!"
                    Else
                        txtFINTIME.Text = CHKRSTR.Substring(1, 4)
                    End If
                End If
            End If
            '
            If sValidOK Then '再作鎖檔日期檢核
                'Dim LCKDT As String = GET_LOCKYM("1", txtSTNID.Text) & "20"
                Dim LCKDT As String = GET_LOCKDT(txtSTNID.Text)
                If Me.txtFINDATE.Text <= LCKDT Then
                    sValidOK = False
                    sMsg = "資料週期已鎖或過帳，不可異動!"
                End If
            End If
            '
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
                sSql = "SELECT EMPID,SHSTDATE,SHSTTIME,SHEDDATE,SHEDTIME,CLASSID,WORKHOUR FROM SCHEDM" _
                        & " WHERE EMPID='" & txtEmpID.Text & "'" _
                        & "   AND (SHSTDATE='" & txtFINDATE.Text & "'" _
                        & "     OR SHEDDATE='" & txtFINDATE.Text & "')" _
                        & " ORDER BY SHSTDATE,SHSTTIME "
                sCmd.CommandText = sSql
                sReader = sCmd.ExecuteReader()
                Dim TIMECHK As Boolean = False
                'modDB.InsertSignRecord("AziTest", "OK-1", My.User.Name)
                If (sReader.HasRows) Then
                    Do While (sReader.Read())
                        'modDB.InsertSignRecord("AziTest", "FIN=" & txtFINDATE.Text & txtFINTIME.Text & " dr=" & sReader.GetString(1) & sReader.GetString(2), My.User.Name)
                        'modDB.InsertSignRecord("AziTest", "FIN=" & txtFINDATE.Text & txtFINTIME.Text & " dr2=" & sReader.GetString(3) & sReader.GetString(4), My.User.Name)
                        If ((txtFINDATE.Text & txtFINTIME.Text) >= (sReader.GetString(1) & sReader.GetString(2))) And _
                           ((txtFINDATE.Text & txtFINTIME.Text) <= (sReader.GetString(3) & sReader.GetString(4))) Then
                            TIMECHK = True
                            'modDB.InsertSignRecord("AziTest", "OK-2", My.User.Name)
                        End If
                    Loop
                End If

                If TIMECHK Then
                    sReader.Close()
                    '-------------------------------------* 找出是否有重複資料
                    sSql = "SELECT EMPID FROM FINGER " _
                         & " WHERE EMPID='" & txtEmpID.Text & "'" _
                         & "   AND FINDATE='" & txtFINDATE.Text & "'" _
                         & "   AND FINTIME='" & txtFINTIME.Text & "00'"

                    sCmd.CommandText = sSql
                    sReader = sCmd.ExecuteReader()
                    sReader.Read()

                    If (sReader.HasRows) Then
                        sValidOK = False
                        sMsg = "已有資料重複!!"
                    End If
                Else
                    sValidOK = False
                    sMsg = "補卡的時間超出核可的範圍!!"
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
        Dim LCKDT As String = GET_LOCKYM("1", txtSTNID.Text) & "20"

        'modUtil.showMsg(Me.Page, "鎖檔日期", LCKDT)

        If pMode = FormViewMode.Edit Then
            If txtEmpName2.Text = "" Then
                modUtil.showMsg(Me.Page, "錯誤訊息(UpdateDB)", "需指定變更有效員工編號")
                e.Cancel = True
                Exit Sub
            End If
        End If
        '
        If Not CheckData(pMode) Then '* 取消新增/修改
            e.Cancel = True
            Exit Sub
        End If
        '20160215 新增增加備註
        If pMode = FormViewMode.Insert Then
            'modDB.InsertSignRecord("AziTest", "ATT_M_053U:Insert _DBupdate", My.User.Name)
            Dim ISql As String
            ISql = " INSERT INTO SCHPASSM (EMPID,SHSTDATE,SHSTTIME,PASSFG,MEMO,inuser,indate,intime)" _
                 + " VALUES ('" + Me.txtEmpID.Text + "','" + Me.txtFINDATE.Text + "','" + Me.txtFINTIME.Text + "','A'" _
                 + " ,'" + Me.TxtMEMO.Text + "','" + Me.Page.User.Identity.Name + "','" + DateTime.Now.ToString("yyyyMMdd") + "','" + DateTime.Now.ToString("HHmmss") + "')"
            '
            EXE_SQL(ISql)
        ElseIf pMode = FormViewMode.ReadOnly Then
            '刪除確認
            Dim DSql As String
            'modDB.InsertSignRecord("AziTest", "ATT_M_053U:Delete _DBupdate:" + Me.txtEmpID.Text + " " + Me.txtFINDATE.Text + " " + Me.txtFINTIME.Text, My.User.Name)
            DSql = " DELETE FROM SCHPASSM WHERE EMPID= '" + Me.txtEmpID.Text + "' AND SHSTDATE='" + Me.txtFINDATE.Text + "' AND SHSTTIME='" + Me.txtFINTIME.Text.Substring(0, 4) + "' AND PASSFG='A'"
            '
            EXE_SQL(DSql)
        End If
        '
        Try
            With e.Command
                .Parameters("@STNID").Value = Me.txtSTNID.Text
                .Parameters("@EMPID").Value = Me.txtEmpID.Text
                .Parameters("@FINDATE").Value = Me.txtFINDATE.Text
                .Parameters("@FINTIME").Value = Me.txtFINTIME.Text
                '-------------------------------------* 刪除舊資料
                If pMode = FormViewMode.Insert Then
                    .Parameters("@FINTIME").Value = Me.txtFINTIME.Text & "00"
                    .Parameters("@FINCLA").Value = "11"
                    .Parameters("@INUSER").Value = Me.Page.User.Identity.Name
                    .Parameters("@INDATE").Value = DateTime.Now.ToString("yyyyMMdd")
                    .Parameters("@INTIME").Value = DateTime.Now.ToString("HHmmss")
                ElseIf pMode = FormViewMode.Edit Then
                    .Parameters("@CHGEMPID").Value = txtEmpID2.Text
                    modDB.InsertSignRecord("ATTM053", "變更打卡員工代號: " + Me.txtEmpID.Text + "_" + Me.txtFINDATE.Text + "_" + Me.txtFINTIME.Text + " -> " + txtEmpID2.Text, My.User.Name)
                End If 'END OF INSERT
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

    Protected Sub txtEmpID2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call modCharge.GetEmpName(Me.txtEmpID2, Me.txtEmpName2, True)
    End Sub

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
End Class
