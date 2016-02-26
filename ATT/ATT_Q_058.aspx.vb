Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms

Partial Class MMS_Report_ATT_Q_058
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sRol As String = ""
        If Not IsPostBack Then
            'modUtil.SetDateObj(Me.TextBox4, False, Nothing, False)
            'Me.TextBox1.Attributes.Add("onkeypress", "return ToUpper()")
            'Me.TextBox2.Attributes.Add("onkeypress", "return ToUpper()")
            'Me.TextBox4.Attributes.Add("onkeypress", "return ToUpper()")
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
                'If ViewState("ROL").Substring(2, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                ViewState("ROL") = modUtil.GetRolData(Request)
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()


                'If modUtil.IsStnRol(Request) Then '* 加油站權限
                '    Me.txtStnID1.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                '    Me.txtStnName1.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                '    Me.txtStnID2.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                '    Me.txtStnName2.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                '    Me.btnStnID1.Visible = False
                '    Me.btnStnId2.Visible = False
                '    modUtil.SetObjReadOnly(form1, "txtStnID1")
                '    modUtil.SetObjReadOnly(form1, "txtStnID2")
                'End If
            End If
            STNID = Request.Cookies("STNID").Value
            STNNAME = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
            '
            Me.txtArID1.Attributes.Add("onkeypress", "return CheckKeyNumber()")

            modUtil.SetDateObj(Me.txtDtFrom, False, Me.txtDtTo, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)
            modUtil.SetDateImgObj(Me.imgDtFrom, Me.txtDtFrom, True, False, Me.txtDtTo)
            modUtil.SetDateImgObj(Me.imgDtTo, Me.txtDtTo, True, False)

            '檢查是否為總部、業務處、區長、站長，列出所屬單位加油站，
            Select Case Left(sRol, 1)
                Case "2"
                    '設定唯讀
                    txtArID1.Attributes("readonly") = "readonly"
                    txtArID1.CssClass = "readonly"
                    Me.txtArID1.Text = modDB.GetCodeName("ARE_ID", "SELECT ARE_ID FROM [MECHSTNM_View] where STNID='" & STNID & "'")
                    Me.txtArName.Text = modDB.GetCodeName("ARE_DESC", "SELECT ARE_DESC FROM [MECHSTNM_View] where STNID='" & STNID & "'")
                    Me.txtArID1.Enabled = False
                    Me.txtArName.Enabled = False
                    Me.btnArea1.Enabled = False
                Case "3"
                    txtArID1.Attributes("readonly") = "readonly"
                    txtArID1.CssClass = "readonly"
                    Me.txtArID1.Text = modDB.GetCodeName("ARE_ID", "SELECT ARE_ID FROM [MECHSTNM_View] where STNID='" & STNID & "'")
                    Me.txtArName.Text = modDB.GetCodeName("ARE_DESC", "SELECT ARE_DESC FROM [MECHSTNM_View] where STNID='" & STNID & "'")
                    Me.txtArID1.Enabled = False
                    Me.txtArName.Enabled = False
                    Me.btnArea1.Enabled = False
            End Select
            '
            Me.txtArID1.Focus()
            Me.MsgLabel.Text = Request("msg")
        End If

        'Me.item_defList.ConnectionString = modDB.GetDataSource.ConnectionString
        'Me.item_defList.SelectCommand = Me.txtSql.Text
        ''--------------------------------------------------* 設定GridView資料來源ID

        Me.ScriptManager1.RegisterPostBackControl(Me.btnPreView)
        Me.ScriptManager1.RegisterPostBackControl(Me.btnClear)
    End Sub

    '******************************************************************************************************
    '* 輸入之代碼達到一定之位數時，顯示代碼名稱
    '******************************************************************************************************
    Public Sub txtArID1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtArID1.TextChanged
        modOrder.GetAreaDesc(Me.txtArID1, Me.txtArName, Me.ViewState("RolDeptID"), True)
    End Sub

    '******************************************************************************************************
    '* 選擇區
    '******************************************************************************************************
    Protected Sub btnArea1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnArea1.Click
        Me.QryArea1.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
    End Sub

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Response.Redirect("ATT_Q_058.aspx")
    End Sub


    Protected Sub btnPreView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreView.Click
        '
        Dim VATMDTBASE As String = "2014/09/21"
        Dim tmpstr As String = ""
        Dim sParas(3) As ReportParameter '* 依參數個數產生陣列大小
        Dim sMsg As String = ""
        Dim sValidOK As Boolean = True
        '
        MsgLabel.Text = ""
        btnPreView.Enabled = False
        btnClear.Enabled = False
        btnSetVatm.Enabled = False

        'sValidOK = sValidOK And modUtil.CheckNotEmpty(Me, "txtArID1", sMsg, "區代號")
        sValidOK = sValidOK And modUtil.CheckNotEmpty(Me, "txtDtFrom", sMsg, "起始日期")
        sValidOK = sValidOK And modUtil.CheckNotEmpty(Me, "txtDtTo", sMsg, "迄止日期")
        If sValidOK And (Trim(txtArName.Text) = "") Then
            sValidOK = False
            sMsg = "區別錯誤"
        End If

        If Not sValidOK Then
            Me.rptViewer.Visible = False
            modUtil.showMsg(Me.Page, "訊息", sMsg)
            Exit Sub
        End If
        '
        Dim sConn As SqlClient.SqlConnection = Nothing
        Dim sCmd As SqlClient.SqlCommand = Nothing

        sConn = New SqlClient.SqlConnection
        sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
        sConn.Open()

        Dim sReader As SqlClient.SqlDataReader = Nothing
        sCmd = New SqlClient.SqlCommand()
        sCmd.Connection = sConn
        'modDB.InsertSignRecord("AziTest", "ok-1", My.User.Name) 'AZITEST
        '
        tmpstr = "SELECT D.ARE_ID,D.STNID,VATM_EMPLID,VATM_SEQ" _
               + "  FROM MP_HR.DBO.VACATM A " _
               + " INNER JOIN MP_HR.DBO.VACAMF B ON A.VATM_VANMID=B.VACA_ID " _
               + " INNER JOIN MP_HR.DBO.EMPLOYEE C ON A.VATM_EMPLID=C.EMPL_ID " _
               + " INNER JOIN [10.1.1.100].SMILE_HQ.DBO.MECHSTNM D ON C.EMPL_DEPTID=D.STNID " _
               + " WHERE D.ARE_ID='" + Me.txtArID1.Text + "' AND VATM_DATE_ST>='" + VATMDTBASE + "' " _
               + "   AND VATM_DATE_ST>='" + Me.txtDtFrom.Text + "' AND VATM_DATE_ST<='" + Me.txtDtTo.Text + "'" _
               + "   AND VATM_VANMID<>'010'" _
               + "   AND VATM_EMPLID+CONVERT(varchar,VATM_SEQ)+CONVERT(char(20), VATM_U_DATE,120) NOT IN " _
               + "(SELECT EMPID+CONVERT(varchar,VATMSEQ)+CONVERT(char(20),VATMDATEU,120) FROM VATMPRNLOG" _
               + "  WHERE EFFECTIVE='Y')"
        '+ "  WHERE AREAID='" + txtArID1.Text + "' AND  EFFECTIVE='Y')"
        '
        sCmd.CommandText = tmpstr
        sReader = sCmd.ExecuteReader()
        sReader.Read()
        If (sReader.HasRows) Then
            'modDB.InsertSignRecord("AziTest", "ok-2", My.User.Name) 'AZITEST
            'sConn.Close()
            'sCmd.Dispose() '手動釋放資源
            'sConn.Dispose()
            'sConn = Nothing '移除指標
            '
            'MsgLabel.Text = "列印OK" 'azitemp
            'modUtil.showMsg(Me.Page, "訊息", "列印OK")
            sReader.Close()
            '作註記處理
            tmpstr = "INSERT INTO VATMPRNLOG " _
                           + " SELECT D.ARE_ID,D.STNID,VATM_EMPLID,VATM_SEQ,VATM_U_DATE,GETDATE() AS PRNDATE" _
                           + "  ,'" + Me.Page.User.Identity.Name + "' AS PRNEMPID,'V' AS EFFITIVE FROM MP_HR.DBO.VACATM A " _
                           + " INNER JOIN MP_HR.DBO.VACAMF B ON A.VATM_VANMID=B.VACA_ID " _
                           + " INNER JOIN MP_HR.DBO.EMPLOYEE C ON A.VATM_EMPLID=C.EMPL_ID " _
                           + " INNER JOIN [10.1.1.100].SMILE_HQ.DBO.MECHSTNM D ON C.EMPL_DEPTID=D.STNID " _
                           + " WHERE D.ARE_ID='" + txtArID1.Text + "' AND VATM_DATE_ST>='" + VATMDTBASE + "' " _
                           + "   AND VATM_VANMID<>'010'" _
                           + "   AND VATM_DATE_ST>='" + Me.txtDtFrom.Text + "' AND VATM_DATE_ST<='" + Me.txtDtTo.Text + "'" _
                           + "   AND VATM_EMPLID+CONVERT(varchar,VATM_SEQ)+CONVERT(char(20), VATM_U_DATE,120) NOT IN " _
                           + "(SELECT EMPID+CONVERT(varchar,VATMSEQ)+CONVERT(char(20),VATMDATEU,120) FROM VATMPRNLOG)"
            'sConn.Open()
            'modDB.InsertSignRecord("AziTest", "ok-3", My.User.Name) 'AZITEST
            sCmd.CommandText = tmpstr
            sCmd.ExecuteNonQuery()
            'modDB.InsertSignRecord("AziTest", "ok-4", My.User.Name) 'AZITEST
            '
            Me.rptViewer.Visible = True
            Try
                'modDB.InsertSignRecord("AziTest", "AREID=" + txtArID1.Text, My.User.Name) 'AZITEST
                '-------------------------------------* 指定報表參數值
                sParas(0) = New ReportParameter("AREID", Me.txtArID1.Text) '* 指定參數值  Me.ddlArea.Text
                'sParas(1) = New ReportParameter("DATEST", VATMDTBASE)
                sParas(1) = New ReportParameter("VARENAME", Me.txtArName.Text)
                sParas(2) = New ReportParameter("DATE1", Me.txtDtFrom.Text)
                sParas(3) = New ReportParameter("DATE2", Me.txtDtTo.Text)
                'modDB.InsertSignRecord("AziTest", "ok-5", My.User.Name) 'AZITEST

                '-------------------------------------* 報表路徑及認證
                Me.rptViewer.ProcessingMode = ProcessingMode.Remote
                Me.rptViewer.ServerReport.ReportServerCredentials = Me.ReportCredentials
                Me.rptViewer.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))
                Me.rptViewer.ServerReport.ReportPath = "/OIL/rptSvrATT058"
                Me.rptViewer.ServerReport.SetParameters(sParas)
                Me.rptViewer.DataBind()
                'modDB.InsertSignRecord("AziTest", "ok-6", My.User.Name) 'AZITEST
                '
            Catch ex As Exception
                modUtil.showMsg(Me.Page, "錯誤訊息", ex.Message)
            End Try
        Else
            MsgLabel.Text = "查無資料"
            modUtil.showMsg(Me.Page, "錯誤訊息", "查無資料")
        End If
        'modDB.InsertSignRecord("AziTest", "ok-7", My.User.Name) 'AZITEST
        '
        sConn.Close()
        sCmd.Dispose() '手動釋放資源
        sConn.Dispose()
        sConn = Nothing '移除指標
        '
        btnPreView.Enabled = True
        btnClear.Enabled = True
        btnSetVatm.Enabled = True
        'End If
    End Sub

    '執行quary不回傳值
    'Public Function EXE_SQL(ByVal TMPSQL1 As String) As Boolean
    '    Dim AA As Integer
    '    Dim CONNSTR1 As String = modDB.GetDataSource.ConnectionString
    '    Dim CONN1 As New System.Data.SqlClient.SqlConnection(CONNSTR1)
    '    Dim SQLCOMMAND1 As New System.Data.SqlClient.SqlCommand
    '    Try
    '        SQLCOMMAND1.CommandText = TMPSQL1
    '        If CONN1.State.Closed = ConnectionState.Open Then CONN1.Close()
    '        SQLCOMMAND1.Connection = CONN1
    '        CONN1.Open()
    '        SQLCOMMAND1.ExecuteNonQuery()
    '        CONN1.Close()
    '    Catch ex As Exception
    '        EXE_SQL = False
    '        Exit Function
    '    End Try
    '    EXE_SQL = True
    'End Function

    Protected Sub btnSetVatm_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sConn As SqlClient.SqlConnection = Nothing
        Dim sCmd As SqlClient.SqlCommand = Nothing
        Dim VATMDTBASE As String = "2014/09/21"

        sConn = New SqlClient.SqlConnection
        sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
        sConn.Open()
        '
        sCmd = New SqlClient.SqlCommand()
        sCmd.Connection = sConn
        '
        Dim tmpstr As String
        Try
            modDB.InsertSignRecord("(ATTQ058)", "列印註記_區域:" + txtArID1.Text, My.User.Name) 'AZITEST
            '
            tmpstr = "UPDATE VATMPRNLOG SET EFFECTIVE='Y' " _
                   + " WHERE AREAID='" + txtArID1.Text + "' AND EFFECTIVE='V' " _
            + "   AND EMPID+CONVERT(VARCHAR,VATMSEQ) IN (SELECT VATM_EMPLID+CONVERT(VARCHAR,VATM_SEQ) FROM MP_HR.DBO.VACATM A" _
            + " INNER JOIN MP_HR.DBO.EMPLOYEE C ON A.VATM_EMPLID=C.EMPL_ID " _
            + " INNER JOIN [10.1.1.100].SMILE_HQ.DBO.MECHSTNM D ON C.EMPL_DEPTID=D.STNID " _
            + " WHERE D.ARE_ID='" + txtArID1.Text + "' AND VATM_DATE_ST>='" + VATMDTBASE + "' " _
            + "   AND VATM_DATE_ST>='" + Me.txtDtFrom.Text + "' AND VATM_DATE_ST<='" + Me.txtDtTo.Text + "'"

            sCmd.CommandText = tmpstr
            sCmd.ExecuteNonQuery()
            '
            modUtil.showMsg(Me.Page, "完成", "已註記完成")
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息", ex.Message)
        End Try
        sConn.Close()
        sCmd.Dispose() '手動釋放資源
        sConn.Dispose()
        sConn = Nothing '移除指標
    End Sub

End Class
