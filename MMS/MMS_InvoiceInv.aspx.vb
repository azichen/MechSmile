Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms

Partial Class MMS_InvoiceInv
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

            modUtil.SetDateObj(Me.txtDtFrom, False, Nothing, False)
            modUtil.SetDateObj(Me.txtDtTo, False, Nothing, False)

            modUtil.SetDateObj(Me.TextBox4, False, Nothing, False)
            modUtil.SetDateObj(Me.TextBox5, False, Nothing, False)

            Me.TextBox1.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox2.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox4.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox5.Attributes.Add("onkeypress", "return ToUpper()")
        End If

        Me.ScriptManager1.RegisterPostBackControl(Me.btnPreView)
        Me.ScriptManager1.RegisterPostBackControl(Me.btnClear)
    End Sub

    '連接資料庫取得datatable
    Public Function get_DataTable(ByVal SQL1 As String) As System.Data.DataTable
        Dim CONNSTR1 As String = modDB.GetDataSource.ConnectionString
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

    Protected Sub btnPreView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreView.Click
        Dim sValidOK As Boolean = True
        rptCmp.Visible = True
        Dim instring1 As String = ""
        'If Me.TextBox4.Text.Trim + Me.TextBox5.Text.Trim + Me.txtDtFrom.Text.Trim + Me.txtDtTo.Text.Trim + Me.TextBox1.Text.Trim + Me.TextBox2.Text.Trim = "" Then
        '    modUtil.showMsg(Me.Page, "訊息", "尚未輸入任何條件！")
        '    Exit Sub
        'End If

        'modDB.InsertSignRecord("AziTest", "ok-1", My.User.Name) 'AZITEST

        If ((Me.TextBox1.Text = "") Or (Me.TextBox2.Text.Trim() = "")) Then
            modUtil.showMsg(Me.Page, "訊息", "需輸入發票號碼！")
            Exit Sub
        End If

        Dim VAreaCode As String = GetAreaCode()
        Dim tmpsql As String

        '機械部可以因其他部門發票作廢申請單
        If VAreaCode = "F" Then
            VAreaCode = ""
        End If

        tmpsql = "Select A.InvoiceNo from MMSInvoiceInvalid A "
        tmpsql += " inner join MMSInvoicem B on A.Invoiceno=B.Invoiceno "
        tmpsql += " inner join MMSCustomers C on B.CustomerNo =C.CustomerNo "
        tmpsql += " where A.InvoiceNo<>'xxxx' "

        'modDB.InsertSignRecord("AziTest", "VAreaCode=" + VAreaCode, My.User.Name) 'AZITEST

        If (VAreaCode.Trim <> "") Then
            tmpsql += " and C.areacode ='" + VAreaCode + "' "
        End If
        '
        If Me.TextBox4.Text.Trim <> "" Then
            tmpsql += " and B.InvoiceDate>='" + Me.TextBox4.Text.Trim + "'"
        End If
        If Me.TextBox5.Text.Trim <> "" Then
            tmpsql += " and B.InvoiceDate<='" + Me.TextBox5.Text.Trim + "'"
        End If
        If Me.txtDtFrom.Text.Trim <> "" Then
            tmpsql += " and ApplyDate>='" + Me.txtDtFrom.Text.Trim + "'"
        End If
        If Me.txtDtTo.Text.Trim <> "" Then
            tmpsql += " and ApplyDate<='" + Me.txtDtTo.Text.Trim + "'"
        End If
        If Me.TextBox1.Text.Trim <> "" Then
            tmpsql += " and A.InvoiceNo>='" + Me.TextBox1.Text.Trim + "'"
        End If

        If Me.TextBox2.Text.Trim <> "" Then
            tmpsql += " and A.InvoiceNo<='" + Me.TextBox2.Text.Trim + "'"
        End If
        'modDB.InsertSignRecord("AziTest", "OK-A", My.User.Name) 'AZITEST
        '
        'LOGSQL(tmpsql)
        Dim tt As DataTable
        tt = get_DataTable(tmpsql)
        If tt.Rows.Count <= 0 Then
            lblMsg.Text = "查無資料!"
            'modDB.InsertSignRecord("AziTest", "無資料", My.User.Name) 'AZITEST
            'modUtil.showMsg(Me.Page, "訊息", "查無資料！")
            'Exit Sub
        Else
            'modDB.InsertSignRecord("AziTest", "SHOW資料", My.User.Name) 'AZITEST
            Dim sParas(6) As ReportParameter '* 依參數個數產生陣列大小
            sParas(0) = New ReportParameter("AA1", Me.TextBox4.Text) '* 指定參數值
            sParas(1) = New ReportParameter("AA2", Me.TextBox5.Text)
            sParas(2) = New ReportParameter("BB1", Me.txtDtFrom.Text)
            sParas(3) = New ReportParameter("BB2", Me.txtDtTo.Text)
            sParas(4) = New ReportParameter("CC1", Me.TextBox1.Text)
            sParas(5) = New ReportParameter("CC2", Me.TextBox2.Text)
            sParas(6) = New ReportParameter("DD", VAreaCode.Trim)
            '-------------------------------------* 報表路徑及認證
            rptCmp.ProcessingMode = ProcessingMode.Remote
            rptCmp.ServerReport.ReportServerCredentials = Me.ReportCredentials '* 身分認證
            rptCmp.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))

            rptCmp.ServerReport.ReportPath = "/MMS/MMS_InvoiceInvalid"
            rptCmp.ServerReport.SetParameters(sParas)
            rptCmp.ShowParameterPrompts = False
            rptCmp.DataBind()
        End If
        'Dim i As Integer
        'For i = 0 To tt.Rows.Count - 1
        '    If instring1.Length > 0 Then
        '        instring1 += ","
        '    End If
        '    instring1 += "'" + tt.Rows(i).Item("InvoiceNo") + "'"
        'Next
        '

    End Sub


    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub BtnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        'Server.Transfer("MMS_R_01.aspx", False)
        Response.Redirect("MMS_InvoiceInv.aspx")
    End Sub

    'Protected Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If Me.TextBox5.Text.Trim = "" Then
    '        Me.TextBox5.Text = Me.TextBox4.Text
    '    End If

    'End Sub

    'Protected Sub txtDtFrom_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    'End Sub

    'Protected Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If Me.TextBox2.Text.Trim = "" Then
    '        Me.TextBox2.Text = Me.TextBox1.Text
    '    End If
    'End Sub

    'Protected Sub TextBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim tmpsql As String
    '    tmpsql = "select * from MMSEmployee where EmployeeNo='" + Me.TextBox3.Text + "' and Salesman='Y'"
    '    Dim tt As System.Data.DataTable
    '    tt = get_DataTable(tmpsql)
    '    If tt.Rows.Count = 0 Then
    '        modUtil.showMsg(Me.Page, "訊息", "無此業務人員！")
    '        Me.Label1.Text = ""
    '        Exit Sub
    '    End If
    '    Me.Label1.Text = tt.Rows(0).Item("EmployeeName")
    'End Sub

    Private Function GetAreaCode() As String
        Dim tmpsql As String
        'Dim STNID As String = Me.ViewState("RolDeptID")
        'modDB.InsertSignRecord("AziTest", "STNID=" + STNID, My.User.Name) 'AZITEST
        tmpsql = "Select AreaCode From MMSArea A " _
               + " INNER JOIN MECHSTNM B ON A.UNITCODE=B.STNUID" _
               + " INNER JOIN USER_PRO C ON B.STNID=C.STNID" _
               + " WHERE Account='" + My.User.Name + "'"
        Try
            GetAreaCode = get_DataTable(tmpsql).Rows(0).Item("AreaCode")
        Catch ex As Exception
            GetAreaCode = ""
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

End Class
