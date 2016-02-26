Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms

Partial Class MMS_R_01
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
            Me.TextBox3.Attributes.Add("onkeypress", "return ToUpper()")
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

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreView.Click
        rptCmp.Visible = True
        Dim sParas(6) As ReportParameter '* 依參數個數產生陣列大小
        sParas(0) = New ReportParameter("AA1", Me.TextBox4.Text) '* 指定參數值
        sParas(1) = New ReportParameter("AA2", Me.TextBox5.Text)
        sParas(2) = New ReportParameter("BB1", Me.txtDtFrom.Text)
        sParas(3) = New ReportParameter("BB2", txtDtTo.Text)
        sParas(4) = New ReportParameter("CC1", Me.TextBox1.Text)
        sParas(5) = New ReportParameter("CC2", Me.TextBox2.Text)
        sParas(6) = New ReportParameter("DD", Me.TextBox3.Text)
        '-------------------------------------* 報表路徑及認證
        Me.rptCmp.Visible = True
        Try
            rptCmp.ProcessingMode = ProcessingMode.Remote
            rptCmp.ServerReport.ReportServerCredentials = Me.ReportCredentials '* 身分認證 2008/04/03 NEC 杜
            rptCmp.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))

            rptCmp.ServerReport.ReportPath = "/MMS/MMS_R01"
            rptCmp.ServerReport.SetParameters(sParas)
            rptCmp.ShowParameterPrompts = False
            rptCmp.DataBind()
        Catch ex As Exception

        End Try
        
    End Sub


    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub BtnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Server.Transfer("MMS_R_01.aspx", False)
    End Sub

    Protected Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.TextBox5.Text.Trim = "" Then
            Me.TextBox5.Text = Me.TextBox4.Text
        End If

    End Sub

    Protected Sub txtDtFrom_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.TextBox2.Text.Trim = "" Then
            Me.TextBox2.Text = Me.TextBox1.Text
        End If
    End Sub

    Protected Sub TextBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tmpsql As String
        tmpsql = "select * from MMSEmployee where EmployeeNo='" + Me.TextBox3.Text + "' and Salesman='Y'"
        Dim tt As System.Data.DataTable
        tt = get_DataTable(tmpsql)
        If tt.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "無此業務人員！")
            Me.Label1.Text = ""
            Exit Sub
        End If
        Me.Label1.Text = tt.Rows(0).Item("EmployeeName")
    End Sub
End Class
