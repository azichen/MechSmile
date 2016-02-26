Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms

Partial Class MMS_R_03
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Me.TextBox1.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox2.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox3.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox4.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox5.Attributes.Add("onkeypress", "return ToUpper()")
            Me.TextBox6.Attributes.Add("onkeypress", "return ToUpper()")

            modUtil.SetDateObj(Me.TextBox4, False, Nothing, True)
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


    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub BtnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        'Server.Transfer("MMS_R_03.aspx", False)
        Response.Redirect("MMS_R_03.aspx")
    End Sub

    Protected Sub TextBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        Dim tmpsql As String
        tmpsql = "select * from MMSEmployee where EmployeeNo='" + Me.TextBox3.Text + "' and accountant='Y'"
        Dim tt As System.Data.DataTable
        tt = get_DataTable(tmpsql)
        If tt.Rows.Count = 0 Then
            modUtil.showMsg(Me.Page, "訊息", "無此會計人員！")
            Me.Label1.Text = ""
            Exit Sub
        End If
        Me.Label1.Text = tt.Rows(0).Item("EmployeeName")
    End Sub

    Protected Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If Me.TextBox2.Text.Trim = "" Then
            Me.TextBox2.Text = Me.TextBox1.Text
        End If
    End Sub

    Protected Sub btnPreView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPreView.Click
        Dim Areacode As String = ""
        Areacode = modUnset.GetAreaCode(My.User.Name).Trim
        '
        Dim sParas(6) As ReportParameter '* 依參數個數產生陣列大小
        rptCmp.Visible = True
        sParas(0) = New ReportParameter("AA", Me.TextBox4.Text) '* 指定參數值
        sParas(1) = New ReportParameter("BB1", Me.TextBox1.Text)
        sParas(2) = New ReportParameter("BB2", TextBox2.Text)
        sParas(3) = New ReportParameter("CC1", Me.TextBox5.Text)
        sParas(4) = New ReportParameter("CC2", Me.TextBox6.Text)
        sParas(5) = New ReportParameter("DD", Me.TextBox3.Text)
        'ADD IN 20140912
        If (Areacode <> "") And (Areacode <> "F") Then
            sParas(6) = New ReportParameter("FF", Areacode)
        Else
            sParas(6) = New ReportParameter("FF", "")
        End If
        '-------------------------------------* 報表路徑及認證
        rptCmp.ProcessingMode = ProcessingMode.Remote
        rptCmp.ServerReport.ReportServerCredentials = Me.ReportCredentials '* 身分認證 2008/04/03 NEC 杜
        rptCmp.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))
        If RadioButtonList1.SelectedIndex = 0 Then
            rptCmp.ServerReport.ReportPath = "/MMS/MMS_R03_1"
        Else
            rptCmp.ServerReport.ReportPath = "/MMS/MMS_R03_2"
        End If
        rptCmp.ServerReport.SetParameters(sParas)
        rptCmp.ShowParameterPrompts = False
        rptCmp.DataBind()

    End Sub

    Protected Sub txtDtFrom_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub TextBox5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.TextBox6.Text.Trim = "" Then
            Me.TextBox6.Text = Me.TextBox5.Text
        End If
    End Sub

    Protected Sub TextBox6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub
End Class
