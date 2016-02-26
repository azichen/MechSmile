Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports Microsoft.Reporting.WebForms
Imports modUtil

Partial Class MMS_R_04
    Inherits System.Web.UI.Page
    Dim STNID As String
    Dim STNNAME As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            modUtil.SetDateObj(Me.txtDatest, False, Nothing, False)
            Me.TextBox1.Attributes.Add("onkeypress", "return oTUpper()")
            Me.TextBox2.Attributes.Add("onkeypress", "return ToUpper()")
            'Me.txtDtFrom.Attributes.Add("onkeypress", "return ToUpper()")
            'Me.txtDtTo.Attributes.Add("onkeypress", "return ToUpper()")
            Dim sConn As SqlClient.SqlConnection = Nothing
            Dim sCmd As SqlClient.SqlCommand = Nothing
            Dim sSql As String

            sConn = New SqlClient.SqlConnection
            sConn.ConnectionString = modDB.GetDataSource.ConnectionString
            sConn.Open()
            sCmd = New SqlClient.SqlCommand()
            sCmd.Connection = sConn
            Dim sReader As SqlClient.SqlDataReader = Nothing
            STNID = Request.Cookies("STNID").Value
            'lblTitle.Text = Request.Cookies("STNID").Value 'Me.ViewState("RolDeptID") 'azitest
            '
            sSql = "SELECT AREACODE,AREANAME FROM MMSAREA"
            If (STNID = "460125") Or (STNID = "460126") Or _
               (STNID = "460127") Or (STNID = "460128") Then
                sSql = sSql & " WHERE UNITCODE IN (SELECT STNUID FROM MECHSTNM WHERE STNID='" & STNID & "')"
            End If
            sSql = sSql & " ORDER BY AREACODE"

            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()
            txtAreaItem.Items.Add("")
            While sReader.Read
                txtAreaItem.Items.Add(sReader.GetString(0) & " " & sReader.GetString(1))
            End While
            sReader.Close()
            If (STNID = "460125") Or (STNID = "460126") Or (STNID = "460127") Or (STNID = "460128") Then
                txtAreaItem.SelectedIndex = 1
                txtAreaItem.Enabled = False
            Else
                txtAreaItem.Text = ""
            End If
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
        Dim sValidOK As Boolean = True
        Dim sMsg As String = ""
        sValidOK = CheckNotEmpty(Me, "txtDateSt", sMsg, "截止日期")
        'sValidOK = sValidOK And CheckNotEmpty(Me, "txtAreaItem", sMsg, "服務區別")
        If sValidOK Then
            Dim Areacode As String = ""
            Areacode = modUnset.GetAreaCode(My.User.Name).Trim
            rptCmp.Visible = True

            Dim sParas(7) As ReportParameter '* 依參數個數產生陣列大小
            sParas(0) = New ReportParameter("AA", Me.txtDateSt.Text) '* 指定參數值
            If (Trim(Me.txtAreaItem.Text)).Length > 2 Then
                sParas(1) = New ReportParameter("BB1", Me.txtAreaItem.Text.Substring(0, 1))
                sParas(2) = New ReportParameter("BB2", Me.txtAreaItem.Text.Substring(1, Me.txtAreaItem.Text.Length - 1))
            Else
                sParas(1) = New ReportParameter("BB1", "")
                sParas(2) = New ReportParameter("BB2", "")
            End If
            sParas(3) = New ReportParameter("CC1", Me.TextBox1.Text)
            sParas(4) = New ReportParameter("CC2", Me.TextBox2.Text)
            'ADD IN 20141003
            'If (Areacode <> "") And (Areacode <> "F") Then
            '    sParas(5) = New ReportParameter("FF", Areacode)
            'Else
            '    sParas(5) = New ReportParameter("FF", "")
            'End If
            ''add in 20141021 增加截止日、回饋日、區名稱
            Dim RebackDate As String = DateTime.Parse(Me.txtDateSt.Text).AddDays(9).ToString("yyyy/MM/dd")
            'lblMsg.Text = txtDateSt.Text.Substring(0, 4) + "年" + txtDateSt.Text.Substring(5, 2) + "月" + txtDateSt.Text.Substring(8, 2) + "日" + " " + RebackDate.ToString
            '            + " " + RebackDate.Substring(0, 4) + "年" + RebackDate.Substring(5, 2) + "月" + RebackDate.Substring(8, 2) + "日"
            sParas(5) = New ReportParameter("DATE1", txtDateSt.Text.Substring(0, 4) + "年" + txtDateSt.Text.Substring(5, 2) + "月" + txtDateSt.Text.Substring(8, 2) + "日")
            sParas(6) = New ReportParameter("DATE2", RebackDate.Substring(0, 4) + "年" + RebackDate.Substring(5, 2) + "月" + RebackDate.Substring(8, 2) + "日")
            If RadioButtonList1.SelectedIndex = 1 Then
                sParas(7) = New ReportParameter("GG", "1")
            ElseIf RadioButtonList1.SelectedIndex = 2 Then
                sParas(7) = New ReportParameter("GG", "2")
            Else
                sParas(7) = New ReportParameter("GG", "")
            End If
            '-------------------------------------* 報表路徑及認證
            'lblTitle.Text = "OK! 截止日=" & sParas(5).Values.ToString & " 回饋日=" & sParas(6).Values.ToString 'azitest
            'showMsg(Me.Page, "OK! 截止日=" & sParas(5).Values.ToString & " 回饋日=" & sParas(6).Values.ToString, sMsg)
            rptCmp.ProcessingMode = ProcessingMode.Remote
            rptCmp.ServerReport.ReportServerCredentials = Me.ReportCredentials '* 身分認證 2008/04/03 NEC 杜
            rptCmp.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))

            rptCmp.ServerReport.ReportPath = "/MMS/MMS_R04"
            rptCmp.ServerReport.SetParameters(sParas)
            rptCmp.ShowParameterPrompts = False
            rptCmp.DataBind()
        Else
            showMsg(Me.Page, "訊息(請修正錯誤欄位)", sMsg)
        End If
    End Sub


    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub BtnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        'Server.Transfer("MMS_R_04.aspx", False)
        Response.Redirect("MMS_R_04.aspx")
    End Sub

    Protected Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.TextBox2.Text.Trim = "" Then
            Me.TextBox2.Text = Me.TextBox1.Text
        End If
    End Sub

End Class
