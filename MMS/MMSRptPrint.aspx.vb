Imports Microsoft.Reporting.WebForms

Partial Class Charge_RptPrint
    Inherits System.Web.UI.Page

    Protected Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Load
        Try
            '-------------------------------------* 報表路徑及認證
            Me.ReportViewer1.ProcessingMode = ProcessingMode.Remote
            Me.ReportViewer1.ServerReport.ReportServerCredentials = Me.ReportCredentials '* 身分認證 2008/04/03 NEC 杜
            Me.ReportViewer1.ServerReport.ReportServerUrl = New Uri(ConfigurationManager.AppSettings("ReportServerUrl"))
            Me.ReportViewer1.ServerReport.ReportPath = Request("ReportPath")

            '-------------------------------------* 報表參數
            'If Session("RptParam") IsNot Nothing Then
            '    Dim sHash As Hashtable = Session("RptParam")
            '    Dim I As Integer, sKeys(sHash.Count - 1), sValu(sHash.Count - 1) As String
            '    Dim sParas(sHash.Count - 1) As ReportParameter

            '    sHash.Keys.CopyTo(sKeys, 0)
            '    sHash.Values.CopyTo(sValu, 0)
            '    For I = 0 To sHash.Count - 1
            '        sParas(I) = New ReportParameter(sKeys(I), sValu(I))
            '    Next
            '    Me.ReportViewer1.ServerReport.SetParameters(sParas)
            '    sHash.Clear() : sHash = Nothing
            '    Session("RptParam") = Nothing
            'End If
            Me.ReportViewer1.DataBind()


            'ClientToolbarReportViewer1_ctl01.LoadPrintControl();return false; window.close();
        Catch ex As Exception
            'MsgBox(ex.Message)
            'modUtil.showMsg(Me.Page, "錯誤訊息", ex.Message)
        End Try
    End Sub

End Class
