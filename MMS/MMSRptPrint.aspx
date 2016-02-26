<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MMSRptPrint.aspx.vb" Inherits="Charge_RptPrint" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    <script type="text/javascript">
    function PrintAndClose() {
      window.blur();
      ClientToolbarReportViewer1_ctl01.LoadPrintControl();
      //setTimeout("window.close()",10000);
      setInterval("window.close()",1000);
    }
    
    </script>
</head>
<body onload="PrintAndClose()"  >
    <form id="form1" runat="server"  >
    <div>
    
    </div>
    
    <uc:ReportCredentials ID="ReportCredentials" runat="server" />
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Height="300px" ProcessingMode="Remote"
            Width="500px">
        </rsweb:ReportViewer>
    </form>
</body>
</html>
