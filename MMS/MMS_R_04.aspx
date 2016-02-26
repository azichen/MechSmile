<%@ Page Language="VB" AutoEventWireup="false" CodeFile="~/MMS/MMS_R_04.aspx.vb"  Debug="true" Inherits="MMS_R_04" EnableEventValidation = "false" Theme="ThemeCHG" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>逾期應收款催收表</title>
<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />  
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form id="form1" runat="server">
    <div>
        <asp:Panel ID="Panel1" runat="server" Height="80px" Width="720px">
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="true" EnableScriptLocalization="true">
          <Scripts> <asp:ScriptReference Path="~/Script/Util.js" /> </Scripts>  
        </asp:ScriptManager>
      <uc:ReportCredentials ID="ReportCredentials" runat="server" />
      <uc:UpdatePanelFix ID="updQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
        <ContentTemplate>
            <table class="table">
                <tr>
                    <td align="center" colspan="3" style="height: 21px">
                        <asp:Label ID="lblTitle" runat="server" CssClass="titles" ForeColor="#000000" Text="逾期應收款催收表"></asp:Label></td>
                </tr>
                <tr>
                    <td align="right" style="width: 50px; height: 25px">
                    </td>
                    <td align="right" style="width: 97px; height: 25px">
                        <span style="color: #ff3300"><span style="color: #333333">*</span><span style="color: #333333">資料截止日</span></span>：</td>
                    <td style="width: 121px; height: 25px">
                        <asp:TextBox ID="txtDateSt" runat="server" AutoPostBack="True" 
                            Width="88px"></asp:TextBox> <img alt="按此可點選日期" onclick="ShowCalendar('txtDateSt',true,false)" src="../Images/date.gif" /> </td> 
                </tr>
                <tr>
                    <td align="right" style="width: 50px; height: 25px">
                    </td>
                    <td align="right" style="width: 97px; height: 25px"><span style="color: #ff3300">
                        <span style="color: #333333">區域：</span></td>
                    <td><asp:DROPDOWNLIST ID="txtAreaItem" runat="server"  Width="200px" AutoPostBack="True" /> </td>
                    
                </tr>
                <tr>
                    <td align="right" style="width: 50px; height: 29px">
                    </td>
                    <td align="right" style="width: 97px; height: 29px">
                        業務員：</td>
                    <td style="width: 300px; height: 29px">
                        <asp:TextBox ID="TextBox1" runat="server" AutoPostBack="True" 
                            Width="107px" OnTextChanged="TextBox1_TextChanged"></asp:TextBox> 至  <asp:TextBox ID="TextBox2" runat="server" Width="107px"></asp:TextBox></td>
                </tr>
                 <tr>
                    <td align="right" style="width: 50px; height: 29px"></td>
                    <td align="right" style="width: 97px; height: 29px"></td>
                    <td>
                        <asp:RadioButtonList ID="RadioButtonList1" runat="server" 
                            RepeatDirection="Horizontal">
                            <asp:ListItem Selected="True">全部</asp:ListItem>
                            <asp:ListItem>未逾期</asp:ListItem>
                            <asp:ListItem>已逾期</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
            </table>
          <table class="table">
            <tr>
                <td style="height: 25px; width: 194px;"></td><td>
                    <asp:Button ID="btnPreView" runat="server" Text="預覽列印" /><asp:Button 
                                ID="btnClear" runat="server" Text="清除重來" /></td>
           </tr>
              <tr>
             <td colspan="2" align="left">
                 <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label></td>
              </tr>
            <tr>
              <td colspan="2" style="height: 5px; background-color:Teal"  ></td>                      
            </tr>
          </table>
        </ContentTemplate>
      </uc:UpdatePanelFix>
        <rsweb:ReportViewer ID="rptCmp" runat="server" Font-Names="Verdana" Font-Size="8pt" Visible=false
            Height="1200px" ProcessingMode="Remote" Width="1100px">
        </rsweb:ReportViewer>
        </asp:Panel>
    </div>
      <uc:UsrCalendar ID="UsrCalendar" runat="server" />  
    </form>
</body>
</html>
