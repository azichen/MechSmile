﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="~/MMS/MMS_R_01.aspx.vb"  Inherits="MMS_R_01" EnableEventValidation = "false" Theme="ThemeCHG" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>簽帳明細表</title>
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
                    <td align="center" colspan="6" style="height: 21px">
                        <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="客戶發票簽收單"></asp:Label></td>
                </tr>
                <tr>
                    <td align="right" style="width: 50px; height: 25px">
                    </td>
                    <td align="right" style="width: 89px; height: 25px">
                        發票日期：</td>
                    <td style="width: 121px; height: 25px">
                        <asp:TextBox ID="TextBox4" runat="server" AutoPostBack="True" OnTextChanged="TextBox4_TextChanged"
                            Width="88px"></asp:TextBox><img alt="按此可點選日期" onclick="ShowCalendar('TextBox4',true,false,'TextBox5')"
                                src="../Images/date.gif" style="border-right: 0px; border-top: 0px; border-left: 0px;
                                border-bottom: 0px" /></td>
                    <td align="center" style="width: 24px; height: 25px">
                        至</td>
                    <td colspan="2">
                        <asp:TextBox ID="TextBox5" runat="server" Width="88px"></asp:TextBox><img alt="按此可點選日期"
                            onclick="ShowCalendar('TextBox5',true,false)" src="../Images/date.gif" style="border-right: 0px;
                            border-top: 0px; border-left: 0px; border-bottom: 0px" /></td>
                </tr>
                <tr>
                    <td align="right" style="width: 50px; height: 25px">
                    </td>
                    <td align="right" style="width: 89px; height: 25px">
                        申請日期：</td>
                    <td style="width: 121px; height: 25px">
                        <asp:TextBox ID="txtDtFrom" runat="server" OnTextChanged="txtDtFrom_TextChanged"
                            Width="88px"></asp:TextBox><img alt="按此可點選日期" onclick="ShowCalendar('txtDtFrom',true,false,'txtDtTo')"
                                src="../Images/date.gif" style="border-right: 0px; border-top: 0px; border-left: 0px;
                                border-bottom: 0px" /></td>
                    <td align="center" style="width: 24px; height: 25px">
                        至</td>
                    <td colspan="2">
                        <asp:TextBox ID="txtDtTo" runat="server" Width="88px"></asp:TextBox><img alt="按此可點選日期"
                            onclick="ShowCalendar('txtDtTo',true,false)" src="../Images/date.gif" style="border-right: 0px;
                            border-top: 0px; border-left: 0px; border-bottom: 0px" />
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="right" style="width: 50px; height: 29px">
                    </td>
                    <td align="right" style="width: 89px; height: 29px">
                        客戶代號：</td>
                    <td style="width: 121px; height: 29px">
                        <asp:TextBox ID="TextBox1" runat="server" AutoPostBack="True" OnTextChanged="TextBox1_TextChanged"
                            Width="107px"></asp:TextBox></td>
                    <td align="center" style="width: 24px; height: 29px">
                        至</td>
                    <td colspan="5" style="height: 29px">
                        <asp:TextBox ID="TextBox2" runat="server" Width="107px"></asp:TextBox>&nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="right" style="width: 50px; height: 29px">
                    </td>
                    <td align="right" style="width: 89px; height: 29px">
                        業務員：</td>
                    <td style="width: 121px; height: 29px">
                        <asp:TextBox ID="TextBox3" runat="server" AutoPostBack="True" OnTextChanged="TextBox3_TextChanged"
                            Width="107px"></asp:TextBox></td>
                    <td colspan="6" style="height: 29px">
                        &nbsp;<asp:Label ID="Label1" runat="server"></asp:Label></td>
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