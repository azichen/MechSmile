<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ATT_Q_058.aspx.vb" Inherits="MMS_Report_ATT_Q_058" EnableEventValidation = "false" Theme="ThemeCHG" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title></title>
<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />  
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
            <Scripts>
                <asp:ScriptReference Path="~/Script/HighLight.js" />
            </Scripts>
            <Scripts>
                <asp:ScriptReference Path="~/Script/Util.js" />
            </Scripts>
        </asp:ScriptManager>
            
     <!-- 查詢輸入區畫面 -->    
     <uc:UpdatePanelFix ID="updQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
     <ContentTemplate>
     
          <table class="table">
            <tr>
                <td align="center" style="height: 15px" colspan="2">
                    <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="請假申請資料檢核列印"></asp:Label></td>
            </tr>
            <tr>
                <td style="height:5px;width:15%" > </td>
                <td style="height:5px;width:85%" > </td>
            </tr> 
            <tr>
               <td align="right" style="height: 25px;"><span class="MustInput" >*</span>責任區：</td>
               <td>
                 <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                 <ContentTemplate>
                   <asp:TextBox ID="txtArID1" runat="server" Width="50px" AutoPostBack="true" MaxLength="5" Enabled="False" /><asp:TextBox
                                ID="txtArName" runat="server" Width="120px" Enabled="False" />
                            <asp:Button ID="btnArea1" runat="server" Text="選擇區別" />
                 </ContentTemplate>
                 </asp:UpdatePanel>             
               </td> 
            </tr> 
            <tr>
                  <td align="right" style="height:25px;"><span class="MustInput" >*</span>請假日期：</td>
                  <td>
                     <asp:TextBox ID="txtDtFrom" runat="server" Width="88px" />
                     <img runat="server" id="imgDtFrom" src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false)" />
                     
                     &nbsp 至 &nbsp<asp:TextBox ID="txtDtTo" runat="server" Width="88px" />
                     <img runat="server" id="imgDtTo" src="../Images/date.gif" alt="按此可點選日期" style="border:0;"  onclick="ShowCalendar('txtDtTo',true,false)" />
                  </td>
            </tr>                                   
            <tr>
                <td style="height: 25px;width:15%">&nbsp</td>
                <td colspan="4">
                    <asp:Button ID="btnPreView" runat="server" Text="預覽列印" OnClick="btnPreView_Click" /><asp:Button 
                                ID="btnClear" runat="server" Text="清除重來" /> <asp:Button 
                                ID="btnSetVatm" runat="server" Text="列印壓註" OnClick="btnSetVatm_Click" /></td>
           </tr>
            <tr>
              <td colspan="2" style="height: 5px; background-color:Teal"  ></td>                      
            </tr>
            <tr>
              <td  align="center"  colspan="2" Font-Size="10pt" style="height: 25px" >   
              <asp:Label ID="MsgLabel" runat="server" CssClass="msg" Text="訊息行" ></asp:Label></td>                  
            </tr>             
          </table>
      </ContentTemplate>
      
      </uc:UpdatePanelFix>  
     
        <rsweb:ReportViewer ID="rptViewer" runat="server" Font-Names="Verdana" Font-Size="8pt" Visible="false"
               ProcessingMode="Remote" Width="1000px" Height="1400px"  >
        </rsweb:ReportViewer>
        
        
        <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" >  
        <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
        </div>
        <uc:ReportCredentials ID="ReportCredentials" runat="server" />
        <uc:QryArea ID="QryArea1" runat="server" AreIDObj="txtArID1" AreDescObj="txtArName" />
        <uc:UsrCalendar ID="UsrCalendar" runat="server" />
    </form>
</body>
</html>
