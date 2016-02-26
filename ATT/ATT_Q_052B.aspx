<%@ Page Language="VB" DEBUG="false" AutoEventWireup="false" CodeFile="ATT_Q_052B.aspx.vb" Inherits="ATT_Q_052B" Theme="ThemeCHG" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
   <title>公出假資料查詢</title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
</head>

<body onkeydown="if(event.keyCode==13)event.keyCode=9">
   <form id="form1" runat="server">
      <!-- 宣告 ScriptManager /SqlDataSource / 使用者控制項 -->
      <asp:ScriptManager ID="ScriptManager1" runat="server">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
      
      <asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>" />
      
      <uc:UsrCalendar ID="UsrCalendar" runat="server" />

      <!-- 查詢輸入區畫面 -->
      <uc:UpdatePanelFix ID="updQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr align="center">
                  <td colspan="2">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="請假資料" ></asp:Label></td>
               </tr>
               <tr style="height:5px">
                  <td style="width:15%;"></td>
                  <td style="width:85%;"></td>
               </tr>
               <tr>
                  <td style="height: 25px; text-align:right;">員工職號：</td>
                  <td style="text-align:left;">
                      <asp:UpdatePanel ID="udpEmp" runat="server" RenderMode="Inline" UpdateMode="Conditional">
                      <ContentTemplate>
                      <asp:TextBox ID="txtEmpID" runat="server" Width="80px" Enabled="False" />
                      <asp:TextBox ID="txtEmpName" runat="server" Width="100px" Enabled="False"></asp:TextBox>
                      </ContentTemplate>
                      </asp:UpdatePanel></td>
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
            </table>
         </ContentTemplate>
      </uc:UpdatePanelFix>
      
      <!-- 查詢結果區畫面 -->
      <uc:UpdatePanelFix ID="updGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td style="height:25px;width:15%">&nbsp</td>
                  <td style="width:85%">
                  <asp:Button ID="btnBack" runat="server" Text="回出勤統計表"  OnClick="btnFunc_Click" />
                     <img alt="" style="border:0" src="../Images/teacher_icon.gif"
                   /><asp:LinkButton ID="btnNew" runat="server" Text="新增一筆資料" style="font-size:medium" />
                  </td>
               </tr>
               <tr>
                  <td colspan="2" align="left">
                  <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label></td>
               </tr>
            </table>
            <asp:GridView ID="grdList" runat="server" />
         </ContentTemplate>
      </uc:UpdatePanelFix>
   </form>
</body>
</html>

