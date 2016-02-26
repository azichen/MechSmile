<%@ Page Language="VB" DEBUG="false" AutoEventWireup="false" CodeFile="ATT_Q_059.aspx.vb" Inherits="ATT_Q_059" Theme="ThemeCHG" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
   <title>雙週出勤周期設定查詢</title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
   <form id="form1" runat="server">
      <!-- 宣告 ScriptManager /SqlDataSource / 使用者控制項 -->
      <asp:ScriptManager ID="ScriptManager1" runat="server">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
      
      <asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:MechPAConnectionString %>" />
      
      <div style="width:10px;height:10px;position:absolute;left:-200px;top:-10px;" >
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
      </div>
      <uc:UsrCalendar ID="UsrCalendar" runat="server" />

      <!-- 查詢輸入區畫面 -->
      <uc:UpdatePanelFix ID="updQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr align="center">
                  <td colspan="2">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="雙週出勤周期設定查詢" ></asp:Label></td>
               </tr>
               <tr style="height:5px">
                  <td style="width:15%;"></td>
                  <td style="width:85%;"></td>
               </tr>
               
               <tr>
                  <td align="right" style="height:25px;"></span>設定：</td>
                  <td> 西元年
                     <asp:DropDownList ID="ddlYear" runat="server" AppendDataBoundItems="true" AutoPostBack="true" OnTextChanged="ddlYear_TextChanged">
            　　　        <asp:ListItem Text="" />
              　　   </asp:DropDownList>
              　　   <asp:Button ID="Button1" runat="server" Text="自動設定" OnClick="Button1_Click" Visible="False"/>
                  </td>
               </tr>
               
                <tr>
                  <td align="right" style="height:25px;"><span class="MustInput" >*</span>日期：</td>
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
                     <asp:Button ID="btnQry" runat="server" Text="查詢"
                   /><asp:Button ID="btnClear" runat="server" Text="清除重來"
                   /><asp:Button ID="btnExcel" runat="server" Text="EXCEL匯出" Enabled="false" />
                     <img alt="" style="border:0" src="../Images/teacher_icon.gif" />
                     </td>
               </tr>
               <tr>
                  <td colspan="2" align="left">
                  <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label></td>
               </tr>
            </table>
            <asp:GridView ID="grdList" runat="server" />
         </ContentTemplate>
         <Triggers>
            <asp:AsyncPostBackTrigger ControlID="btnQry" EventName="Click" />
         </Triggers>
      </uc:UpdatePanelFix>        
   </form>
</body>
</html>

