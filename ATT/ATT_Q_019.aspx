<%@ Page Language="C#" debug="false" AutoEventWireup="true" CodeFile="~/ATT/ATT_Q_019.aspx.cs" Inherits="ATT_Q_019" EnableEventValidation = "false" Theme="ThemeCHG" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
   <title>區長出勤統計查詢</title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />  
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
   <form id="form1" runat="server">     
      <!-- 宣告 ScriptManager /SqlDataSource / 使用者控制項 -->
      <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="true" EnableScriptLocalization="true">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
      
      <asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>" />
      
      <div style="width:10px; height:10px; position:absolute; left:-200px; top:-10px;" >
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
         <uc:QryEmp ID="QryEmp" runat="server" UpdatePanelObj="udpQry" EmpIDObj="txtEmpID" EmpNameObj="txtEmpName"/>
      </div>
      <uc:UsrCalendar ID="UsrCalendar" runat="server" />
      
      <!-- 查詢輸入區畫面 --> 
      <asp:UpdatePanel ID="udpQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td align="center" style="height:12px" colspan="2">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="區長出勤統計查詢" /></td>
               </tr>
               
               <tr>
                  <td style="height:5px;width:15%" > </td>
                  <td style="height:5px;width:85%" > </td>
               </tr>
               
               <tr>
                  <td align="right" style="height:25px;"><span class="MustInput" >*</span>查詢日期：</td>
                  <td>
                     <asp:TextBox ID="txtDtFrom" runat="server" Width="88px" />
                     <img runat="server" id="imgDtFrom" src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false)" />
                     
                     &nbsp 至 &nbsp<asp:TextBox ID="txtDtTo" runat="server" Width="88px" />
                     <img runat="server" id="imgDtTo" src="../Images/date.gif" alt="按此可點選日期" style="border:0;"  onclick="ShowCalendar('txtDtTo',true,false)" />
                  </td>
               </tr>
               
               <tr>
                  <td align="right" style="height:25px;">員工職號：</td>
                  <td>
                     <asp:TextBox ID="txtEmpID" runat="server" Width="88px" AutoPostBack="true" MaxLength="8" OnTextChanged="txtEmpID_TextChanged"
                   /> <asp:TextBox ID="txtEmpName" runat="server" Width="150px" />
                  </td>
               </tr>
               
               <tr>
                 <td align="right" style="height:25px;"><span class="MustInput" >*</span>類別：</td>
                 <td><asp:RadioButtonList ID="rdoType" runat="server" RepeatColumns="2" RepeatDirection="Horizontal" RepeatLayout="Flow" >
                          <asp:ListItem Value="1" Selected="True">統計表</asp:ListItem>
                          <asp:ListItem Value="2">明細表</asp:ListItem>
                     </asp:RadioButtonList></td>
              </tr>
            </table>
         </ContentTemplate>
      </asp:UpdatePanel>
      
      <!-- 查詢結果區畫面 -->       
      <asp:UpdatePanel ID="udpGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td style="height:25px; width:15%">&nbsp</td>
                  <td style="width:85%">
                     <asp:Button ID="btnQry" runat="server" Text="查詢" OnClick="btnQry_Click"
                   /><asp:Button ID="btnClear" runat="server" Text="清除重來" OnClick="btnClear_Click"
                   /><asp:Button ID="btnExcel" runat="server" Text="EXCEL匯出" Enabled="false" OnClick="btnExcel_Click" />
                     </td>
               </tr>
               <tr>
                  <td colspan="2" align="left">
                     <asp:Label ID="lblMsg" runat="server" CssClass="msg" /></td>
               </tr>
            </table>
            <asp:GridView ID="grdList" runat="server" DataSourceID="DSCLIST" OnRowDataBound="grdList_RowDataBound" OnDataBound="grdList_DataBound" OnPageIndexChanged="grdList_PageIndexChanged" />
         </ContentTemplate>
      </asp:UpdatePanel>
   </form>
</body>
</html>
