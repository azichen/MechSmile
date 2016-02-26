<%@ Page Language="C#" debug="false" AutoEventWireup="true" CodeFile="~/ATT/ATT_Q_012.aspx.cs" Inherits="ATT_Q_012" EnableEventValidation = "false" Theme="ThemeCHG" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
   <title>PT申報加班時數檢核表</title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />  
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
      <asp:SqlDataSource ID="dscBus" runat="server" ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>"
           SelectCommand="SELECT RTrim(Bus_ID) As Bus_ID, Bus_Desc FROM [Stn_Bus] Order By Bus_ID" >
      </asp:SqlDataSource>        
      <asp:SqlDataSource ID="dscArea" runat="server" ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>"
           SelectCommand="SELECT RTrim(Are_ID) As Are_ID, Are_Desc FROM [Stn_Area] WHERE Bus_ID = @BusID Order By Are_ID" OnSelecting="dscArea_Selecting" >
           <SelectParameters>
              <asp:ControlParameter ControlID="ddlBus" Name="BusID" PropertyName="SelectedValue" Type="String" />
           </SelectParameters>
      </asp:SqlDataSource>  
              
      <div style="width:10px; height:10px; position:absolute; left:-200px; top:-10px;" >
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
         <uc:QryStn ID="QryStn" runat="server" UpdatePanelObj="udpQry" StnIDObj="txtStnID" StnNameObj="txtStnName"/>
         <uc:QryEmp ID="QryEmp" runat="server" UpdatePanelObj="udpQry1" EmpIDObj="txtEmpID" EmpNameObj="txtEmpName"/>
      </div>
      
      <uc:UsrCalendar ID="UsrCalendar" runat="server" />
      
      <!-- 查詢輸入區畫面 --> 
      <asp:UpdatePanel ID="udpQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td align="center" style="height:12px" colspan="2">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="PT申報加班時數檢核表" /></td>
               </tr>
               <tr>
                  <td style="height:5px;width:15%" > </td>
                  <td style="height:5px;width:85%" > </td>
               </tr>
         
               <tr>
                  <td align="right" style="height:25px;">業務處：</td>
                  <td>
                     <asp:DropDownList ID="ddlBus" runat="server" AppendDataBoundItems="true" AutoPostBack="true"
                          DataSourceID="dscBus" DataTextField="Bus_Desc" DataValueField="Bus_ID" OnDataBound="ddlArea_DataBound" OnSelectedIndexChanged="ddlBus_SelectedIndexChanged" >
            　　　        <asp:ListItem Text="" />
              　　   </asp:DropDownList>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
             責任區：<asp:DropDownList ID="ddlArea" runat="server" AppendDataBoundItems="true"
                          DataSourceID="dscArea" DataTextField="Are_Desc" DataValueField="Are_ID" OnDataBound="ddlArea_DataBound" OnSelectedIndexChanged="ddlBus_SelectedIndexChanged" >
            　 　         <asp:ListItem Text="" />
                     </asp:DropDownList></td>
               </tr>

               <tr>
                  <td align="right" style="height:25px;"><span class="MustInput" >*</span>加油站號：</td>
                  <td>
                     <asp:TextBox ID="txtStnID" runat="server" Width="48px" AutoPostBack="true" MaxLength="6" OnTextChanged="txtStnID_TextChanged"
                   /><asp:TextBox ID="txtStnName" runat="server" Width="150px" />
                     <asp:Button ID="btnStnID" runat="server" Text="選擇單位" style="width:80px" OnClick="btnStnID_Click" /></td>
               </tr>
               <tr>
                <td align="right" style="height: 25px; width: 120px;"><span class="MustInput" >*</span>查詢月份：</td>
                <td style="height: 25px" colspan="4">
                  <asp:TextBox ID="txtDtStart" runat="server" Width="60px" MaxLength="7" OnTextChanged="txtDtStart_TextChanged" ></asp:TextBox>
                  <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtStart',true,true)" /></td>
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
               <tr>
                  <td colspan="2" align="left">
                  <asp:Label ID="lblMsg2" runat="server" ></asp:Label>
                  </td>
               </tr>
            </table>
            <asp:GridView ID="grdList" runat="server" DataSourceID="DSCLIST" OnRowDataBound="grdList_RowDataBound" />
         </ContentTemplate>
      </asp:UpdatePanel>
   </form>
</body>
</html>
