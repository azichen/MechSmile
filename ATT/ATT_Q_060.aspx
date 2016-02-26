<%@ Page Language="VB" AutoEventWireup="false" DEBUG="true" CodeFile="ATT_Q_060.aspx.vb" Inherits="ATT_Q_060" Theme="ThemeCHG" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
   <title>考勤鎖檔處理作業</title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
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
      <asp:SqlDataSource ID="dscBus" runat="server" ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>"
           SelectCommand="SELECT RTrim(Bus_ID) As Bus_ID, Bus_Desc FROM [Stn_Bus] Order By Bus_ID">
      </asp:SqlDataSource>
      <asp:SqlDataSource ID="dscArea" runat="server" ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>"
           SelectCommand="SELECT RTrim(Are_ID) As Are_ID, Are_Desc FROM [Stn_Area] WHERE Bus_ID = @BusID Order By Are_ID">
           <SelectParameters>
              <asp:ControlParameter ControlID="ddlBus" Name="BusID" PropertyName="SelectedValue" Type="String" />
           </SelectParameters>
      </asp:SqlDataSource>
      
      <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" >  
        <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
        <uc:QryStn ID="QryStn" runat="server" UpdatePanelObj="udpQry" StnIDObj="txtStnID" StnNameObj="txtStnName" />     
      </div>

      <!-- 查詢輸入區畫面 --> 
      <asp:UpdatePanel ID="udpQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td align="center" colspan="2">
                     <asp:Label ID="TitleLabel" runat="server" CssClass="titles" Text="考勤鎖檔設定作業"></asp:Label></td>
               </tr>
               
            </table>
         </ContentTemplate>
      </asp:UpdatePanel>
      
      <uc:UpdatePanelFix ID="updGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
            
                <tr>
                  <td align="right" style="height:25px;">業務處：</td>
                  <td>
                     <asp:DropDownList ID="ddlBus" runat="server" AppendDataBoundItems="true" AutoPostBack="true"
                          DataSourceID="dscBus" DataTextField="Bus_Desc" DataValueField="Bus_ID" >
            　　　        <asp:ListItem Text="" />
              　　   </asp:DropDownList>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
             責任區：<asp:DropDownList ID="ddlArea" runat="server" AppendDataBoundItems="true"
                          DataSourceID="dscArea" DataTextField="Are_Desc" DataValueField="Are_ID" >
            　 　         <asp:ListItem Text="" />
                     </asp:DropDownList></td>
               </tr>
                <tr>
                  <td align="right" style="height: 32px;">加油站號：</td>
                  <td style="height: 32px"><asp:TextBox ID="txtStnID" runat="server"  Width="48px" AutoPostBack="true" MaxLength="6" /><asp:TextBox 
                                   ID="txtStnName" runat="server"  Width="150px" />
               <asp:Button ID="btnStnID" runat="server" Text="選擇單位" style="width:80px" /></td>                       
              </tr> 
            
               <tr>
                 <td align="right" style="height:27px;">週期別：</td>
                 <td style="height: 27px">
                     &nbsp;<asp:Button ID="Button3" runat="server" ForeColor="DarkGreen" OnClick="Button3_Click"
                         Text="計薪月" Font-Bold="True" />&nbsp;
                     <asp:Button ID="Button4" runat="server" ForeColor="DimGray" OnClick="Button4_Click"
                         Text="雙週" Font-Bold="False" />
                     &nbsp;
                     <asp:RadioButtonList ID="rdoType" runat="server" RepeatColumns="2" RepeatDirection="Horizontal" RepeatLayout="Flow" Visible="False" >
                          <asp:ListItem Value="1" Selected="True">計薪月</asp:ListItem>
                          <asp:ListItem Value="2">雙週</asp:ListItem>
                     </asp:RadioButtonList></td>
               </tr>
               <tr>
                  <td align="right" style="height: 32px;">起始日期：</td>
                  <td style="height: 32px">
                      <asp:Label ID="LblDATEST" runat="server" Text="2015/06/22" Font-Bold="True" Font-Size="Large"></asp:Label>
                      <asp:Button ID="Button1" runat="server" Text="前一週期" style="width:80px" OnClick="Button1_Click" /> <asp:Label ID="LblACCYM" runat="server" Font-Bold="True"
                     Font-Size="Large" Text="2015年7月"></asp:Label></td>                       
               </tr> 
               <tr>
                  <td align="right" style="height: 32px;">迄止日期：</td>
                  <td style="height: 32px">
                      <asp:Label ID="LblDATEEN" runat="server" Text="2015/07/05" Font-Bold="True" Font-Size="Large"></asp:Label>
                      <asp:Button ID="Button2" runat="server" Text="下一週期" style="width:80px" OnClick="Button2_Click" />
                      <asp:Label ID="LblFWDSID" runat="server" Text="Label" Visible="False"></asp:Label></td>                       
               </tr> 
               <tr>
                  <td align="right" style="height: 32px;">狀態：</td>
                  <td style="height: 32px">
                      <asp:Label ID="LblSTATUS" runat="server" Text="已鎖檔(總部)" Font-Bold="True" Font-Size="Large" ForeColor="Red"></asp:Label></td>                       
               </tr> 
               <tr>
                  <td style="height:25px; width:15%"></td>
                  <td style="width:85%">
                     <asp:Button ID="btnLOCK" runat="server" Text="鎖檔" OnClick="btnLOCK_Click" OnClientClick="return confirm('麻煩「再次」確認，是否鎖定? ');" 
                   /><asp:Button ID="btnUNLOCK" runat="server" Text="解鎖檔" OnClick="btnUNLOCK_Click"
                   />&nbsp;
                  </td>
               </tr>
      
               <tr>
                  <td colspan="2" align="left">
                  <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label></td>
               </tr>
            </table>
            <asp:GridView ID="grdList2" runat="server" />
         </ContentTemplate>      
      </uc:UpdatePanelFix>     
   </form>
</body>
</html>
