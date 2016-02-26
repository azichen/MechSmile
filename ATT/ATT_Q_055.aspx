<%@ Page Language="VB" AutoEventWireup="false" DEBUG="true" CodeFile="ATT_Q_055.aspx.vb" Inherits="ATT_Q_055" Theme="ThemeCHG" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
   <title>加油站合理工時維護</title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
   <form id="form1" runat="server">
      <!-- 宣告 ScriptManager /SqlDataSource / 使用者控制項 -->
      <asp:ScriptManager ID="ScriptManager1" runat="server">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
      
      <asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:Smile_HQConnectionString %>" />
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
                     <asp:Label ID="TitleLabel" runat="server" CssClass="titles" Text="功能名稱"></asp:Label></td>
               </tr>
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
                  <td align="right" style="height: 25px;">加油站號：</td>
                  <td><asp:TextBox ID="txtStnID" runat="server"  Width="48px" AutoPostBack="true" MaxLength="6" /><asp:TextBox 
                                   ID="txtStnName" runat="server"  Width="150px" />
               <asp:Button ID="btnStnID" runat="server" Text="選擇單位" style="width:80px" /></td>                       
               </tr>  
            
            </table>
         </ContentTemplate>
      </asp:UpdatePanel>
      
      <!-- 查詢結果區畫面 -->       
      <asp:UpdatePanel ID="udpGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td style="height:25px; width:15%"></td>
                  <td style="width:85%">
                     <asp:Button ID="btnQry" runat="server" Text="開始查詢"
                   /><asp:Button ID="btnClear" runat="server" Text="清除重來"
                   /><asp:Button ID="btnExcel" runat="server" Text="EXCEL匯出" Enabled="False" />
                     <img alt="" src="../Images/teacher_icon.gif" style="border: 0" />
                     <asp:LinkButton ID="btnNew" runat="server" Text="新增一筆資料" Style="font-size: medium" /></td>
               </tr>
               <tr>
                  <td colspan="2" align="left">
                     <asp:Label ID="lblMsg" runat="server" CssClass="msg" /></td>
               </tr>
            </table>
            <asp:GridView ID="grdList" runat="server" />
         </ContentTemplate>
      </asp:UpdatePanel>
   </form>
</body>
</html>
