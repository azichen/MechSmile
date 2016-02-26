<%@ Page Language="VB" AutoEventWireup="false" Debug="true" CodeFile="ATT_M_059.aspx.vb" Inherits="ATT_M_059" Theme="ThemeCHG" %>
<%@ Register TagPrefix="ucCHG" TagName="ATT_M_059U"  Src="~/ATT/ATT_M_059U.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>出勤雙週週期設定維護</title>
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" /> 
</head>

<body onkeydown="if(event.keyCode==13)event.keyCode=9">
  <form runat="server" id="form1" >
    <div style="width:500px">
    
<!-- 宣告 ScriptManager /SqlDataSource / 使用者控制項 -->
      <asp:ScriptManager ID="ScriptManager1" runat="server">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>     
    
     <asp:SqlDataSource ID="dscMain" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>"
         SelectCommand="SELECT FWDS_DATEST,FWDS_DATEEN,FWDS_STNWH FROM FWDUTYSH WHERE FWDS_DATEST= @FWDS_DATEST" 
         UpdateCommand="UPDATE FWDUTYSH SET FWDS_STNWH = @FWDS_STNWH WHERE FWDS_DATEST= @FWDS_DATEST" DataSourceMode="DataReader">
            <SelectParameters>
                <asp:QueryStringParameter Name="FWDS_DATEST"  QueryStringField="FWDS_DATEST" Type="String" />
            </SelectParameters>
            
            <UpdateParameters>
                <asp:Parameter Name="FWDS_DATEST" Type="String" />
                <asp:Parameter Name="FWDS_STNWH" Type="Double" />
            </UpdateParameters>
            
     </asp:SqlDataSource>
        
     <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" >
        <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
     </div>
     
     <uc:UsrCalendar ID="UsrCalendar1" runat="server" /> 

     <!-- 宣告 FormView 流程控制主體 -->  
     <table class="table" >
      <tr> <td style="height:30px"></td></tr>   
      <tr align="center" >
        <td style="height: 30px">
          <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="功能名稱"></asp:Label></td>
      </tr>    
      <tr> <td style="height:15px"></td></tr>  
      <tr>
        <td align="left" style="height: 25px; background-color:#fffbd6">
          <asp:Label ID="lblMsg" runat="server" CssClass="msg" /><br />
          <asp:FormView ID="fmvEdit" runat="server" DataKeyNames="FWDS_DATEST" DataSourceID="dscMain" >
            <ItemTemplate>
               <ucCHG:ATT_M_059U ID="ATT_M_059U" runat="server"/>
               <table class="table">
                 <tr> <td style="height:50px"></td></tr>
                 <tr>
                   <td align="center">
                    <asp:Button ID="btnEdit" runat="server" Text="修改" CommandName="Edit" /><asp:Button 
                                ID="btnBack" runat="server" Text="回主畫面"  OnClick="btnFunc_Click" />                   
                   </td>
                 </tr>
               </table>
            </ItemTemplate>
            
            <EditItemTemplate>
               <ucCHG:ATT_M_059U ID="ATT_M_059U" runat="server"/>
               <table class="table">
                 <tr> <td style="height:10px"></td></tr>
                 <tr>
                   <td align="center">
                     <asp:Button ID="btnMod" runat="server" Text="資料送出" CommandName="Update" /><asp:Button 
                                 ID="btnCancel" runat="server"  Text="取消設定" CommandName="Cancel" />                  
                   </td>
                 </tr>
               </table>
             </EditItemTemplate>   
            
            
        </asp:FormView> </td>
      </tr>
     </table>
    </div> 
  </form>
</body>
</html>
