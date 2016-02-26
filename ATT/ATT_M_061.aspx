<%@ Page Language="VB" AutoEventWireup="false" Debug="true" CodeFile="ATT_M_061.aspx.vb" Inherits="ATT_M_061" Theme="ThemeCHG" %>
<%@ Register TagPrefix="ucCHG" TagName="ATT_M_061U"  Src="~/ATT/ATT_M_061U.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>颱風假設定維護</title>
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
    <!-- DeleteCommand="DELETE FROM HOLYSETM WHERE HODDATE= @HODDATE AND HDTMST=@HDTMST AND HDTMEN=@HDTMEN AND STNFLAG='1' "  -->
     <asp:SqlDataSource ID="dscMain" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>"
         DeleteCommand="UPDATE HOLYSETM SET DELFLAG='Y' WHERE HODDATE= @HODDATE AND HDTMST=@HDTMST AND HDTMEN=@HDTMEN AND STNFLAG='1' "
         InsertCommand="INSERT INTO HOLYSETM (HODDATE,HDTMST,HDTMEN,HODHOURS,HODMEMO,CITYSTR,STNFLAG,MODUSER,MODDT,DELFLAG) VALUES (@HODDATE,@HDTMST,@HDTMEN,@HODHOURS,@HODMEMO,@CITYSTR,'1',@MODUSER,GETDATE(),'')" 
         SelectCommand="SELECT HODDATE,HDTMST,HDTMEN,HODHOURS,HODMEMO,CITYSTR FROM HOLYSETM WHERE HODDATE= @HODDATE AND HDTMST=@HDTMST AND HDTMEN=@HDTMEN AND STNFLAG='1'" 
         UpdateCommand="UPDATE HOLYSETM SET HODHOURS=@HODHOURS,HODMEMO=@HODMEMO,CITYSTR=@CITYSTR,MODUSER=@MODUSER,MODDT=GETDATE() WHERE HODDATE=@HODDATE AND HDTMST=@HDTMST AND HDTMEN=@HDTMEN AND STNFLAG='1'" 
         DataSourceMode="DataReader">
            <DeleteParameters>
                <asp:Parameter Name="HODDATE" Type="String" />
                <asp:Parameter Name="HDTMST" Type="String" />
                <asp:Parameter Name="HDTMEN" Type="String" />
            </DeleteParameters>
            
            <InsertParameters>
                <asp:Parameter Name="HODDATE" Type="String" />
                <asp:Parameter Name="HDTMST" Type="String" />
                <asp:Parameter Name="HDTMEN" Type="String" />
                <asp:Parameter Name="HODHOURS"   Type="Double" />
                <asp:Parameter Name="HODMEMO" Type="String" />
                <asp:Parameter Name="CITYSTR" Type="String" />
                <asp:Parameter Name="MODUSER" Type="String" />
            </InsertParameters>
         
            <SelectParameters>
                <asp:QueryStringParameter Name="HODDATE"  QueryStringField="HODDATE" Type="String" />
                <asp:QueryStringParameter Name="HDTMST"  QueryStringField="HDTMST" Type="String" />
                <asp:QueryStringParameter Name="HDTMEN"  QueryStringField="HDTMEN" Type="String" />
            </SelectParameters>
            
            <UpdateParameters>
                <asp:Parameter Name="HODDATE" Type="String" />
                <asp:Parameter Name="HDTMST" Type="String" />
                <asp:Parameter Name="HDTMEN" Type="String" />
                <asp:Parameter Name="HODHOURS"   Type="Double" />
                <asp:Parameter Name="CITYSTR" Type="String" />
                <asp:Parameter Name="HODMEMO" Type="String" />
                <asp:Parameter Name="MODUSER" Type="String" />
            </UpdateParameters>                     
            
     </asp:SqlDataSource>
        
     <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" >
        <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
     </div>
     
     <uc:UsrCalendar ID="UsrCalendar1" runat="server" /> 

     <!-- 宣告 FormView 流程控制主體 -->  
     <!-- <td align="center">
                    <asp:Button ID="btnEdit" runat="server" Text="修改" CommandName="Edit" /><asp:Button
                                ID="btnDel" runat="server" Text="刪除" OnClick="btnFunc_Click" 
                                    OnClientClick="return confirm('確定要刪除此筆資料嗎?')" /><asp:Button 
                                ID="btnBack" runat="server" Text="回主畫面"  OnClick="btnFunc_Click" />                   
                   </td>  -->  
                   
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
          <asp:FormView ID="fmvEdit" runat="server" DataKeyNames="HODDATE" DataSourceID="dscMain" >
            <ItemTemplate>
               <ucCHG:ATT_M_061U ID="ATT_M_061U" runat="server"/>
               <table class="table">
                 <tr> <td style="height:50px"></td></tr>
                 <tr>
                   <td align="center">
                    <asp:Button ID="btnDel" runat="server" Text="刪除" OnClick="btnFunc_Click" 
                                    OnClientClick="return confirm('確定要刪除此筆資料嗎?')" /><asp:Button 
                                ID="btnBack" runat="server" Text="回主畫面"  OnClick="btnFunc_Click" />                   
                   </td>
                 </tr>
               </table>
            </ItemTemplate>
            
            <EditItemTemplate>
               <ucCHG:ATT_M_061U ID="ATT_M_061U" runat="server"/>
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
            
             <InsertItemTemplate>
               <ucCHG:ATT_M_061U ID="ATT_M_061U" runat="server"/>
               <table class="table">
                 <tr> <td style="height:50px"></td></tr>
                 <tr>
                   <td align="center">
                    <asp:Button ID="btnIns" runat="server" Text="資料送出" CommandName="Insert" /><asp:Button 
                                ID="btnClear" runat="server" Text="清除重來" OnClick="btnFunc_Click" /><asp:Button 
                                ID="btnBack" runat="server" Text="回主畫面" OnClick="btnFunc_Click" />                  
                   </td>
                 </tr>
               </table>
            </InsertItemTemplate>
            
        </asp:FormView> </td>
      </tr>
     </table>
    </div> 
  </form>
</body>
</html>
