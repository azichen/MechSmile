<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ATT_M_054.aspx.vb" Inherits="ATT_M_054" Theme="ThemeCHG" %>
<%@ Register TagPrefix="ucCHG" TagName="ATT_M_054U"  Src="~/ATT/ATT_M_054U.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>排班代號維護</title>
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
         DeleteCommand="DELETE FROM CLSTIME WHERE STNID=@STNID AND CLSID=@CLSID " 
         InsertCommand="INSERT INTO CLSTIME (STNID,CLSID,STTIME,EDTIME,NFLAG) VALUES (@STNID,@CLSID,@STTIME,@EDTIME,@NFLAG) " 
         SelectCommand="SELECT A.STNID,STNNAME,CLSID,STTIME,EDTIME,NFLAG FROM CLSTIME A INNER JOIN MECHSTNM C ON A.STNID=C.STNID WHERE A.STNID=@STNID AND CLSID=@CLSID " 
         UpdateCommand="" DataSourceMode="DataReader">
            <DeleteParameters>
                <asp:Parameter Name="STNID" Type="String" />
                <asp:Parameter Name="CLSID" Type="String" />
            </DeleteParameters>
 
            <SelectParameters>
                <asp:QueryStringParameter Name="STNID"    QueryStringField="STNID" Type="String" />
                <asp:QueryStringParameter Name="CLSID"    QueryStringField="CLSID" Type="String" />
            </SelectParameters>
            
            <InsertParameters>
                <asp:Parameter Name="STNID"   Type="String" />
                <asp:Parameter Name="CLSID"   Type="String" />
                <asp:Parameter Name="FINDATE" Type="String" />
                <asp:Parameter Name="STTIME"  Type="String" />
                <asp:Parameter Name="EDTIME"  Type="String" />
                <asp:Parameter Name="NFLAG"   Type="String" />
            </InsertParameters>
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
          <asp:FormView ID="fmvEdit" runat="server" DataKeyNames="CLSID" DataSourceID="dscMain" >
            <ItemTemplate>
               <ucCHG:ATT_M_054U ID="ATT_M_054U" runat="server"/>
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
            
            <InsertItemTemplate>
               <ucCHG:ATT_M_054U ID="ATT_M_054U" runat="server"/>
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
