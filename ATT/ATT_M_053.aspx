<%@ Page Language="VB" AutoEventWireup="false" Debug="true"  CodeFile="ATT_M_053.aspx.vb" Inherits="ATT_M_053" Theme="ThemeCHG" %>
<%@ Register TagPrefix="ucCHG" TagName="ATT_M_053U"  Src="~/ATT/ATT_M_053U.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>打卡資料維護</title>
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
         DeleteCommand="DELETE FROM FINGER WHERE STNID=@STNID AND EMPID=@EMPID AND FINDATE=@FINDATE and FINTIME=@FINTIME" 
         InsertCommand="INSERT INTO FINGER (STNID,EMPID,FINDATE,FINTIME,INUSER,INDATE,INTIME,FINCLA) VALUES (@STNID,@EMPID,@FINDATE,@FINTIME,@INUSER,@INDATE,@INTIME,@FINCLA)" 
         SelectCommand="SELECT A.STNID,STNNAME,A.EMPID,ISNULL(B.EMPL_NAME,'') AS EMPNAME,FINDATE,FINTIME,MEMO FROM FINGER A INNER JOIN MECHSTNM C ON A.STNID=C.STNID LEFT JOIN [MP_HR].DBO.EMPLOYEE B ON A.EMPID=B.EMPL_ID LEFT JOIN SCHPASSM S ON S.EMPID=A.EMPID AND S.SHSTDATE=A.FINDATE AND S.SHSTTIME=SUBSTRING(A.FINTIME,1,4) AND S.PASSFG='A' WHERE A.STNID=@STNID AND A.EMPID=@EMPID AND FINDATE=@FINDATE AND FINTIME=@FINTIME " 
         UpdateCommand="UPDATE FINGER SET EMPID=@CHGEMPID WHERE STNID=@STNID AND EMPID=@EMPID AND FINDATE=@FINDATE AND FINTIME=@FINTIME AND SUBSTRING(FINCLA,1,1)='0'" 
         DataSourceMode="DataReader">
            <DeleteParameters>
                <asp:Parameter Name="STNID" Type="String" />
                <asp:Parameter Name="EMPID" Type="String" />
                <asp:Parameter Name="FINDATE" Type="String" />
                <asp:Parameter Name="FINTIME" Type="String" />
            </DeleteParameters>
 
            <SelectParameters>
                <asp:QueryStringParameter Name="STNID"    QueryStringField="STNID" Type="String" />
                <asp:QueryStringParameter Name="EMPID"    QueryStringField="EMPID" Type="String" />
                <asp:QueryStringParameter Name="FINDATE"  QueryStringField="FINDATE" Type="String" />
                <asp:QueryStringParameter Name="FINTIME"  QueryStringField="FINTIME" Type="String" />        
            </SelectParameters>
            
            <InsertParameters>
                <asp:Parameter Name="STNID"   Type="String" />
                <asp:Parameter Name="EMPID"   Type="String" />
                <asp:Parameter Name="FINDATE" Type="String" />
                <asp:Parameter Name="FINTIME" Type="String" />
                <asp:Parameter Name="FINCLA"  Type="String" />
                <asp:Parameter Name="INUSER"  Type="String" />
                <asp:Parameter Name="INDATE"  Type="String" />
                <asp:Parameter Name="INTIME"  Type="String" />
            </InsertParameters>
            
            <UpdateParameters>
                <asp:Parameter Name="STNID" Type="String" />
                <asp:Parameter Name="EMPID" Type="String" />
                <asp:Parameter Name="FINDATE" Type="String" />
                <asp:Parameter Name="FINTIME" Type="String" />
                <asp:Parameter Name="CHGEMPID" Type="String" />
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
          <asp:FormView ID="fmvEdit" runat="server" DataKeyNames="EMPID" DataSourceID="dscMain" >
            <ItemTemplate>
               <ucCHG:ATT_M_053U ID="ATT_M_053U" runat="server"/>
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
               <ucCHG:ATT_M_053U ID="ATT_M_053U" runat="server"/>
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
            
            <EditItemTemplate>
               <ucCHG:ATT_M_053U ID="ATT_M_053U" runat="server"/>
               <table class="table">
                 <tr> <td style="height:50px"></td></tr>
                 <tr>
                   <td align="center">
                    <asp:Button ID="btnEdit" runat="server" Text="員工代號變更" CommandName="Update" 
                                    OnClientClick="return confirm('確定要變更此員工代號嗎?')" /><asp:Button 
                                ID="btnBack" runat="server" Text="回主畫面" OnClick="btnFunc_Click" />                  
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
