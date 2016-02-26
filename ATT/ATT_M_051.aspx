<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ATT_M_051.aspx.vb" Inherits="ATT_M_051" Theme="ThemeCHG" %>
<%@ Register TagPrefix="ucCHG" TagName="ATT_M_051U"  Src="~/ATT/ATT_M_051U.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>排班資料維護</title>
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
         DeleteCommand="DELETE FROM SCHEDM WHERE STNID=@STNID AND EMPID=@EMPID AND SHSTDATE=@SHSTDATE and SHSTTIME=@SHSTTIME" 
         InsertCommand="INSERT INTO SCHEDM (EMPID,STNID,SHSTDATE,SHSTTIME,SHEDDATE,SHEDTIME,CLASSID,WORKHOUR,RtSTTIME,RtEDTIME,FACFLAG,inuser,indate,intime,WKSTTIME,WKEDTIME,WKSTMARK,WKEDMARK,NetWH,DayOver,TWKOver,HODWH,OVTMSEQ,OVTMINYM) VALUES (@EMPID,@STNID,@SHSTDATE,@SHSTTIME,@SHEDDATE,@SHEDTIME,@CLASSID,@WORKHOUR,@RtSTTIME,@RtEDTIME,'N',@inuser,@indate,@intime,'','','','',0,0,0,0,0,'')"
         SelectCommand="SELECT A.STNID,STNNAME,EMPID,B.EMPL_NAME AS EMPNAME,SHSTDATE,SHSTTIME,SHEDDATE,SHEDTIME,WORKHOUR,RTSTTIME,RTEDTIME FROM SCHEDM A INNER JOIN [MP_HR].DBO.EMPLOYEE B ON A.EMPID=B.EMPL_ID INNER JOIN MECHSTNM C ON A.STNID=C.STNID WHERE A.STNID=@STNID and EMPID=@EMPID and SHSTDATE=@SHSTDATE and SHSTTIME=@SHSTTIME "
         UpdateCommand="UPDATE SCHEDM SET SHSTDATE=@SHSTDATE,SHSTTIME=@SHSTTIME,SHEDDATE=@SHEDDATE,SHEDTIME=@SHEDTIME,WORKHOUR=@WORKHOUR,RTSTTIME=@RTSTTIME,RTEDTIME=@RTEDTIME,INUSER=@INUSER,INDATE=@INDATE,INTIME=@INTIME,FACFLAG='N' WHERE STNID=@STNID AND EMPID=@EMPID AND SHSTDATE=@VSTDATE AND SHSTTIME=@VSTTIME"
         DataSourceMode="DataReader">
            <DeleteParameters>
                <asp:Parameter Name="STNID" Type="String" />
                <asp:Parameter Name="EMPID" Type="String" />
                <asp:Parameter Name="SHSTDATE" Type="String" />
                <asp:Parameter Name="SHSTTIME" Type="String" />
            </DeleteParameters>
 
            <SelectParameters>
                <asp:QueryStringParameter Name="STNID"    QueryStringField="STNID" Type="String" />
                <asp:QueryStringParameter Name="EMPID"    QueryStringField="EMPID" Type="String" />
                <asp:QueryStringParameter Name="SHSTDATE" QueryStringField="SHSTDATE" Type="String" />
                <asp:QueryStringParameter Name="SHSTTIME" QueryStringField="SHSTTIME" Type="String" />        
            </SelectParameters>
            
            <InsertParameters>
                <asp:Parameter Name="STNID" Type="String" />
                <asp:Parameter Name="EMPID" Type="String" />
                <asp:Parameter Name="SHSTDATE" Type="String" />
                <asp:Parameter Name="SHSTTIME" Type="String" />
                <asp:Parameter Name="SHEDDATE" Type="String" />
                <asp:Parameter Name="SHEDTIME" Type="String" />
                <asp:Parameter Name="CLASSID"  Type="String" />
                <asp:Parameter Name="WORKHOUR" Type="String" />
                <asp:Parameter Name="RtSTTIME" Type="String" />
                <asp:Parameter Name="RtEDTIME" Type="String" />
                <asp:Parameter Name="inuser" Type="String" />
                <asp:Parameter Name="indate" Type="String" />
                <asp:Parameter Name="intime" Type="String" />
            </InsertParameters>
            
            <UpdateParameters>
                <asp:Parameter Name="STNID" Type="String" />
                <asp:Parameter Name="EMPID" Type="String" />
                <asp:Parameter Name="SHSTDATE" Type="String" />
                <asp:Parameter Name="SHSTTIME" Type="String" />
                <asp:Parameter Name="SHEDDATE" Type="String" />
                <asp:Parameter Name="SHEDTIME" Type="String" />
                <asp:Parameter Name="CLASSID"  Type="String" />
                <asp:Parameter Name="WORKHOUR" Type="String" />
                <asp:Parameter Name="RtSTTIME" Type="String" />
                <asp:Parameter Name="RtEDTIME" Type="String" />
                <asp:Parameter Name="INUSER" Type="String" />
                <asp:Parameter Name="INDATE" Type="String" />
                <asp:Parameter Name="INTIME" Type="String" />
                
                <asp:Parameter Name="VSTDATE" Type="String" />
                <asp:Parameter Name="VSTTIME" Type="String" />
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
               <ucCHG:ATT_M_051U ID="ATT_M_051U" runat="server"/>
               <table class="table">
                 <tr> <td style="height:50px"></td></tr>
                 <tr>
                   <td align="center">
                    <asp:Button ID="btnEdit" runat="server" Text="修改" CommandName="Edit" /><asp:Button
                                ID="btnDel" runat="server" Text="刪除" OnClick="btnFunc_Click" 
                                    OnClientClick="return confirm('確定要刪除此筆資料嗎?')" /><asp:Button 
                                ID="btnPrn" runat="server" Text="列印" Enabled="false" /><asp:Button 
                                ID="btnBack" runat="server" Text="回主畫面"  OnClick="btnFunc_Click" />                   
                   </td>
                 </tr>
               </table>
            </ItemTemplate>
            
            <EditItemTemplate>
                  <ucCHG:ATT_M_051U ID="ATT_M_051U" runat="server"/>
                  <table class="table">
                    <tr>
                       <td align="center">
                          <asp:Button ID="btnMod" runat="server" Text="資料送出" CommandName="Update" 
                        /><asp:Button ID="btnCancel" runat="server"  Text="取消修改" CommandName="Cancel" />                  
                       </td>
                    </tr>
                  </table>
               </EditItemTemplate>  
            
            <InsertItemTemplate>
               <ucCHG:ATT_M_051U ID="ATT_M_051U" runat="server"/>
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
