<%@ Page Language="VB" AutoEventWireup="false" DEBUG="true" CodeFile="ATT_M_057.aspx.vb" Inherits="ATT_M_057" Theme="ThemeCHG" %>
<%@ Register TagPrefix="ucCHG" TagName="ATT_M_057U"  Src="~/ATT/ATT_M_057U.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>表單申請資料維護</title>
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
         DeleteCommand="UPDATE HAFORMAPPLY SET HAP_DELETOR=@DELETOR,HAP_DELETEDATE=GETDATE() WHERE HAP_Serid=@SERID" 
         InsertCommand="INSERT INTO HAFORMAPPLY (HAP_Stnid,HAP_Empid,HAP_Formid,HAP_Status,HAP_Creator,HAP_ApplyDate,HAP_ApplyTime,HAP_Deletor,HAP_DeleteDate,HAP_REFSTRA,HAP_REFSTRB,HAP_REFSTRC,HAP_REFSTRD,HAP_REFSTRE,HAP_Contacts) VALUES (@STNID,@EMPID,@FORMID,'N',@CREATOR,@APPLYDATE,@APPLYTIME,'','',@REFSTRA,@REFSTRB,@REFSTRC,@REFSTRD,@REFSTRE,@Contacts)"
         SelectCommand="SELECT HAP_STNID,STNNAME,HAP_SERID,HAP_EMPID AS EMPID,M.EMPL_NAME AS EMPNAME,(HAP_FORMID+' '+FORM_NAME) AS FORMNAME,HAP_Contacts AS Contacts,HAP_APPLYDATE,CASE HAP_STATUS WHEN 'N' THEN '未處理' WHEN 'Y' THEN '已處理' WHEN 'I' THEN '不處理' END  AS STATUS,HAP_REFSTRA,HAP_REFSTRB,HAP_REFSTRC,HAP_REFSTRD,HAP_REFSTRE FROM HAFORMApply A LEFT JOIN [MP_HR].DBO.EMPLOYEE M ON A.HAP_EMPID=M.EMPL_ID LEFT JOIN MECHSTNM S ON A.HAP_STNID=S.STNID LEFT JOIN HAFORM F ON A.HAP_FORMID=F.FORM_ID WHERE HAP_SERID=@SERID "
         DataSourceMode="DataReader">
         
            <DeleteParameters>
                <asp:Parameter Name="DELETOR" Type="String" />
                <asp:Parameter Name="SERID" Type="Int32" />
            </DeleteParameters>
 
            <InsertParameters>
                <asp:Parameter Name="STNID" Type="String" />
                <asp:Parameter Name="EMPID" Type="String" />
                <asp:Parameter Name="FORMID" Type="String" />
                <asp:Parameter Name="CREATOR" Type="String" />
                <asp:Parameter Name="APPLYDATE" Type="String" />
                <asp:Parameter Name="APPLYTIME" Type="String" />
                <asp:Parameter Name="REFSTRA"  Type="String" />
                <asp:Parameter Name="REFSTRB" Type="String" />
                <asp:Parameter Name="REFSTRC" Type="String" />
                <asp:Parameter Name="REFSTRD" Type="String" />
                <asp:Parameter Name="REFSTRE" Type="String" />
                <asp:Parameter Name="Contacts" Type="String" />
            </InsertParameters>
                       
            <SelectParameters>
                <asp:QueryStringParameter Name="SERID" QueryStringField="SERID" Type="INT32" />
            </SelectParameters>
            
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
               <ucCHG:ATT_M_057U ID="ATT_M_057U" runat="server"/>
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
               <ucCHG:ATT_M_057U ID="ATT_M_057U" runat="server"/>
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
