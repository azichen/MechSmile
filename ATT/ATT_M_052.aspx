<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ATT_M_052.aspx.vb" Inherits="ATT_M_052" Theme="ThemeCHG" %>
<%@ Register TagPrefix="ucCHG" TagName="ATT_M_052U"  Src="~/ATT/ATT_M_052U.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>請假資料維護</title>
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
         DeleteCommand="DELETE FROM MP_HR.DBO.VACATM WHERE VATM_EMPLID=@VATM_EMPLID AND VATM_SEQ=@VATM_SEQ " 
         InsertCommand="INSERT INTO MP_HR.DBO.VACATM (VATM_COMPID,VATM_EMPLID,VATM_SEQ,VATM_VANMID,VATM_DATE_ST,VATM_DATE_EN,VATM_TIME_ST,VATM_TIME_EN,VATM_DAYS,VATM_HOURS,VATM_REMARK,VATM_INYM,VATM_A_USER_ID,VATM_A_USER_NM,VATM_A_DATE,VATM_U_USER_ID,VATM_U_USER_NM,VATM_U_DATE) VALUES ('01',@VATM_EMPLID,@VATM_SEQ,@VATM_VANMID,@VATM_DATE_ST,@VATM_DATE_EN,@VATM_TIME_ST,@VATM_TIME_EN,0,@VATM_HOURS,@VATM_REMARK,@VATM_INYM,@VATM_A_USER_ID,@VATM_A_USER_NM,@VATM_A_DATE,@VATM_U_USER_ID,@VATM_U_USER_NM,@VATM_U_DATE)" 
         SelectCommand="SELECT VATM_COMPID,VATM_EMPLID,EMPL_NAME,VATM_SEQ,VATM_VANMID,CONVERT(char(10), VACA_NAME) AS VATM_NAME,CONVERT(char(8), VATM_DATE_ST,112) AS VATM_DATE_ST,CONVERT(char(8), VATM_DATE_EN,112) AS VATM_DATE_EN,VATM_TIME_ST,VATM_TIME_EN,VATM_DAYS,VATM_HOURS,VATM_REMARK,VATM_A_USER_ID,VATM_A_USER_NM,VATM_A_DATE FROM MP_HR.DBO.VACATM A INNER JOIN MP_HR.DBO.EMPLOYEE B ON A.VATM_EMPLID=B.EMPL_ID INNER JOIN MP_HR.DBO.VACAMF C ON A.VATM_VANMID=C.VACA_ID WHERE VATM_EMPLID=@VATM_EMPLID AND VATM_SEQ=@VATM_SEQ " 
         UpdateCommand="UPDATE MP_HR.DBO.VACATM SET VATM_DATE_ST=@VATM_DATE_ST,VATM_DATE_EN=@VATM_DATE_EN,VATM_TIME_ST=@VATM_TIME_ST,VATM_TIME_EN=@VATM_TIME_EN,VATM_DAYS=0,VATM_HOURS=@VATM_HOURS,VATM_REMARK=@VATM_REMARK,VATM_U_USER_ID=@VATM_U_USER_ID,VATM_U_USER_NM=@VATM_U_USER_NM,VATM_U_DATE=@VATM_U_DATE WHERE VATM_EMPLID=@VATM_EMPLID AND VATM_SEQ=@VATM_SEQ" 
         DataSourceMode="DataReader">
            <DeleteParameters>
                <asp:Parameter Name="VATM_EMPLID"  Type="String" />
                <asp:Parameter Name="VATM_SEQ"     Type="Int32" />
            </DeleteParameters>
 
            <SelectParameters>
                <asp:QueryStringParameter Name="VATM_EMPLID" QueryStringField="VATM_EMPLID" Type="String" />
                <asp:QueryStringParameter Name="VATM_SEQ"    QueryStringField="VATM_SEQ" Type="Int32" />
            </SelectParameters>
            
            <InsertParameters>
                <asp:Parameter Name="VATM_EMPLID"  Type="String" />
                <asp:Parameter Name="VATM_SEQ"     Type="Int32"  />
                <asp:Parameter Name="VATM_VANMID"  Type="String" />
                <asp:Parameter Name="VATM_DATE_ST" Type="String" />
                <asp:Parameter Name="VATM_DATE_EN" Type="String" />
                <asp:Parameter Name="VATM_TIME_ST" Type="String" />
                <asp:Parameter Name="VATM_TIME_EN" Type="String" />
                <asp:Parameter Name="VATM_HOURS"   Type="Double" />
                <asp:Parameter Name="VATM_REMARK"  Type="String" />
                <asp:Parameter Name="VATM_INYM"  Type="String" />
                <asp:Parameter Name="VATM_A_USER_ID" Type="String" />
                <asp:Parameter Name="VATM_A_USER_NM" Type="String" />
                <asp:Parameter Name="VATM_A_DATE"  Type="String" />
                <asp:Parameter Name="VATM_U_USER_ID" Type="String" />
                <asp:Parameter Name="VATM_U_USER_NM" Type="String" />
                <asp:Parameter Name="VATM_U_DATE"  Type="String" />
            </InsertParameters>
            
            <UpdateParameters>
                <asp:Parameter Name="VATM_EMPLID"  Type="String" />
                <asp:Parameter Name="VATM_SEQ"     Type="Int32"  />
                <asp:Parameter Name="VATM_DATE_ST" Type="String" />
                <asp:Parameter Name="VATM_DATE_EN" Type="String" />
                <asp:Parameter Name="VATM_TIME_ST" Type="String" />
                <asp:Parameter Name="VATM_TIME_EN" Type="String" />
                <asp:Parameter Name="VATM_HOURS"   Type="Double" />
                <asp:Parameter Name="VATM_REMARK"  Type="String" />
                <asp:Parameter Name="VATM_U_USER_ID" Type="String" />
                <asp:Parameter Name="VATM_U_USER_NM" Type="String" />
                <asp:Parameter Name="VATM_U_DATE"  Type="String" />
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
          <asp:FormView ID="fmvEdit" runat="server" DataKeyNames="VATM_EMPLID" DataSourceID="dscMain" >
            <ItemTemplate>
               <ucCHG:ATT_M_052U ID="ATT_M_052U" runat="server"/>
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
                  <ucCHG:ATT_M_052U ID="ATT_M_052U" runat="server"/>
                  <table class="table">
                    <tr>
                       <td align="center">
                          <asp:Button ID="btnMod" runat="server" Text="資料送出" CommandName="Update" 
                        /><asp:Button ID="btnCancel" runat="server"  Text="取消設定" CommandName="Cancel" />                  
                       </td>
                    </tr>
                  </table>
               </EditItemTemplate>  
            
            <InsertItemTemplate>
               <ucCHG:ATT_M_052U ID="ATT_M_052U" runat="server"/>
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
