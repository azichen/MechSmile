<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Contract.aspx.vb" Inherits="MMS_Contract" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
    </head>

<body onkeydown="if(event.keyCode==13)event.keyCode=9">
   <form id="form1" runat="server">      
      <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="true" EnableScriptLocalization="true">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
      <asp:SqlDataSource ID="dscList" runat="server" 
          ConnectionString="<%$ ConnectionStrings:Smile_HQConnectionString %>" 
          SelectCommand="SELECT MMSContract.* FROM MMSContract"></asp:SqlDataSource>
                              <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                                  
                                  
                                  SelectCommand="SELECT AreaCode, AreaCode+'_'+AreaName  as AreaName FROM MMSArea WHERE (Effective = 'Y')">
                              </asp:SqlDataSource>
                              <asp:SqlDataSource ID="SqlDataSource3" runat="server" 
                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" SelectCommand="SELECT EmployeeNo, EmployeeNo+'-'+EmployeeName  as EmployeeName  FROM MMSEmployee
where Effective = 'Y'"></asp:SqlDataSource>
                              <asp:SqlDataSource ID="SqlDataSource4" runat="server" 
                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" SelectCommand="SELECT EmployeeNo, EmployeeNo+'-'+EmployeeName  as EmployeeName  FROM MMSEmployee
where Effective = 'Y'
"></asp:SqlDataSource>
      <div style="width:10px; height:10px; position:absolute; left:-200px; top:-10px;" >
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
         <uc:QryStn ID="QryStn" runat="server" UpdatePanelObj="updQry" StnIDObj="txtContractStnID" StnNameObj="txtContractStnName" />
      </div>
      <uc:UsrCalendar ID="UsrCalendar" runat="server" />           
        
      <!-- 查詢輸入區畫面 -->
      <asp:UpdatePanel ID="updQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td align="center" style="height: 12px" colspan="2">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="合約基本資料維護-查詢"></asp:Label></td>
               </tr>
               <tr>
                  <td style="height:5px; width:15%"></td>
                  <td style="height:5px; width:85%"></td>
               </tr>
               <tr>
                  <td align="right" style="height:25px;">客戶代號：</td>
                  <td><asp:TextBox ID="txtCmpID" runat="server" MaxLength="15" Width="110px" AutoPostBack="true"
                    />&nbsp;至 
                      <asp:TextBox ID="txtCmpID0" runat="server" AutoPostBack="true" MaxLength="15" 
                          Width="110px" />
                   </td>
               </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        客戶名稱：</td>
                    <td>
                        <asp:TextBox ID="txtContractID1" runat="server" MaxLength="150" Width="257px" />
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        聯絡人：</td>
                    <td>
                        <asp:TextBox ID="TextBox1" runat="server" Width="109px"></asp:TextBox>
                        &nbsp;&nbsp;&nbsp; 連絡電話：<asp:TextBox ID="TextBox2" runat="server" Width="109px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        <asp:CheckBox ID="CheckBox2" runat="server" AutoPostBack="True" />
                        區域：</td>
                    <td>
                        <asp:DropDownList ID="DropDownList5" runat="server" AutoPostBack="True" 
                            DataSourceID="SqlDataSource2" DataTextField="AreaName" 
                            DataValueField="AreaCode" Enabled="False" Width="121px">
                        </asp:DropDownList>
                        &nbsp;
                        <asp:CheckBox ID="CheckBox3" runat="server" AutoPostBack="True" />
                        業務員：<asp:DropDownList ID="DropDownList6" runat="server" 
                            DataSourceID="SqlDataSource3" DataTextField="EmployeeName" 
                            DataValueField="EmployeeNo" Enabled="False" Width="121px">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        <asp:CheckBox ID="CheckBox4" runat="server" AutoPostBack="True" />
                        收款員：</td>
                    <td>
                        <asp:DropDownList ID="DropDownList7" runat="server" 
                            DataSourceID="SqlDataSource4" DataTextField="EmployeeName" 
                            DataValueField="EmployeeNo" Enabled="False" Width="121px">
                        </asp:DropDownList>
                    </td>
                </tr>
               <tr>
                  <td align="right" style="height:25px;">合約編號：</td>
                  <td><asp:TextBox ID="txtContractID" runat="server" MaxLength="14" Width="110px" 
                          AutoPostBack="True" />&nbsp;至 
                      <asp:TextBox ID="txtContractID0" runat="server" MaxLength="14" Width="110px" />
                   </td>
               </tr>
               <tr>
                  <td align="right" style="height:25px;">合約日期訖：</td>
                  <td><asp:TextBox ID="txtDtFrom" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false,'txtDtTo')" />
                      至&nbsp;<asp:TextBox ID="txtDtTo" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtTo',true,false)" />
                  </td>
               </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        大樓(案場)名稱：</td>
                    <td>
                        <asp:TextBox ID="txtContractID2" runat="server" MaxLength="150" Width="257px" />
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        &nbsp;</td>
                    <td>
                        <asp:CheckBox ID="CheckBox1" runat="server" Text="只顯示未歸檔合約" />
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        &nbsp;</td>
                    <td>
                        <asp:Button ID="btnQry" runat="server" Text="查詢" />
                        <asp:Button ID="btnClear" runat="server" Text="清除重來" />
                        <asp:Button ID="btnExcel" runat="server" Enabled="false" Text="EXCEL匯出" />
                        <img alt="" style=" border:0" src="../Images/teacher_icon.gif"
                   />
                        <asp:LinkButton ID="btnNew" runat="server" style="font-size:medium" 
                            Text="新增一筆資料" />
                        <asp:TextBox ID="txtSql" runat="server" Visible="False" Width="272px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" colspan="2" style="height:25px; text-align: left;">
                        <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right" colspan="2" style="height:25px; text-align: left;" 
                        valign="top">
                        <asp:GridView ID="grdList" runat="server" />
                    </td>
                </tr>
            </table>
         </ContentTemplate>
      </asp:UpdatePanel>

      <asp:UpdatePanel ID="updGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td style="height:25px; width:15%">&nbsp</td>
                  <td style="width:85%">
                      &nbsp;</td>
               </tr>
               <tr>
                  <td colspan="2" align="left">
                      &nbsp;</td>
               </tr>
            </table>
         </ContentTemplate>
         <Triggers>
            <asp:AsyncPostBackTrigger ControlID="btnQry" EventName="Click" />
         </Triggers>
      </asp:UpdatePanel>
   </form>
</body>
</html>
