<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InvoiceApply.aspx.vb" Inherits="MMS_InvoiceApply" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
    <style type="text/css">
        .style1
        {
            height: 5px;
            width: 11%;
        }
        .style2
        {
            height: 25px;
            width: 11%;
        }
        .style3
        {
            height: 5px;
            width: 50px;
        }
        .style5
        {
            height: 25px;
            width: 50px;
        }
        .style6
        {
            height: 14px;
        }
    </style>
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
                  <td align="center" style="height: 12px" colspan="3">
                      <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="發票開立需求維護-查詢"></asp:Label>
                   </td>
               </tr>
               <tr>
                  <td class="style3">&nbsp;</td>
                   <td class="style1">
                   </td>
                  <td style="height:5px; width:85%"></td>
               </tr>
               <tr>
                  <td align="right" class="style5">&nbsp;</td>
                   <td align="right" class="style2">
                       業務員：</td>
                  <td>
                      <asp:TextBox ID="txtSaleID1" runat="server" AutoPostBack="True" MaxLength="6" 
                          Width="110px" />
                      <asp:Label ID="Label1" runat="server"></asp:Label>
                   </td>
               </tr>
                <tr>
                    <td align="right" class="style5">
                        &nbsp;</td>
                    <td align="right" class="style2">
                        客戶代號：</td>
                    <td>
                        <asp:TextBox ID="txtCmpID" runat="server" AutoPostBack="True" MaxLength="6" 
                            Width="110px" />
                        <asp:Label ID="Label2" runat="server"></asp:Label>
                    </td>
                </tr>
               <tr>
                  <td align="right" class="style5">&nbsp;</td>
                   <td align="right" class="style2">
                       合約編號：</td>
                  <td><asp:TextBox ID="txtContractID" runat="server" MaxLength="14" Width="110px" 
                          AutoPostBack="True" />&nbsp; 至 
                      <asp:TextBox ID="txtContractID0" runat="server" MaxLength="14" Width="110px" />
                   </td>
               </tr>
               <tr>
                  <td align="right" class="style5">&nbsp;</td>
                   <td align="right" class="style2">
                       申請日期：</td>
                  <td><asp:TextBox ID="txtDtFrom" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false,'txtDtTo')" />
                      至&nbsp;<asp:TextBox ID="txtDtTo" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtTo',true,false)" />
                  </td>
               </tr>
                <tr>
                    <td align="right" class="style5">
                        &nbsp;</td>
                    <td align="right" class="style2">
                        &nbsp;</td>
                    <td>
                        <asp:Button ID="btnQry" runat="server" Text="    查    詢    " />
                        <asp:Button ID="btnClear" runat="server" Text="清除重來" Visible="False" />
                        <asp:Button ID="btnExcel" runat="server" Enabled="false" Text="EXCEL匯出" />
                        <img alt="" style=" border:0" src="../Images/teacher_icon.gif"
                   />
                        <asp:LinkButton ID="btnNew" runat="server" style="font-size: medium" 
                            Text="一般申請" />
                        &nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:LinkButton ID="btnNew0" runat="server" style="font-size: medium" 
                            Text="常態申請" />
                    </td>
                </tr>
            </table>
         </ContentTemplate>
      </asp:UpdatePanel>

      <asp:UpdatePanel ID="updGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td align="left">
                     <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label>
                      <asp:TextBox ID="txtSql" runat="server" Visible="False" Width="272px"></asp:TextBox>
                   </td>
               </tr>
                <tr>
                    <td align="left">
                        <asp:GridView ID="grdList" runat="server" DataKeyNames="DocNo" />
                    </td>
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
