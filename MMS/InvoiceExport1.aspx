<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InvoiceExport1.aspx.vb" Inherits="MMS_InvoiceExport1" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
    <style type="text/css">


.modalPopup {
	background-color:#ffffdd;
	border-width:3px;
	border-style:solid;
	border-color:Blue;
	padding:3px;
	width:250px;
}

    .msg
{
    font-weight: bold;
    font-size: 12pt;
    color: #ff0066;
    font-variant: normal;
}

        .style1
        {
            height: 5px;
            width: 15%;
        }
        .style2
        {
            height: 25px;
            }

    </style>
    </head>

<body onkeydown="if(event.keyCode==13)event.keyCode=9">
   <form id="form1" runat="server">      
      <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="true" EnableScriptLocalization="true">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
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
                  <td align="center" style="height: 12px; background-color: #1c5e55;" colspan="2">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="發票資料轉出" 
                          ForeColor="White"></asp:Label></td>
               </tr>
               <tr>
                  <td class="style1"></td>
                  <td style="height:5px; width:85%">&nbsp;&nbsp; &nbsp;</td>
               </tr>
               <tr>
                  <td align="right" class="style2">申請日期：</td>
                  <td>
                      <asp:TextBox ID="txtDtFrom" runat="server" Width="88px"></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false,'txtDtTo')" />
                      至&nbsp;<asp:TextBox ID="txtDtTo" runat="server" Width="88px"></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtTo',true,false)" />
                  </td>
               </tr>
                <tr>
                    <td align="right" class="style2">
                        發票日期：</td>
                    <td>
                        <asp:TextBox ID="txtDtFrom0" runat="server" Width="88px"></asp:TextBox>
                        <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                            onclick="ShowCalendar('txtDtFrom0',true,false,'txtDtTo0')" />
                        至&nbsp;<asp:TextBox ID="txtDtTo0" runat="server" Width="88px"></asp:TextBox>
                        <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                            onclick="ShowCalendar('txtDtTo0',true,false)" />
                    </td>
                </tr>
                <tr>
                    <td align="right" class="style2">
                        客戶代號：</td>
                    <td>
                        <asp:TextBox ID="txtDtFrom1" runat="server" Width="88px" AutoPostBack="True"></asp:TextBox>
                        &nbsp;&nbsp;&nbsp; 至 
                        <asp:TextBox ID="txtDtFrom2" runat="server" Width="88px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" class="style2">
                        發票號碼：</td>
                    <td>
                        <asp:TextBox ID="txtDtFrom3" runat="server" Width="88px" AutoPostBack="True"></asp:TextBox>
                        &nbsp;&nbsp;&nbsp; 至&nbsp;<asp:TextBox ID="txtDtFrom4" runat="server" Width="88px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" class="style2">
                        <asp:CheckBox ID="CheckBox2" runat="server" AutoPostBack="True" Text="業務員" />
                        ：</td>
                    <td>
                        <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" 
                            Width="100px" Enabled="False">
                        </asp:DropDownList>
                        <asp:DropDownList ID="DropDownList2" runat="server" AutoPostBack="True" 
                            Enabled="False" Width="100px">
                        </asp:DropDownList>
                        &nbsp;&nbsp;&nbsp;
                        <asp:CheckBox ID="CheckBox1" runat="server" Text="顯示作廢發票" />
                    </td>
                </tr>
                <tr>
                    <td align="right" class="style2">
                        &nbsp;</td>
                    <td>
                        <asp:Button ID="Button1" runat="server" style="height: 21px" Text="查詢" />
                        <asp:Button ID="btnQry1" runat="server" Text="清  除" />
                        <asp:Button ID="ExcelButton" runat="server" Text="EXCEL匯出" Enabled="False" />
                        <asp:TextBox ID="txtSql" runat="server" Visible="False" Width="272px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="style2" colspan="2">
                        <asp:GridView ID="grdList" runat="server" />
                        <asp:SqlDataSource ID="dscList" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:Smile_HQConnectionString %>">
                        </asp:SqlDataSource>
                    </td>
                </tr>
                <tr>
                    <td align="left" style="height: 25px;" colspan="2">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td align="left" colspan="2" style="height: 25px;">
                        &nbsp;</td>
                </tr>
            </table>
         </ContentTemplate>
      </asp:UpdatePanel>

   </form>
</body>
</html>
