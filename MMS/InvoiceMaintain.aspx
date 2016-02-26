<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InvoiceMaintain.aspx.vb" Inherits="MMS_InvoiceMaintain" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
    <style type="text/css">
        .style1
        {
            width: 507px;
        }
    </style>
    </head>

<body onkeydown="if(event.keyCode==13)event.keyCode=9">
   <form id="form1" runat="server">      
      <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="true" EnableScriptLocalization="true">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
      
      <uc:ReportCredentials ID="ReportCredentials" runat="server" />
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
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="發票處理作業" 
                          ForeColor="White"></asp:Label></td>
               </tr>
               <tr>
                  <td style="height:5px; width:15%"></td>
                  <td style="height:5px; width:85%"></td>
               </tr>
               <tr>
                  <td align="right" style="height:25px;">發票日期：</td>
                  <td><asp:TextBox ID="txtDtFrom" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false,'txtDtTo')" />
                      至&nbsp;<asp:TextBox ID="txtDtTo" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtTo',true,false)" />
                  </td>
               </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        申請日期：</td>
                    <td>
                        <asp:TextBox ID="txtDtFrom1" runat="server" Width="88px"></asp:TextBox>
                        <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                          onclick="ShowCalendar('txtDtFrom1',true,false,'txtDtTo1')" />
                        至&nbsp;<asp:TextBox ID="txtDtTo1" runat="server" Width="88px"></asp:TextBox>
                        <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                          onclick="ShowCalendar('txtDtTo1',true,false)" />
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        發票號碼：</td>
                    <td>
                        <asp:TextBox ID="txtDtFrom0" runat="server" Width="88px"></asp:TextBox>
                        &nbsp;&nbsp;&nbsp; 至&nbsp;<asp:TextBox ID="txtDtTo0" runat="server" Width="88px"></asp:TextBox>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        發票狀態：</td>
                    <td>
                        <asp:RadioButtonList ID="RadioButtonList1" runat="server" 
                            RepeatDirection="Horizontal" AutoPostBack="True">
                            <asp:ListItem Selected="True">已審核未開立</asp:ListItem>
                            <asp:ListItem>已開立</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        &nbsp;</td>
                    <td>
                         <asp:RadioButtonList ID="RadioButtonList2" runat="server" visible="false"
                            RepeatDirection="Horizontal" AutoPostBack="True" >
                            <asp:ListItem Selected="True">全部</asp:ListItem>
                            <asp:ListItem>有效</asp:ListItem>
                            <asp:ListItem>作廢</asp:ListItem>
                            <asp:ListItem>申請作廢</asp:ListItem>
                            <asp:ListItem>作廢核准</asp:ListItem>                            
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        &nbsp;</td>
                    <td>
                        <asp:Button ID="btnQry" runat="server" style="height: 21px" Text="查詢" 
                            Height="21px" />
                        <asp:Button ID="ClearButton" runat="server" Text="清除重來" Height="21px" />
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
                   </td>
               </tr>
                <tr>
                    <td align="left">
                        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                            AutoGenerateColumns="False" DataKeyNames="sn" 
                            EnableModelValidation="True" Width="800px">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("checkf") %>' 
                                             />
                                    </ItemTemplate>
                                    <ItemStyle Width="20px" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="sn" Visible="False">
                                    <ItemTemplate>
                                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("sn") %>'></asp:Label>
                                    </ItemTemplate>
                                    <EditItemTemplate>
                                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("sn") %>'></asp:TextBox>
                                    </EditItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="InvoiceTypeM" HeaderText="發票類別">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField DataField="InvoicePeriod" HeaderText="期別">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="發票號碼" DataField="InvoiceNo" >
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="客戶代號" DataField="CustomerNo" >
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="發票日期" DataField="InvoiceDate1" >
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="統一編號" DataField="GUINumber">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="抬頭" DataField="TitleOfInvoice">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="發票金額" DataField="Amount1">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="狀態" DataField="StatusN">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="" DataField="Status">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:CommandField ShowSelectButton="True" >
                                <HeaderStyle ForeColor="White" />
                                </asp:CommandField>
                                <asp:BoundField DataField="snM" Visible="False" />
                            </Columns>
                            <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                        </asp:GridView>
                        <br />
                        發票內容：<br /> 
                        <br />
                    </td>
                </tr>
                <tr>
                    <td align="left">
                        <asp:Button ID="Button1" runat="server" Text="全選" VISIBLE="False" 
                            Height="25px" />
                        <asp:Button ID="Button2" runat="server" Text="全不選" VISIBLE="False" 
                            Height="25px" />
                        <asp:Button ID="Button3" runat="server" Text="作廢發票" VISIBLE="False" 
                            Height="25px" />
                        <asp:Button ID="Button4" runat="server" Text="列印發票" 
                            VISIBLE="False" Height="25px" />
                        <asp:Button ID="Button6" runat="server" VISIBLE="False" Height="25px" 
                            Text="[電子]發票給號" />
                        <asp:Button ID="Button5" runat="server" VISIBLE="False" 
                            Text="取消審核" Height="25px" />
                        <asp:Button ID="Button7" runat="server" VISIBLE="FALSE" 
                            Text="核准作廢" Height="25px" OnClick="Button7_Click" />
                        <asp:Button ID="Button8" runat="server" VISIBLE="False" 
                            Text="取消作廢申請" Height="25px" OnClick="Button8_Click" />        
                    </td>
                </tr>
                <tr>
                    <td align="left">
                        <table style="width:100%;">
                            <tr>
                                <td class="style1" valign="top">
                                    <asp:GridView ID="GridView3" runat="server" AllowPaging="True" 
                                        AutoGenerateColumns="False" EnableModelValidation="True" 
                                        PageSize="4" Width="620px">
                                        <Columns>
                                            <asp:BoundField DataField="ItemName" HeaderText="品名" ReadOnly="True">
                                            <HeaderStyle ForeColor="White" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="UnitPrice" HeaderText="單價" ReadOnly="True">
                                            <HeaderStyle ForeColor="White" />
                                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Quantity" HeaderText="數量" ReadOnly="True">
                                            <HeaderStyle ForeColor="White" />
                                            <ItemStyle HorizontalAlign="Right" Width="40px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Amount" HeaderText="金額" ReadOnly="True">
                                            <HeaderStyle ForeColor="White" />
                                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Memo" HeaderText="備註" ReadOnly="True">
                                            <HeaderStyle ForeColor="White" />
                                            </asp:BoundField>
                                        </Columns>
                                        <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                                    </asp:GridView>
                                </td>
                                <td valign="top">
                                    <asp:GridView ID="GridView4" runat="server" AllowPaging="True" 
                                        AutoGenerateColumns="False" EnableModelValidation="True" 
                                        PageSize="4" Width="169px">
                                        <Columns>
                                            <asp:BoundField DataField="PrintDateTime" HeaderText="列印日期" ReadOnly="True">
                                            <HeaderStyle ForeColor="White" />
                                            <ItemStyle HorizontalAlign="Left" Width="100px" />
                                            </asp:BoundField>
                                        </Columns>
                                        <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                                    </asp:GridView>
                                </td>
                            </tr>
                        </table>
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
