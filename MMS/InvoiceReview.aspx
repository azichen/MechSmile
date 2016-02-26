<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InvoiceReview.aspx.vb" Inherits="MMS_InvoiceReview" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
    <style type="text/css">
        .style19
        {
            width: 67%;
        }
        .style20
        {
            width: 118px;
        }
        .style22
        {
            width: 195px;
        }
        .style23
        {
            width: 134px;
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
                  <td align="center" style="height: 12px; background-color: #1c5e55;" colspan="3">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="發票開立審核" 
                          ForeColor="White"></asp:Label></td>
               </tr>
            </table>
             <table  width="800px" class="style19">
                 <tr>
                     <td align="right" class="style20">
                         申請日期：</td>
                     <td class="style23">
                         <asp:TextBox ID="txtDtFrom" runat="server" Width="88px"></asp:TextBox>
                         <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false,'txtDtTo')" />
                     </td>
                     <td class="style22">
                         至 &nbsp;<asp:TextBox ID="txtDtTo" runat="server" Width="88px"></asp:TextBox>
                         <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtTo',true,false)" />
                     </td>
                     <td>
                         &nbsp;</td>
                 </tr>
                 <tr>
                     <td align="right" class="style20">
                         發票日期：</td>
                     <td class="style23">
                         <asp:TextBox ID="txtDtFrom0" runat="server" Width="88px"></asp:TextBox>
                         <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                          onclick="ShowCalendar('txtDtFrom0',true,false,'txtDtTo0')" />
                     </td>
                     <td class="style22">
                         至 &nbsp;<asp:TextBox ID="txtDtTo0" runat="server" Width="88px"></asp:TextBox>
                         <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                          onclick="ShowCalendar('txtDtTo0',true,false)" />
                         &nbsp;&nbsp;</td>
                     <td>
                         &nbsp;</td>
                 </tr>
                 <tr>
                     <td align="right" class="style20">
                         客戶代號：</td>
                     <td class="style23">
                         <asp:TextBox ID="CustFrom" runat="server" Width="88px" AutoPostBack="True"></asp:TextBox>
                     </td>
                     <td class="style22">
                         至 &nbsp;<asp:TextBox ID="CustTo" runat="server" Width="88px"></asp:TextBox>
                     </td>
                     <td>
                         <asp:Button ID="btnQry" runat="server" style="height: 21px" Text="查詢" />
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
                            EnableModelValidation="True" Width="787px">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("checkf") %>' 
                                            AutoPostBack="True" oncheckedchanged="CheckBox2_CheckedChanged" 
                                             />
                                    </ItemTemplate>
                                    <ItemStyle Width="20px" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="sn" Visible="False" />
                                <asp:TemplateField HeaderText="申請單號">
                                    <EditItemTemplate>
                                        <asp:Label ID="Label3" runat="server" Text='<%# Eval("DocNo") %>'></asp:Label>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("DocNo") %>'></asp:Label>
                                        <br />
                                        <asp:Label ID="Label6" runat="server" Text='<%# Bind("InvoiceType") %>'></asp:Label>
                                    </ItemTemplate>
                                    <HeaderStyle ForeColor="White" />
                                    <ItemStyle Width="80px" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="客戶代號">
                                    <HeaderTemplate>
                                        客戶代號<br />
                                    </HeaderTemplate>
                                    <ItemTemplate>
                                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("CustomerNo") %>'></asp:Label>
                                        <br />
                                    </ItemTemplate>
                                    <EditItemTemplate>
                                        <asp:Label ID="Label1" runat="server" Text='<%# Eval("CustomerNo") %>'></asp:Label>
                                    </EditItemTemplate>
                                    <HeaderStyle ForeColor="White" />
                                    <ItemStyle Width="90px" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="統一編號">
                                    <HeaderTemplate>
                                        統一編號<br /> 抬&nbsp;&nbsp;&nbsp; 頭
                                    </HeaderTemplate>
                                    <ItemTemplate>
                                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("GUINumber") %>'></asp:Label>
                                        <br />
                                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("TitleOfInvoice") %>'></asp:Label>
                                    </ItemTemplate>
                                    <EditItemTemplate>
                                        <asp:Label ID="Label2" runat="server" Text='<%# Eval("GUINumber") %>'></asp:Label>
                                    </EditItemTemplate>
                                    <HeaderStyle ForeColor="White" />
                                    <ItemStyle Width="220px" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="InvoiceDate" HeaderText="發票日期">
                                <ControlStyle Width="70px" />
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="70px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="InvoiceNo" HeaderText="發票號碼">
                                <ControlStyle Width="90px" />
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="80px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="amt" HeaderText="發票金額">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:CommandField ShowEditButton="True" ButtonType="Button" />
                                <asp:CommandField SelectText="瀏覽內容" ShowSelectButton="True" />
                            </Columns>
                            <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                        </asp:GridView>
                        <br />
                        發票內容：<br /> 
                        <asp:GridView ID="GridView2" runat="server" AllowPaging="True" 
                            AutoGenerateColumns="False" DataKeyNames="sn" EnableModelValidation="True" 
                            PageSize="5" Width="702px">
                            <Columns>
                                <asp:BoundField DataField="sn" Visible="False" />
                                <asp:BoundField DataField="Serial" HeaderText="Serial" Visible="False" />
                                <asp:BoundField DataField="ItemName" HeaderText="品名" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField DataField="UnitPrice" HeaderText="單價" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Quantity" HeaderText="數量" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Amount" HeaderText="金額" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Memo" HeaderText="備註">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                            </Columns>
                            <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                        </asp:GridView>
                        <br />
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Button ID="Button1" runat="server" Text="    全選    " Enabled="False" />
                        <asp:Button ID="Button2" runat="server" Text="  全不選  " Enabled="False" />
                        <asp:Button ID="Button3" runat="server" Text="資料送出" Enabled="False" />
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
