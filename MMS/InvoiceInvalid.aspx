<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InvoiceInvalid.aspx.vb" Inherits="MMS_InvoiceInvalid" %>

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
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="發票作廢申請" 
                          ForeColor="White"></asp:Label></td>
               </tr>
               <tr>
                  <td style="height:5px; width:15%"></td>
                  <td style="height:5px; width:85%"></td>
               </tr>
               <tr>
                  <td align="right" style="height:25px;">申請日期：</td>
                  <td>
                      <asp:Label ID="Label2" runat="server"></asp:Label>
                  </td>
               </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        申請單位：</td>
                    <td>
                        <asp:DropDownList ID="DropDownList1" runat="server" Width="121px" 
                            AutoPostBack="True">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        申請人：</td>
                    <td>
                        <asp:DropDownList ID="DropDownList2" runat="server" Width="121px">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        發票號碼：</td>
                    <td>
                        <asp:TextBox ID="TextBox1" runat="server" AutoPostBack="True" Width="157px" 
                            MaxLength="10"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        發票期別：</td>
                    <td>
                        <asp:Label ID="Label6" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        發票日期：</td>
                    <td>
                        <asp:Label ID="Label5" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        發票金額：</td>
                    <td>
                        <asp:Label ID="Label4" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        統編：</td>
                    <td>
                        <asp:Label ID="Label7" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        發票抬頭：</td>
                    <td>
                        <asp:Label ID="Label3" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;" valign="top">
                        發票內容：</td>
                    <td>
                        <asp:GridView ID="GridView3" runat="server" AllowPaging="True" 
                            AutoGenerateColumns="False" EnableModelValidation="True" PageSize="4" 
                            Width="620px">
                            <Columns>
                                <asp:BoundField DataField="ItemName" HeaderText="品名" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                <asp:BoundField DataField="UnitPrice1" HeaderText="單價" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle HorizontalAlign="Right" Width="100px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Quantity" HeaderText="數量" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Amount1" HeaderText="金額" ReadOnly="True">
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
                </tr>
                <tr>
                    <td align="right" style="height: 25px;" valign="top">
                        作廢原因：</td>
                    <td>
                        <asp:TextBox ID="TextBox2" runat="server" Height="62px" TextMode="MultiLine" 
                            Width="569px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        &nbsp;</td>
                    <td>
                        <asp:Button ID="btnQry0" runat="server" style="height: 21px" Text="清除重來" />
                        <asp:Button ID="btnQry" runat="server" style="height: 21px" Text="資料送出" />
                        <asp:Button ID="btnQry1" runat="server" style="height: 21px" Text="回主畫面" />
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
                        &nbsp;</td>
                </tr>
                <tr>
                    <td align="left">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td align="left">
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
