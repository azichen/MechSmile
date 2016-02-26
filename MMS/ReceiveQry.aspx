<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ReceiveQry.aspx.vb" Inherits="MMS_ReceiveQry" %>

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
      
      <asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:Smile_HQConnectionString %>" />
      
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
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="應收帳款查詢"></asp:Label></td>
               </tr>
               <tr>
                  <td style="height:5px; width:15%"></td>
                  <td style="height:5px; width:85%"></td>
               </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        客戶代號：</td>
                    <td>
                        <asp:TextBox ID="txtCmpID" runat="server" AutoPostBack="True" MaxLength="6" 
                            Width="110px" OnTextChanged="txtCmpID_TextChanged" />
                        &nbsp;至&nbsp;<asp:TextBox ID="txtCmpID0" runat="server" AutoPostBack="True" MaxLength="6" 
                            Width="110px" />
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        客戶名稱：</td>
                    <td>
                        <asp:TextBox ID="txtSaleID2" runat="server" AutoPostBack="True" MaxLength="6" 
                            Width="254px" />
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        發票號碼：</td>
                    <td>
                        <asp:TextBox ID="txtCmpID1" runat="server" AutoPostBack="True" 
                            Width="110px" />
                    </td>
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
                        沖銷狀態：</td>
                    <td>
                        <asp:RadioButtonList ID="RadioButtonList1" runat="server" 
                            RepeatDirection="Horizontal">
                            <asp:ListItem Selected="True">全部</asp:ListItem>
                            <asp:ListItem>未沖銷完畢</asp:ListItem>
                            <asp:ListItem>已沖銷完畢</asp:ListItem>
                        </asp:RadioButtonList>
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
                     <asp:Button ID="btnQry" runat="server" Text="查詢" style="height: 21px"
                   /><asp:Button ID="btnClear" runat="server" Text="清除重來" Visible="False"
                   /><asp:Button ID="btnExcel" runat="server" Text="EXCEL匯出" Enabled="false" />
                      &nbsp;
                      </td>
               </tr>
               <tr>
                  <td colspan="2" align="left">
                     <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label>
                      <asp:TextBox ID="txtSql" runat="server" Visible="False" Width="272px"></asp:TextBox>
                   </td>
               </tr>
            </table>
            
            <asp:GridView ID="grdList" runat="server" />
            
            <table class="table">
                 <tr>
                     <td align="left">
                         發票內容:</td>
                 </tr>
                 <tr>
                     <td align="left">
                         <asp:GridView ID="GridView5" runat="server" AutoGenerateColumns="False" 
                             EnableModelValidation="True" Width="800px">
                             <Columns>
                                 <asp:BoundField DataField="InvoiceNo" HeaderText="發票號碼" ReadOnly="True">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="Serial" HeaderText="項次"></asp:BoundField>
                                 <asp:BoundField DataField="ItemName" HeaderText="品名" ReadOnly="True">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="UnitPrice" HeaderText="單價" ReadOnly="True">
                                 <HeaderStyle ForeColor="White" />
                                 <ItemStyle HorizontalAlign="Right" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="Quantity" HeaderText="數量" ReadOnly="True">
                                 <ItemStyle HorizontalAlign="Right" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="Amount" HeaderText="金額">
                                 <HeaderStyle ForeColor="White" />

                                 <ItemStyle HorizontalAlign="Right" />

                                 </asp:BoundField>
                             </Columns>
                             <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                         </asp:GridView>
                     </td>
                 </tr>
                 <tr>
                     <td align="left">
                         收款紀錄:</td>
                 </tr>
                 <tr>
                     <td align="left">
                         <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" 
                             EnableModelValidation="True" Width="467px">
                             <Columns>
                                 <asp:BoundField DataField="收款日期" HeaderText="收款日期" ReadOnly="True">
                                 <HeaderStyle ForeColor="White" />
                                 <ItemStyle HorizontalAlign="Left" Width="100px" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="發票號碼" HeaderText="發票號碼" ReadOnly="True">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="沖銷金額" HeaderText="沖銷金額">
                                 <HeaderStyle ForeColor="White" />
                                 <ItemStyle HorizontalAlign="Right" Width="100px" />
                                 </asp:BoundField>
                             </Columns>
                             <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                         </asp:GridView>
                     </td>
                 </tr>
             </table>
             <br />
             <br />
         </ContentTemplate>
         <Triggers>
            <asp:AsyncPostBackTrigger ControlID="btnQry" EventName="Click" />
         </Triggers>
      </asp:UpdatePanel>
   </form>
</body>
</html>
