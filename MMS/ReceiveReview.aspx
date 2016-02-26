<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ReceiveReview.aspx.vb" Inherits="MMS_ReceiveReview" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
    <style type="text/css">

    
        .style79
        {
            color: #FF3300;
        }
        </style>
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
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="收款資料審核"></asp:Label></td>
               </tr>
               <tr>
                  <td style="height:5px; width:15%"></td>
                  <td style="height:5px; width:85%"></td>
               </tr>
               <tr>
                  <td align="right" style="height:25px;">收款員：</td>
                  <td>
                      <asp:TextBox ID="txtSaleID1" runat="server" AutoPostBack="True" MaxLength="6" 
                          Width="110px" />
                      <asp:Label ID="Label1" runat="server"></asp:Label>
                   </td>
               </tr>
                <tr>
                    <td align="right" style="height:25px;">
                        客戶代號：</td>
                    <td>
                        <asp:TextBox ID="txtCmpID" runat="server" AutoPostBack="True" MaxLength="6" 
                            Width="110px" />
                        <asp:Label ID="Label2" runat="server"></asp:Label>
                    </td>
                </tr>
               <tr>
                  <td align="right" style="height:25px;">申請日期：</td>
                  <td><asp:TextBox ID="txtDtFrom" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false,'txtDtTo')" />
                      至&nbsp;<asp:TextBox ID="txtDtTo" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtTo',true,false)" />
                  </td>
               </tr>
               <tr>
                  <td align="right" style="height:25px;">核可日期：</td>
                  <td><asp:TextBox ID="txtDtFr2" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFr2',true,false,'txtDtTo2')" />
                      至&nbsp;<asp:TextBox ID="txtDtTo2" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtTo2',true,false)" />
                  </td>
               </tr>
               <tr>
                    <td align="right" style="height: 25px;">
                        審核狀態：</td>
                    <td>
                        <asp:RadioButtonList ID="RadioButtonList1" runat="server" 
                            RepeatDirection="Horizontal" AutoPostBack="True">
                            <asp:ListItem Selected="True">全部</asp:ListItem>
                            <asp:ListItem>已審核</asp:ListItem>
                            <asp:ListItem>未審核</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 25px;">
                        付款別：</td>
                    <td>
                        <asp:CheckBox ID="CheckBox1" runat="server" Text="現金" ></asp:CheckBox>
                        <asp:CheckBox ID="CheckBox3" runat="server" Text="匯款" ></asp:CheckBox>
                        <asp:CheckBox ID="CheckBox4" runat="server" Text="支票" ></asp:CheckBox>
                        <asp:CheckBox ID="CheckBox5" runat="server" Text="其他" ></asp:CheckBox>
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
                     <asp:Button ID="btnQry" runat="server" Text="查詢"
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
             <table class="table">
                 <tr>
                     <td align="left">
                         <asp:GridView ID="GridView4" runat="server" AllowPaging="True" 
                             AutoGenerateColumns="False" DataKeyNames="ReceivablesNo" 
                             EnableModelValidation="True" Width="800px">
                             <Columns>
                                 <asp:TemplateField>
                                     <ItemTemplate>
                                         <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("checkf") %>' />
                                     </ItemTemplate>
                                     <ItemStyle Width="20px" />
                                 </asp:TemplateField>
                                 <asp:BoundField DataField="ReceivablesNo" HeaderText="申請單號" ReadOnly="True">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="CustomerNo" HeaderText="客戶代碼">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="CustomerName" HeaderText="客戶名稱">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="CASH" HeaderText="現金"  ItemStyle-HorizontalAlign="Right">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="AMOUNTOFCHECK" HeaderText="應收票據"  ItemStyle-HorizontalAlign="Right">
                                 <HeaderStyle ForeColor="White" />
                                 <ItemStyle HorizontalAlign="Right" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="AmountOfRemittances" HeaderText="匯款金額"  ItemStyle-HorizontalAlign="Right">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="OtherAmt" HeaderText="其他金額"  ItemStyle-HorizontalAlign="Right">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="RemittanceDate" HeaderText="匯款日">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="ReviewDate" HeaderText="核可日期">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 <asp:BoundField DataField="ReviewerName" HeaderText="審核人">
                                 <HeaderStyle ForeColor="White" />
                                 </asp:BoundField>
                                 
                                 <asp:TemplateField ShowHeader="False">
                                     <ItemTemplate>
                                         <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" 
                                             CommandName="Select" Text="選取"></asp:LinkButton>
                                     </ItemTemplate>
                                 </asp:TemplateField>
                             </Columns>
                             <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                         </asp:GridView>
                     </td>
                 </tr>
                 <tr>
                     <td align="left">
                         <asp:Button ID="Button1" runat="server" Text="全選" Visible="False" />
                         <asp:Button ID="Button2" runat="server" Text="全不選" Visible="False" />
                         <asp:Button ID="Button3" runat="server" style="height: 21px" Text="審核" 
                             Visible="False" />
                          <asp:Button ID="Button4" runat="server" style="height: 21px" Text="取消審核" 
                             Visible="False" />    
                             &nbsp;&nbsp;&nbsp; 「審核」按鈕需為未審核項才會顯現，「取消審核」按鈕需為已審核項才會顯現!
                     </td>
                 </tr>
             </table>
             <br />
             <asp:Panel ID="Panel1" runat="server">
                 <table class="table">
                     <tr>
                         <td align="center" colspan="2" style="height: 12px; background-color: #1c5e55;">
                             <asp:Label ID="Label9" runat="server" Font-Bold="True" Font-Size="Large" 
                                 ForeColor="White">收款單資料</asp:Label>
                         </td>
                     </tr>
                     <tr>
                         <td class="style4">
                         </td>
                         <td style="height: 5px; width: 85%">
                             <asp:Label ID="Label10" runat="server" Text="Label" Visible="False"></asp:Label>
                         </td>
                     </tr>
                     <tr>
                         <td align="left" class="style5" colspan="2">
                             <span class="style79">*</span>收款日期：<asp:TextBox ID="txtDtFrom2" runat="server" 
                                 ReadOnly="True" Width="88px"></asp:TextBox>
                             &nbsp;&nbsp;
                             <asp:Label ID="Label11" runat="server" Text="建檔日期："></asp:Label>
                             <asp:Label ID="Label3" runat="server"></asp:Label>
                             &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                             <asp:Label ID="Label4" runat="server" Text="單據狀態："></asp:Label>
                             <asp:Label ID="Label5" runat="server"></asp:Label>
                         </td>
                     </tr>
                     <tr>
                         <td align="left" class="style5" colspan="2">
                             <span class="style79">*</span>客戶代號：<asp:TextBox ID="txtDtFrom0" runat="server" 
                                 AutoPostBack="True" ReadOnly="True" Width="88px"></asp:TextBox>
                             &nbsp;<asp:Label ID="Label6" runat="server"></asp:Label>
                             <table class="table">
                                 <tr>
                                     <td align="left">
                                         <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                                             AutoGenerateColumns="False" DataKeyNames="Serial" EnableModelValidation="True" 
                                             PageSize="3" Width="800px">
                                             <Columns>
                                                 <asp:BoundField DataField="Serial" ReadOnly="True">
                                                 <ItemStyle Width="30px" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="ChequeNo" HeaderText="支票號碼">
                                                 <HeaderStyle ForeColor="White" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="AmountOfCheck" HeaderText="支票金額">
                                                 <HeaderStyle ForeColor="White" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="PayingBank" HeaderText="付款銀行">
                                                 <HeaderStyle ForeColor="White" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="BankAccount" HeaderText="銀行帳號">
                                                 <HeaderStyle ForeColor="White" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="ChequeExpiryDate" HeaderText="支票到期日">
                                                 <HeaderStyle ForeColor="White" />
                                                 </asp:BoundField>
                                             </Columns>
                                             <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                                         </asp:GridView>
                                     </td>
                                 </tr>
                                 <tr>
                                     <td align="left">
                                         &nbsp;</td>
                                 </tr>
                                 <tr>
                                     <td align="left">
                                         現&nbsp;&nbsp;&nbsp; 金：<asp:TextBox ID="TextBox2" runat="server" ReadOnly="True"></asp:TextBox>
                                     </td>
                                 </tr>
                                 <tr>
                                     <td align="left">
                                         匯款金額：<asp:TextBox ID="TextBox3" runat="server" ReadOnly="True"></asp:TextBox>
                                         &nbsp; 匯款手續費：<asp:TextBox ID="TextBox4" runat="server" ReadOnly="True"></asp:TextBox>
                                         &nbsp; 匯款日期：<asp:TextBox ID="txtDtFrom1" runat="server" ReadOnly="True" Width="88px"></asp:TextBox>
                                         &nbsp;</td>
                                 </tr>
                                 <tr>
                                     <td align="left" class="style2">
                                         銷貨折讓：<asp:TextBox ID="TextBox5" runat="server" ReadOnly="True"></asp:TextBox>
                                         &nbsp;&nbsp;&nbsp; 銷項稅額：<asp:TextBox ID="TextBox6" runat="server" ReadOnly="True"></asp:TextBox>
                                         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 其他：<asp:TextBox ID="TextBox7" runat="server" ReadOnly="True"></asp:TextBox>
                                         &nbsp;</td>
                                 </tr>
                                 <tr>
                                     <td align="left">
                                         &nbsp;&nbsp;&nbsp;&nbsp; 合計：<asp:Label ID="Label7" runat="server"></asp:Label>
                                         &nbsp;&nbsp;&nbsp;&nbsp; 製單人：<asp:Label ID="Label8" runat="server"></asp:Label>
                                         &nbsp;&nbsp;&nbsp;&nbsp; 收款員：<asp:DropDownList ID="DropDownList1" runat="server" Enabled="False" 
                                             Width="131px">
                                         </asp:DropDownList>
                                     </td>
                                 </tr>
                                 <tr>
                                     <td align="left" valign="top">
                                         &nbsp;&nbsp;&nbsp;&nbsp; 備註：<asp:TextBox ID="TextBox8" runat="server" Height="44px" ReadOnly="True" 
                                             TextMode="MultiLine" Width="624px"></asp:TextBox>
                                     </td>
                                 </tr>
                                 <tr>
                                     <td align="left" class="style3" valign="bottom">
                                         沖銷發票發票：</td>
                                 </tr>
                                 <tr>
                                     <td align="left" valign="top">
                                         <asp:GridView ID="GridView3" runat="server" 
                                             AutoGenerateColumns="False" EnableModelValidation="True" 
                                             Width="800px">
                                             <Columns>
                                                 <asp:BoundField DataField="InvoiceDate" HeaderText="發票日期" ReadOnly="True">
                                                 <HeaderStyle ForeColor="White" />
                                                 <ItemStyle HorizontalAlign="Left" Width="100px" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="InvoiceNo" HeaderText="發票號碼" ReadOnly="True">
                                                 <HeaderStyle ForeColor="White" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="Amount" HeaderText="發票金額" ReadOnly="True">
                                                 <HeaderStyle ForeColor="White" />
                                                 <ItemStyle HorizontalAlign="Right" Width="100px" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="ReceiptAmount" HeaderText="已收款金額" ReadOnly="True">
                                                 <HeaderStyle ForeColor="White" />
                                                 <ItemStyle HorizontalAlign="Right" Width="100px" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="UnReceiptAmount" HeaderText="未收款金額" ReadOnly="True">
                                                 <HeaderStyle ForeColor="White" />
                                                 <ItemStyle HorizontalAlign="Right" Width="100px" />
                                                 </asp:BoundField>
                                                 <asp:BoundField DataField="revicea" HeaderText="沖銷金額">
                                                 <HeaderStyle ForeColor="White" />
                                                 <ItemStyle HorizontalAlign="Right" Width="100px" />
                                                 </asp:BoundField>
                                             </Columns>
                                             <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                                         </asp:GridView>
                                     </td>
                                 </tr>
                                 <tr>
                                     <td align="left">
                                         <table style="width: 100%;">
                                             <tr>
                                                 <td class="style1" valign="top">
                                                     &nbsp;&nbsp;</td>
                                                 <td valign="top">
                                                     &nbsp;</td>
                                             </tr>
                                         </table>
                                     </td>
                                 </tr>
                             </table>
                         </td>
                     </tr>
                 </table>
             </asp:Panel>
             <br />
         </ContentTemplate>
         <Triggers>
            <asp:AsyncPostBackTrigger ControlID="btnQry" EventName="Click" />
         </Triggers>
      </asp:UpdatePanel>
   </form>
</body>
</html>
