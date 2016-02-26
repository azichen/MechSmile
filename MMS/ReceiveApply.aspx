<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ReceiveApply.aspx.vb" Inherits="MMS_Contract_001" %>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <style type="text/css">

    
        .style79
        {
            color: #FF3300;
        }
        .style80
        {
            width: 99px;
        }
        .style81
        {
            width: 90px;
        }
        .style82
        {
            width: 120px;
        }
        .style83
        {
            width: 128px;
        }
        .style84
        {
            width: 141px;
        }
        </style>
    <script language="javascript" type="text/javascript">
// <![CDATA[

        function Button6_onclick() {

        }

// ]]>
    </script>
</head>

<body onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form id="form1" runat="server" >
          <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True">
                  <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
      <uc:ReportCredentials ID="ReportCredentials" runat="server" />
            <table class="table">
               <tr>
                  <td align="center" style="height: 12px; background-color: #1c5e55;" colspan="2">
                     <asp:Label ID="Label9" runat="server" 
                          ForeColor="White" Font-Bold="True" Font-Size="Large"></asp:Label></td>
               </tr>
               <tr>
                  <td class="style4"></td>
                  <td style="height:5px; width:85%">
                      <asp:Label ID="Label10" runat="server" Text="Label" Visible="False"></asp:Label>
                   </td>
               </tr>
               <tr>
                  <td align="left" class="style5" colspan="2">
                      <span class="style79">*</span>收款日期：<asp:TextBox ID="txtDtFrom" 
                          runat="server" Width="88px"></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                          onclick="ShowCalendar('txtDtFrom',true,false)" />
                      &nbsp;
                      <asp:Label ID="Label2" runat="server" Text="建檔日期："></asp:Label>
                      <asp:TextBox ID="TextBox18" runat="server" Width="97px"></asp:TextBox>
                                    &nbsp;&nbsp;&nbsp;&nbsp;
                      <asp:Label ID="Label4" runat="server" Text="單據狀態："></asp:Label>
                      <asp:TextBox ID="TextBox19" runat="server" Width="97px"></asp:TextBox>
                   </td>
               </tr>
                <tr>
                    <td align="left" class="style5" colspan="2">
                        <span class="style79">*</span>客戶代號：<asp:TextBox 
                            ID="txtDtFrom0" runat="server" AutoPostBack="True" 
                            Width="88px"></asp:TextBox>
                        <asp:TextBox ID="TextBox9" runat="server" Width="336px"></asp:TextBox>
                        <table class="table">
                            <tr>
                                <td align="left" colspan="2">
                                    <table style="width:100%;background-color:skyblue" cellpadding="0">
                                        <tr>
                                            <td class="style80" align="center">
                                                支票號碼</td>
                                            <td align="center" class="style81">
                                                支票金額</td>
                                            <td align="center" class="style82">
                                                付款銀行</td>
                                            <td align="center" class="style83">
                                                銀行帳號</td>
                                            <td align="center" class="style84">
                                                支票到期日</td>
                                            <td>
                                                &nbsp;</td>
                                        </tr>
                                        <tr>
                                            <td class="style80">
                                                <asp:TextBox ID="TextBox10" runat="server" Width="99px" TabIndex="1"></asp:TextBox>
                                            </td>
                                            <td class="style81">
                                                <asp:TextBox ID="TextBox11" runat="server" Width="93px" TabIndex="2"></asp:TextBox>
                                            </td>
                                            <td class="style82">
                                                <asp:TextBox ID="TextBox12" runat="server" Width="125px" TabIndex="3"></asp:TextBox>
                                            </td>
                                            <td class="style83">
                                                <asp:TextBox ID="TextBox13" runat="server" TabIndex="4"></asp:TextBox>
                                            </td>
                                            <td class="style84">
                                                <asp:TextBox ID="TextBox14" runat="server" Width="88px" TabIndex="5"></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                          onclick="ShowCalendar('TextBox14',true,false)" /></td>
                                            <td>
                                                <asp:TextBox ID="TextBox15" runat="server" Width="58px" Visible="False"></asp:TextBox>
                                                <asp:Button ID="Button6" runat="server" Text="確認" />
                                                <asp:Button ID="Button7" runat="server" Text="Button" Visible="False" />
                                                </td>
                                        </tr>
                                        <tr>
                                            <td class="style80">
                                                &nbsp;</td>
                                            <td class="style81">
                                                &nbsp;</td>
                                            <td class="style82">
                                                &nbsp;</td>
                                            <td class="style83">
                                                &nbsp;</td>
                                            <td class="style84">
                                                &nbsp;</td>
                                            <td>
                                                &nbsp;</td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" colspan="2">
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
                                            <asp:CommandField ButtonType="Button" SelectText="修改" ShowSelectButton="True" >
                                            <ItemStyle HorizontalAlign="Right" />
                                            </asp:CommandField>
                                            <asp:CommandField ShowDeleteButton="True" ButtonType="Button">
                                            <ItemStyle Width="40px" />
                                            </asp:CommandField>
                                        </Columns>
                                        <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                                    </asp:GridView>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" colspan="2">
                                    &nbsp;</td>
                            </tr>
                            <tr>
                                <td align="left" style="background-color: skyblue" colspan="2">
                                    現&nbsp;&nbsp;&nbsp; 金：<asp:TextBox ID="TextBox2" runat="server" text-align="right"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" style="background-color: skyblue" colspan="2">
                                    匯款金額：<asp:TextBox ID="TextBox3" runat="server"></asp:TextBox>
                                    &nbsp; 匯款手續費：<asp:TextBox ID="TextBox4" runat="server"></asp:TextBox>
                                    &nbsp; 匯款日期：<asp:TextBox ID="txtDtFrom1" runat="server" Width="100px"></asp:TextBox>
                                    <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                                        onclick="ShowCalendar('txtDtFrom1',true,false)" />
                                </td>
                            </tr>
                            <tr>
                                <td align="left" class="style2" style="background-color: skyblue" colspan="2">
                                    銷貨折讓：<asp:TextBox ID="TextBox5" runat="server"></asp:TextBox>
                                    &nbsp;&nbsp;&nbsp; 銷項稅額：<asp:TextBox ID="TextBox6" runat="server"></asp:TextBox>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 其他：<asp:TextBox ID="TextBox7" runat="server" 
                                        Width="120px"></asp:TextBox>
                                    &nbsp;</td>
                            </tr>
                            <tr>
                                <td align="left" style="background-color: skyblue" colspan="2">
                                    &nbsp;&nbsp;&nbsp; 合計：<asp:TextBox 
                                        ID="TextBox16" runat="server"></asp:TextBox>
                                    &nbsp;&nbsp; &nbsp;&nbsp; 製單人：<asp:TextBox ID="TextBox17" runat="server"></asp:TextBox>
                                    &nbsp;&nbsp;&nbsp; 收款員：<asp:DropDownList ID="DropDownList1" 
                                        runat="server" Width="131px">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" valign="top" style="background-color: skyblue" colspan="2">
                                    &nbsp;&nbsp;&nbsp; 備註：<asp:TextBox ID="TextBox8" runat="server" Height="44px" 
                                        TextMode="MultiLine" Width="624px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" class="style3" valign="bottom">
                                    <asp:Label ID="Label1" runat="server" Text="尚未收款發票："></asp:Label></td>
                                <td align="right" class="style3" valign="bottom">
                                    <asp:Button ID="Button10" runat="server" Text="分配沖帳發票" />
                                </td>
                            </tr>
                            <tr>
                                <td align="left" valign="top" colspan="2">
                                    <asp:GridView ID="GridView3" runat="server" AllowPaging="True" 
                                        AutoGenerateColumns="False" EnableModelValidation="True" PageSize="4" 
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
                                            <asp:CommandField ShowEditButton="True" />
                                        </Columns>
                                        <HeaderStyle BackColor="Teal" Font-Size="Small" ForeColor="White" />
                                    </asp:GridView>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" colspan="2">
                                    <table style="width:100%;">
                                        <tr>
                                            <td class="style1" valign="top" align="center">
                                                <asp:Button ID="Button9" runat="server" Text="    修改    " style="height: 21px" 
                                                    Visible="False" />
                                                <asp:Button ID="Button3" runat="server" Text="    刪除    " OnClick="Button3_Click" 
                                    OnClientClick="return confirm('確定要刪除此筆資料嗎?')" style="height: 21px" Visible="False" />    
                                                <asp:Button ID="Button1" runat="server" Text="資料送出" style="height: 21px" 
                                                    Visible="False" />
                                                <asp:Button ID="Button5" runat="server" Text="清除重來" 
                                                    style="height: 21px" />
                                                <asp:Button ID="Button4" runat="server" Text="資料送出" Visible="False" />
                                                <asp:Button ID="Button2" runat="server" Text="回主畫面" />
                                                <asp:Button ID="Button8" runat="server" Text="取消設定" />
                                                <asp:Button ID="Button11" runat="server" Text="核帳單" /></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <rsweb:ReportViewer ID="rptCmp" runat="server" Font-Names="Verdana" Font-Size="8pt" Visible=false
                                    Height="1200px" ProcessingMode="Remote" Width="1100px">
            </rsweb:ReportViewer>
          <br />
      <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;">
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
      </div>
      <uc:UsrCalendar ID="UsrCalendar" runat="server" /> 
    <div>
    
    </div>
    </form>
</body>
</html>
