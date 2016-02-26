<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Cust_001.aspx.vb" Inherits="MMS_Cust_001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
    <title>未命名頁面</title>
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .style2
        {
        }
        .style4
        {
        }
        .style7
        {
        }
        .style14
        {
            width: 125px;
        }
        .style15
        {
            width: 141px;
        }
        .style16
        {
            width: 149px;
        }
        .style17
        {
            width: 82px;
        }
    </style>
</head>
<body  onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form id="form1" runat="server">    
      <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True">
         <Scripts> <asp:ScriptReference Path="~/Script/Util.js" /> </Scripts>
      </asp:ScriptManager>
      <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;">
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
      </div>
    <div>
    
        <table style="width: 800px;">
            <tr>
                    <td valign="middle" style="height: 25px; text-align: center;background-color:#1c5e55" align="center">
                        <asp:Label ID="TitleLabel" runat="server" CssClass="titles2" Text="功能名稱"></asp:Label></td>
            </tr>
            <tr>
                <td>
                    <asp:FormView ID="FormView1" runat="server" DataKeyNames="CustomerNo" 
                        DataSourceID="SqlDataSource1" EnableModelValidation="True" Width="100%">
                        <EditItemTemplate>
                            &nbsp;<table cellpadding="5" style="width: 669px;">
                                <tr>
                                    <td align="right" class="style7">
                                        <asp:Label ID="Label1" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        客戶代號:</td>
                                    <td class="style16">
                                        <asp:Label ID="CustomerNoLabel1" runat="server" 
                                            Text='<%# Eval("CustomerNo") %>' />
                                    </td>
                                    <td align="right" class="style17">
                                        <asp:Label ID="Label6" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        區域:</td>
                                    <td>
                                        <asp:DropDownList ID="DropDownList3" runat="server" Width="148px" 
                                            DataSourceID="SqlDataSource2" DataTextField="AreaName" 
                                            DataValueField="AreaCode"  
                                            AutoPostBack="True" 
                                            onselectedindexchanged="DropDownList3_SelectedIndexChanged1" >
                                        </asp:DropDownList>
                                        <asp:TextBox ID="AreaCodeTextBox" runat="server" Text='<%# Bind("AreaCode") %>' 
                                            Visible="False" Width="30px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        <asp:Label ID="Label2" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        客戶名稱:</td>
                                    <td class="style2" colspan="3">
                                        <asp:TextBox ID="CustomerNameTextBox" runat="server" 
                                            Text='<%# Bind("CustomerName") %>' Width="445px" MaxLength="75" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        客戶地址:</td>
                                    <td class="style2" colspan="3">
                                        <asp:TextBox ID="AddressTextBox" runat="server" Text='<%# Bind("Address") %>' 
                                            Width="445px" MaxLength="100" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        e-mail</td>
                                    <td class="style2" colspan="3">
                                        <asp:TextBox ID="EMAILTextBox" runat="server" Text='<%# Bind("EMAIL") %>' 
                                            Width="445px" MaxLength="50" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        聯絡電話:</td>
                                    <td class="style16">
                                        <asp:TextBox ID="TelNoTextBox" runat="server" Text='<%# Bind("TelNo") %>' 
                                            MaxLength="20" />
                                    </td>
                                    <td align="right" class="style17">
                                        傳真號碼:</td>
                                    <td>
                                        <asp:TextBox ID="FaxNoTextBox" runat="server" Text='<%# Bind("FaxNo") %>' 
                                            MaxLength="20" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        聯絡人:</td>
                                    <td class="style16">
                                        <asp:TextBox ID="ContactTextBox" runat="server" Text='<%# Bind("Contact") %>' 
                                            MaxLength="25" />
                                    </td>
                                    <td align="right" class="style17">
                                        <asp:Label ID="Label7" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        業務員:</td>
                                    <td>
                                        <asp:DropDownList ID="DropDownList1" runat="server" Width="148px" 
                                            DataSourceID="SqlDataSource3" DataTextField="EmployeeName" 
                                            DataValueField="EmployeeNo" 
                                            Enabled="True" 
                                            onselectedindexchanged="DropDownList1_SelectedIndexChanged" >
                                        </asp:DropDownList>
                                        <asp:TextBox ID="SalesmanTextBox" runat="server" Text='<%# Bind("Salesman") %>' 
                                            Width="30px" Visible="False" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        統一編號:</td>
                                    <td class="style16">
                                        <asp:TextBox ID="GUINumberTextBox" runat="server" 
                                            Text='<%# Bind("GUINumber") %>' MaxLength="8" 
                                            onkeypress="return CheckKeyNumber()" AutoPostBack="True" 
                                            ontextchanged="GUINumberTextBox_TextChanged" />
                                    </td>
                                    <td align="right" class="style17">
                                        <asp:Label ID="Label8" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        收款員:</td>
                                    <td>
                                        <asp:DropDownList ID="DropDownList2" runat="server" Width="148px" 
                                            DataSourceID="SqlDataSource4" DataTextField="EmployeeName" 
                                            DataValueField="EmployeeNo" 
                                            onselectedindexchanged="DropDownList2_SelectedIndexChanged" >
                                        </asp:DropDownList>
                                        <asp:TextBox ID="CashierTextBox" runat="server" Text='<%# Bind("Cashier") %>' 
                                            Width="30px" Visible="False" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        <asp:Label ID="Label3" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        有統編催收天數:</td>
                                    <td class="style16">
                                        <asp:TextBox ID="Date1TextBox" runat="server" Text='<%# Bind("Date1") %>' />
                                    </td>
                                    <td align="right" class="style17">
                                        <asp:Label ID="Label5" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        無統編催收天數:</td>
                                    <td>
                                        <asp:TextBox ID="Date2TextBox" runat="server" Text='<%# Bind("Date2") %>' />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        <asp:Label ID="Label14" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        收款天數:</td>
                                    <td class="style4" colspan="3">
                                        <asp:TextBox ID="DaysAllowedTextBox" runat="server" 
                                            Text='<%# Bind("DaysAllowed") %>' Width="445px" MaxLength="75" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        <asp:Label ID="Label4" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        發票抬頭:</td>
                                    <td class="style4" colspan="3">
                                        <asp:TextBox ID="TitleOfInvoiceTextBox" runat="server" 
                                            Text='<%# Bind("TitleOfInvoice") %>' Width="445px" MaxLength="75" />
                                        <asp:TextBox ID="EffectiveTextBox" runat="server" 
                                            Text='<%# Bind("Effective") %>' Visible="False" Width="10px" />
                                        <asp:TextBox ID="TextBox1" runat="server" Width="30px" 
                                            Text='<%# Bind("Used") %>' Visible="False"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style7">
                                        備註:</td>
                                    <td class="style4" colspan="3">
                                        <asp:TextBox ID="MemoTextBox" runat="server" Height="54px" 
                                            Text='<%# Bind("Memo") %>' TextMode="MultiLine" Width="445px" 
                                            MaxLength="100" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" class="style7" colspan="4">
                                        <asp:Button ID="Button6" runat="server"  Text="    修改    " 
                                            onclick="Button6_Click" />
                                        <asp:Button ID="Button5" runat="server" onclick="Button5_Click" 
                                            Text="    刪除    " />
                                        <asp:Button ID="Button3" runat="server" onclick="Button3_Click" Text="資料送出" />
                                        <asp:Button ID="Button4" runat="server" onclick="Button4_Click" Text="取消設定" />
                                    </td>
                                </tr>
                            </table>
                            <br />
                            <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" 
                                CommandName="Update" Text="更新" Visible="False" />
                            &nbsp;<asp:LinkButton ID="UpdateCancelButton" runat="server" 
                                CausesValidation="False" CommandName="Cancel" Text="取消" Visible="False" />
                        </EditItemTemplate>
                        <InsertItemTemplate>
                            <table cellpadding="5" style="width: 650px;">
                                <tr>
                                    <td align="right" class="style14">
                                        <asp:Label ID="Label6" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        客戶代號:</td>
                                    <td class="style15">
                                        <asp:TextBox ID="CustomerNoTextBox" runat="server" 
                                            Text='<%# Bind("CustomerNo") %>' AutoPostBack="True" 
                                            ontextchanged="CustomerNoTextBox_TextChanged"  
                                            onkeypress="return ToUpper()" MaxLength="15" />
                                    </td>
                                    <td align="right" class="style14">
                                        <asp:Label ID="Label12" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        區域:</td>
                                    <td>
                                        <asp:DropDownList ID="DropDownList3" runat="server" 
                                            DataSourceID="SqlDataSource2" DataTextField="AreaName" 
                                            DataValueField="AreaCode" 
                                            Width="148px" AutoPostBack="True"  
                                            onselectedindexchanged="DropDownList3_SelectedIndexChanged">
                                        </asp:DropDownList>
                                        <asp:TextBox ID="AreaCodeTextBox" runat="server" Text='<%# Bind("AreaCode") %>' 
                                            Visible="False" Width="10px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        <asp:Label ID="Label7" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        客戶名稱:</td>
                                    <td class="style2" colspan="3">
                                        <asp:TextBox ID="CustomerNameTextBox" runat="server" 
                                            Text='<%# Bind("CustomerName") %>' Width="445px" MaxLength="75" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        客戶地址:</td>
                                    <td class="style2" colspan="3">
                                        <asp:TextBox ID="AddressTextBox" runat="server" Text='<%# Bind("Address") %>' 
                                            Width="445px" MaxLength="100" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        e-mail</td>
                                    <td class="style2" colspan="3">
                                        <asp:TextBox ID="EMAILTextBox" runat="server" Text='<%# Bind("EMAIL") %>' 
                                            Width="445px" MaxLength="50" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        聯絡電話:</td>
                                    <td class="style15">
                                        <asp:TextBox ID="TelNoTextBox" runat="server" Text='<%# Bind("TelNo") %>' 
                                            MaxLength="20" />
                                    </td>
                                    <td align="right" class="style14">
                                        傳真號碼:</td>
                                    <td>
                                        <asp:TextBox ID="FaxNoTextBox" runat="server" Text='<%# Bind("FaxNo") %>' 
                                            MaxLength="20" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        聯絡人:</td>
                                    <td class="style15">
                                        <asp:TextBox ID="ContactTextBox" runat="server" Text='<%# Bind("Contact") %>' 
                                            MaxLength="25" />
                                    </td>
                                    <td align="right" class="style14">
                                        <asp:Label ID="Label10" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        業務員:</td>
                                    <td>
                                        <asp:DropDownList ID="DropDownList1" runat="server" 
                                            DataSourceID="SqlDataSource3" DataTextField="EmployeeName" 
                                            DataValueField="EmployeeNo" 
                                            Width="148px">
                                        </asp:DropDownList>
                                        <asp:TextBox ID="SalesmanTextBox" runat="server" Text='<%# Bind("Salesman") %>' 
                                            Visible="False" Width="10px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        統一編號:</td>
                                    <td class="style15">
                                        <asp:TextBox ID="GUINumberTextBox" runat="server" 
                                            Text='<%# Bind("GUINumber") %>' MaxLength="8" 
                                            onkeypress="return CheckKeyNumber()"  />
                                    </td>
                                    <td align="right" class="style14">
                                        <asp:Label ID="Label11" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        收款員:</td>
                                    <td>
                                        <asp:DropDownList ID="DropDownList2" runat="server" 
                                            DataSourceID="SqlDataSource4" DataTextField="EmployeeName" 
                                            DataValueField="EmployeeNo"  
                                            Width="148px">
                                        </asp:DropDownList>
                                        <asp:TextBox ID="CashierTextBox" runat="server" Text='<%# Bind("Cashier") %>' 
                                            Visible="False" Width="10px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        <asp:Label ID="Label8" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        有統編催收天數:</td>
                                    <td class="style15">
                                        <asp:TextBox ID="Date1TextBox" runat="server" Text='<%# Bind("Date1") %>' />
                                    </td>
                                    <td align="right" class="style14">
                                        <asp:Label ID="Label9" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        無統編催收天數:</td>
                                    <td>
                                        <asp:TextBox ID="Date2TextBox" runat="server" Text='<%# Bind("Date2") %>' />
                                    </td>
                                </tr>
                                 <tr>
                                    <td align="right" class="style7">
                                        <asp:Label ID="Label14" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        收款天數:</td>
                                    <td class="style4" colspan="3">
                                        <asp:TextBox ID="DaysAllowedTextBox" runat="server" 
                                            Text='<%# Bind("DaysAllowed") %>' Width="445px" MaxLength="75" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        <asp:Label ID="Label13" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                        發票抬頭:</td>
                                    <td class="style4" colspan="3">
                                        <asp:TextBox ID="TitleOfInvoiceTextBox" runat="server" 
                                            Text='<%# Bind("TitleOfInvoice") %>' Width="445px" MaxLength="75" />
                                        <asp:TextBox ID="EffectiveTextBox" runat="server" 
                                            Text='<%# Bind("Effective") %>' Visible="False" Width="10px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style14">
                                        備註:</td>
                                    <td class="style4" colspan="3">
                                        <asp:TextBox ID="MemoTextBox" runat="server" Height="39px" 
                                            Text='<%# Bind("Memo") %>' TextMode="MultiLine" Width="445px" 
                                            MaxLength="100" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" class="style7" colspan="4">
                                        <asp:Button ID="Button1" runat="server" onclick="Button1_Click" Text="資料送出" />
                                        <asp:Button ID="Button2" runat="server" onclick="Button2_Click" Text="回主畫面" />
                                    </td>
                                </tr>
                            </table>
                            <br />
                            <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" 
                                CommandName="Insert" Text="插入" Visible="False" />
                            &nbsp;<asp:LinkButton ID="InsertCancelButton" runat="server" 
                                CausesValidation="False" CommandName="Cancel" Text="取消" Visible="False" />
                        </InsertItemTemplate>
                        <ItemTemplate>
                            CustomerNo:
                            <asp:Label ID="CustomerNoLabel" runat="server" 
                                Text='<%# Eval("CustomerNo") %>' />
                            <br />
                            CustomerName:
                            <asp:Label ID="CustomerNameLabel" runat="server" 
                                Text='<%# Bind("CustomerName") %>' />
                            <br />
                            Address:
                            <asp:Label ID="AddressLabel" runat="server" Text='<%# Bind("Address") %>' />
                            <br />
                            AreaCode:
                            <asp:Label ID="AreaCodeLabel" runat="server" Text='<%# Bind("AreaCode") %>' />
                            <br />
                            EMAIL:
                            <asp:Label ID="EMAILLabel" runat="server" Text='<%# Bind("EMAIL") %>' />
                            <br />
                            TelNo:
                            <asp:Label ID="TelNoLabel" runat="server" Text='<%# Bind("TelNo") %>' />
                            <br />
                            Contact:
                            <asp:Label ID="ContactLabel" runat="server" Text='<%# Bind("Contact") %>' />
                            <br />
                            FaxNo:
                            <asp:Label ID="FaxNoLabel" runat="server" Text='<%# Bind("FaxNo") %>' />
                            <br />
                            GUINumber:
                            <asp:Label ID="GUINumberLabel" runat="server" Text='<%# Bind("GUINumber") %>' />
                            <br />
                            TitleOfInvoice:
                            <asp:Label ID="TitleOfInvoiceLabel" runat="server" 
                                Text='<%# Bind("TitleOfInvoice") %>' />
                            <br />
                            Salesman:
                            <asp:Label ID="SalesmanLabel" runat="server" Text='<%# Bind("Salesman") %>' />
                            <br />
                            Cashier:
                            <asp:Label ID="CashierLabel" runat="server" Text='<%# Bind("Cashier") %>' />
                            <br />
                            Memo:
                            <asp:Label ID="MemoLabel" runat="server" Text='<%# Bind("Memo") %>' />
                            <br />
                            Date1:
                            <asp:Label ID="Date1Label" runat="server" Text='<%# Bind("Date1") %>' />
                            <br />
                            Date2:
                            <asp:Label ID="Date2Label" runat="server" Text='<%# Bind("Date2") %>' />
                            <br />
                            DaysAllowed:
                            <asp:Label ID="DaysAllowedLabel" runat="server" Text='<%# Bind("DaysAllowed") %>' />
                            <br />
                            Effective:
                            <asp:Label ID="EffectiveLabel" runat="server" Text='<%# Bind("Effective") %>' />
                            <br />
                        </ItemTemplate>
                    </asp:FormView>
                </td>
            </tr>
            <tr>
                <td>
                        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>"
                            SelectCommand="SELECT * FROM [MMSCustomers]">
                        </asp:SqlDataSource>
   
                        <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                            
                            
                            SelectCommand="SELECT AreaCode, AreaCode + '-' + AreaName AS AreaName FROM MMSArea where [Effective] = 'Y'">
                        </asp:SqlDataSource>
                        <asp:SqlDataSource ID="SqlDataSource3" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                            
                            SelectCommand="SELECT EmployeeNo, EmployeeNo+'-'+EmployeeName as EmployeeName FROM MMSEmployee">
                        </asp:SqlDataSource>
                        <asp:SqlDataSource ID="SqlDataSource4" runat="server" 
                            ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                            SelectCommand="SELECT EmployeeNo, EmployeeNo+'-'+EmployeeName as EmployeeName FROM MMSEmployee">
                        </asp:SqlDataSource>
   
                    </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
