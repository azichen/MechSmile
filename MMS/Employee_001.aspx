<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Employee_001.aspx.vb" Inherits="MMS_Employee_001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .style1
        {
        }
        .style3
        {
            width: 187px;
        }
        .style5
        {
        }
        .style6
        {
            width: 65px;
        }
        
        .style79
        {
            color: #FF3300;
        }
        </style>
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
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
                    <asp:FormView ID="FormView1" runat="server" DataKeyNames="EmployeeNo" 
                        DataSourceID="SqlDataSource1" EnableModelValidation="True" Width="100%">
                        <EditItemTemplate>
                            &nbsp;<table cellpadding="5" style="width: 613px;">
                                <tr>
                                    <td class="style6">
                                        員工代號:</td>
                                    <td class="style3">
                                        <asp:Label ID="EmployeeNoLabel1" runat="server" 
                                            Text='<%# Eval("EmployeeNo") %>' />
                                    </td>
                                    <td class="style6">
                                        員工姓名:</td>
                                    <td>
                                        <asp:TextBox ID="EmployeeNameTextBox" runat="server" 
                                            Text='<%# Bind("EmployeeName") %>' MaxLength="5" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style1" colspan="2">
                                        <span class="style79">*</span><asp:CheckBox ID="CheckBox1" runat="server" Text="業務員" />
                                        &nbsp;&nbsp;<asp:CheckBox ID="CheckBox2" runat="server" Text="收款員" />
                                        &nbsp;
                                        <asp:CheckBox ID="CheckBox3" runat="server" Text="會計人員" />
                                    </td>
                                    <td class="style6">
                                        分機:</td>
                                    <td>
                                        <asp:TextBox ID="ExtNoTextBox" runat="server" Text='<%# Bind("ExtNo") %>' 
                                            onkeypress="return CheckKeyNumber()" MaxLength="10" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style6">
                                        會計代號:</td>
                                    <td class="style3">
                                        <asp:TextBox ID="AccountsTextBox" runat="server" 
                                            Text='<%# Bind("Accounts") %>' MaxLength="10" />
                                    </td>
                                    <td class="style6">
                                        <span class="style79">*</span>區域:</td>
                                    <td>
                                        <asp:DropDownList ID="DropDownList1" runat="server" 
                                            DataSourceID="SqlDataSource2" DataTextField="AreaName" 
                                            DataValueField="AreaCode"  Width="121px">
                                        </asp:DropDownList>
                                        <asp:TextBox ID="AreaCodeTextBox" runat="server" Text='<%# Bind("AreaCode") %>' 
                                            Width="40px" Visible="False" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style6">電子郵件:</td>
                                    <td class="style3" colspan="3">
                                        <asp:TextBox ID="EmailAdrTextBox" runat="server" 
                                            Text='<%# Bind("EmailAdr") %>' MaxLength="30" />
                                    </td>
                                   
                                </tr>
                                <tr>
                                    <td class="style6">
                                        <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("Used") %>' 
                                            Visible="False" Width="20px" />
                                    </td>
                                    <td class="style3">
                                        <asp:TextBox ID="EffectiveTextBox" runat="server" 
                                            Text='<%# Bind("Effective") %>' Visible="False" />
                                    </td>
                                    <td class="style6">
                                        &nbsp;</td>
                                    <td>
                                        <asp:TextBox ID="TextBox2" runat="server" Width="20px" 
                                            Text='<%# Bind("Salesman") %>' Visible="False"></asp:TextBox>
                                        <asp:TextBox ID="TextBox3" runat="server" Width="20px" 
                                            Text='<%# Bind("Cashier") %>' Visible="False"></asp:TextBox>
                                        <asp:TextBox ID="TextBox4" runat="server" Width="20px" 
                                            Text='<%# Bind("accountant") %>' Visible="False"></asp:TextBox>
                                    </td>
                                </tr>
                                 
                                <tr>
                                    <td align="center" class="style5" colspan="4">
                                        <asp:Button ID="Button6" runat="server" 
                                            Text="    修改    " onclick="Button6_Click" />
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
                            &nbsp;<table cellpadding="5" style="width: 613px;">
                                <tr>
                                    <td class="style5">
                                        <span class="style79">*</span>員工代號:</td>
                                    <td class="style3">
                                        <asp:TextBox ID="EmployeeNoTextBox" runat="server" 
                                            Text='<%# Bind("EmployeeNo") %>' Width="130px"  
                                            onkeypress="return ToUpper()" MaxLength="10"/>
                                    </td>
                                    <td align="right" class="style6">
                                        員工姓名:</td>
                                    <td>
                                        <asp:TextBox ID="EmployeeNameTextBox" runat="server" 
                                            Text='<%# Bind("EmployeeName") %>' Width="130px" MaxLength="10" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style1" colspan="2">
                                        <span class="style79">*</span><asp:CheckBox ID="CheckBox1" runat="server" Text="業務員" />
                                        &nbsp;&nbsp;<asp:CheckBox ID="CheckBox2" runat="server" Text="收款員" />
                                        &nbsp;
                                        <asp:CheckBox ID="CheckBox3" runat="server" Text="會計人員" />
                                    </td>
                                    <td align="right" class="style6">
                                        分機:</td>
                                    <td>
                                        <asp:TextBox ID="ExtNoTextBox" runat="server" Text='<%# Bind("ExtNo") %>' 
                                           onkeypress="return CheckKeyNumber()"  Width="130px" MaxLength="10" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style5">
                                        會計代號:</td>
                                    <td class="style3">
                                        <asp:TextBox ID="AccountsTextBox" runat="server" Text='<%# Bind("Accounts") %>' 
                                            Width="130px"  onkeypress="return CheckKeyNumber()" MaxLength="10"/>
                                    </td>
                                    <td align="right" class="style6">
                                        <span class="style79">*</span>區域:</td>
                                    <td>
                                        <asp:DropDownList ID="DropDownList1" runat="server" 
                                            DataSourceID="SqlDataSource2" DataTextField="AreaName" 
                                            DataValueField="AreaCode" SelectedValue='<%# Bind("AreaCode") %>' Width="135px">
                                        </asp:DropDownList>
                                        <asp:TextBox ID="AreaCodeTextBox" runat="server" Text='<%# Bind("AreaCode") %>' 
                                            Visible="False" Width="30px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style6">
                                        電子郵件:</td>
                                    <td class="style3" colspan="3">
                                        <asp:TextBox ID="EmailAdrTextBox" runat="server" 
                                            Text='<%# Bind("EmailAdr") %>' MaxLength="30" />
                                    </td>
                                   
                                </tr>
                                <tr>
                                    <td align="center" class="style5" colspan="4">
                                        <asp:Button ID="Button1" runat="server" onclick="Button1_Click" Text="資料送出" />
                                        <asp:Button ID="Button2" runat="server" onclick="Button2_Click" Text="回主畫面" />
                                        <asp:TextBox ID="EffectiveTextBox" runat="server" 
                                            Text='<%# Bind("Effective") %>' Visible="False" Width="130px" />
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
                            EmployeeNo:
                            <asp:Label ID="EmployeeNoLabel" runat="server" 
                                Text='<%# Eval("EmployeeNo") %>' />
                            <br />
                            EmployeeName:
                            <asp:Label ID="EmployeeNameLabel" runat="server" 
                                Text='<%# Bind("EmployeeName") %>' />
                            <br />
                            Accounts:
                            <asp:Label ID="AccountsLabel" runat="server" Text='<%# Bind("Accounts") %>' />
                            <br />
                            AreaCode:
                            <asp:Label ID="AreaCodeLabel" runat="server" Text='<%# Bind("AreaCode") %>' />
                            <br />
                            ExtNo:
                            <asp:Label ID="ExtNoLabel" runat="server" Text='<%# Bind("ExtNo") %>' />
                            <br />
                            EmployeeType:
                            <asp:Label ID="EmployeeTypeLabel" runat="server" 
                                Text='<%# Bind("EmployeeType") %>' />
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
                            SelectCommand="SELECT * FROM [MMSEmployee]">
                        </asp:SqlDataSource>
   
                        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>"
                            SelectCommand="SELECT AreaCode,AreaCode+'-'+AreaName as AreaName FROM [MMSArea] ">
                            <SelectParameters>
                                <asp:Parameter DefaultValue="Y" Name="Effective" Type="String" />
                            </SelectParameters>
                        </asp:SqlDataSource>
   
                    </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
