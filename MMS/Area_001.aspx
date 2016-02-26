<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Area_001.aspx.vb" Inherits="MMS_Area_001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>未命名頁面</title>
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .style1
        {
        }
        .style2
        {
        }
        .style3
        {
            width: 81px;
        }
        .style4
        {
            width: 167px;
        }
        .style5
        {
        }
        .style6
        {
            width: 71px;
        }
    
        .style79
        {
            color: #FF3300;
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
                    <asp:FormView ID="FormView1" runat="server" DataKeyNames="AreaCode" 
                        DataSourceID="SqlDataSource1" EnableModelValidation="True" Width="100%">
                        <EditItemTemplate>
                            <table cellpadding="5" cellspacing="2" style="width: 68%;">
                                <tr>
                                    <td align="right" class="style5">
                                        區域代號:</td>
                                    <td class="style4">
                                        <asp:Label ID="AreaCodeLabel1" runat="server" Text='<%# Eval("AreaCode") %>' />
                                    </td>
                                    <td align="right" class="style6">
                                        名稱:</td>
                                    <td>
                                        <asp:TextBox ID="AreaNameTextBox" runat="server" 
                                            Text='<%# Bind("AreaName") %>' />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style5">
                                        電話:</td>
                                    <td class="style4">
                                        <asp:TextBox ID="TelNoTextBox" runat="server" Text='<%# Bind("TelNo") %>' />
                                    </td>
                                    <td align="right" class="style6">
                                        傳真:</td>
                                    <td>
                                        <asp:TextBox ID="FaxNoTextBox" runat="server" Text='<%# Bind("FaxNo") %>' />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style5">
                                        地址:</td>
                                    <td class="style2" colspan="3">
                                        <asp:TextBox ID="AddressTextBox" runat="server" Text='<%# Bind("Address") %>' 
                                            Width="411px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style5">
                                        稅額單位:</td>
                                    <td class="style4">
                                        <asp:TextBox ID="TaxUnitTextBox" runat="server" Text='<%# Bind("TaxUnit") %>' onkeypress="return ToUpper()" />
                                    </td>
                                    <td align="right" class="style6">
                                        中心代號:</td>
                                    <td>
                                        <asp:TextBox ID="ProfitNoTextBox" runat="server" 
                                            Text='<%# Bind("ProfitNo") %>'  onkeypress="return ToUpper()" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style5">
                                        單位代號:</td>
                                    <td class="style4">
                                        <asp:TextBox ID="UnitCodeTextBox" runat="server" 
                                            Text='<%# Bind("UnitCode") %>'  onkeypress="return ToUpper()"/>
                                    </td>
                                    <td class="style6" align="right">
                                        <asp:TextBox ID="TextBox1" runat="server" Width="20px"  
                                            Text='<%# Bind("Used") %>' Visible="False" />
                                    </td>
                                    <td>
                                        <asp:TextBox ID="EffectiveTextBox" runat="server" 
                                            Text='<%# Bind("Effective") %>' Visible="False" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" class="style5" colspan="4">
                                        <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" 
                                            CommandName="Update" Text="更新" Visible="False" />
                                        &nbsp;
                                        <asp:LinkButton ID="UpdateCancelButton" runat="server" CausesValidation="False" 
                                            CommandName="Cancel" Text="取消" Visible="False" />
                                        <asp:Button ID="Button6" runat="server" 
                                            Text="    修改    " onclick="Button6_Click" />
                                        <asp:Button ID="Button5" runat="server" onclick="Button5_Click" 
                                            Text="    刪除    " />
                                        <asp:Button ID="Button3" runat="server" onclick="Button3_Click" Text="資料送出" />
                                        <asp:Button ID="Button4" runat="server" onclick="Button4_Click" Text="取消設定" />
                                    </td>
                                </tr>
                            </table>
                        </EditItemTemplate>
                        <InsertItemTemplate>
                            <table cellpadding="5" cellspacing="2" style="width: 68%;">
                                <tr>
                                    <td align="right" class="style1">
                                        <span class="style79">*</span>區域代號:</td>
                                    <td class="style4">
                                        <asp:TextBox ID="AreaCodeTextBox" runat="server" 
                                            Text='<%# Bind("AreaCode") %>' MaxLength="1" />
                                    </td>
                                    <td align="right" class="style3">
                                        名稱:</td>
                                    <td>
                                        <asp:TextBox ID="AreaNameTextBox" runat="server" 
                                            Text='<%# Bind("AreaName") %>' />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style1">
                                        電話:</td>
                                    <td class="style4">
                                        <asp:TextBox ID="TelNoTextBox" runat="server" Text='<%# Bind("TelNo") %>' />
                                    </td>
                                    <td align="right" class="style3">
                                        傳真:</td>
                                    <td>
                                        <asp:TextBox ID="FaxNoTextBox" runat="server" Text='<%# Bind("FaxNo") %>' />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style1">
                                        地址:</td>
                                    <td class="style2" colspan="3">
                                        <asp:TextBox ID="AddressTextBox" runat="server" Text='<%# Bind("Address") %>' 
                                            Width="411px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style1">
                                        稅額單位:</td>
                                    <td class="style4">
                                        <asp:TextBox ID="TaxUnitTextBox" runat="server" Text='<%# Bind("TaxUnit") %>'  onkeypress="return ToUpper()" />
                                    </td>
                                    <td align="right" class="style3">
                                        中心代號:</td>
                                    <td>
                                        <asp:TextBox ID="ProfitNoTextBox" runat="server" 
                                            Text='<%# Bind("ProfitNo") %>'  onkeypress="return ToUpper()"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" class="style1">
                                        單位代號:</td>
                                    <td class="style4">
                                        <asp:TextBox ID="UnitCodeTextBox" runat="server" 
                                            Text='<%# Bind("UnitCode") %>'  onkeypress="return ToUpper()"/>
                                    </td>
                                    <td class="style3" align="right">
                                        &nbsp;</td>
                                    <td>
                                        <asp:TextBox ID="EffectiveTextBox" runat="server" 
                                            Text='<%# Bind("Effective") %>' Visible="False" />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" class="style1" colspan="4">
                                        <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" 
                                            CommandName="Insert" Text="插入" Visible="False" />
                                        &nbsp;
                                        <asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" 
                                            CommandName="Cancel" Text="取消" Visible="False" />
                                        <asp:Button ID="Button1" runat="server" onclick="Button1_Click" Text="資料送出" />
                                        <asp:Button ID="Button2" runat="server" onclick="Button2_Click" Text="回主畫面" />
                                    </td>
                                </tr>
                            </table>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            AreaCode:
                            <asp:Label ID="AreaCodeLabel" runat="server" Text='<%# Eval("AreaCode") %>' />
                            <br />
                            AreaName:
                            <asp:Label ID="AreaNameLabel" runat="server" Text='<%# Bind("AreaName") %>' />
                            <br />
                            TelNo:
                            <asp:Label ID="TelNoLabel" runat="server" Text='<%# Bind("TelNo") %>' />
                            <br />
                            FaxNo:
                            <asp:Label ID="FaxNoLabel" runat="server" Text='<%# Bind("FaxNo") %>' />
                            <br />
                            Address:
                            <asp:Label ID="AddressLabel" runat="server" Text='<%# Bind("Address") %>' />
                            <br />
                            TaxUnit:
                            <asp:Label ID="TaxUnitLabel" runat="server" Text='<%# Bind("TaxUnit") %>' />
                            <br />
                            ProfitNo:
                            <asp:Label ID="ProfitNoLabel" runat="server" Text='<%# Bind("ProfitNo") %>' />
                            <br />
                            UnitCode:
                            <asp:Label ID="UnitCodeLabel" runat="server" Text='<%# Bind("UnitCode") %>' />
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
                            SelectCommand="SELECT * FROM [MMSArea]">
                        </asp:SqlDataSource>
   
                    </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
