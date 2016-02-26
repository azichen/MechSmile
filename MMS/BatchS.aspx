<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BatchS.aspx.vb" Inherits="MMS_BatchS" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

    <style type="text/css">

	 

.titles
{
    text-decoration: underline;
    font-weight: bold;
    font-size: 12pt;

}
            
.modalPopup {
	background-color:#ffffdd;
	border-width:3px;
	border-style:solid;
	border-color:Blue;
	padding:3px;
	width:250px;
}

        .style2
        {
        }
        .style4
        {
            height: 26px;
        }

        .style5
        {
            width: 50px;
        }
        .style6
        {
            height: 26px;
            width: 50px;
        }
        .style7
        {
            width: 56px;
        }
        .style8
        {
            height: 26px;
            width: 56px;
        }

    </style>
</head>
<body  onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form id="form1" runat="server">
    <div>
    
      <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="true" EnableScriptLocalization="true">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
        <table style="width: 800px;">
            <tr>
                <td align="center" class="style2" colspan="3">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" 
                        Text="業務員\收款員批次修改-查詢"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style5">
                    &nbsp;</td>
                <td class="style7">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style6" style="style4">
                </td>
                <td class="style8" style="font-size: small">
                    <asp:Label ID="Lable2" runat="server" Text="區        域: " Font-Size="Small"></asp:Label>
                </td>
                <td class="style4" style="font-size: small">
                    <asp:DropDownList 
                        ID="DropDownList3" runat="server" Width="150px" 
                        AutoPostBack="True">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style6" style="style4">
                    &nbsp;</td>
                <td class="style8" style="font-size: small">
                    &nbsp;</td>
                <td class="style4" style="font-size: small">
                    <asp:RadioButtonList ID="RadioButtonList1" runat="server" AutoPostBack="True" 
                        RepeatDirection="Horizontal" Font-Size="Small">
                        <asp:ListItem Selected="True" Value="1">業務員</asp:ListItem>
                        <asp:ListItem Value="2">收款員</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="style5">
                    &nbsp;</td>
                <td class="style7">
                    <asp:Label ID="Lable1" runat="server" Text="原業務員: " Font-Size="Small"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="DropDownList1" runat="server" Width="150px" 
                        AutoPostBack="True">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style5">
                    &nbsp;</td>
                <td class="style7">
                    <asp:Label ID="Label2" runat="server" Text="新業務員: " Font-Size="Small"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="DropDownList2" runat="server" Width="150px">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style5">
                    &nbsp;</td>
                <td class="style7">
                    &nbsp;</td>
                <td>
                    <asp:Button ID="Button5" runat="server" Text="查    詢" Width="78px" 
                        Height="22px" />
                    <asp:Button ID="Button2" runat="server" Text="全  選" Enabled="False" 
                        Width="78px" Height="22px" />
                    <asp:Button ID="Button3" runat="server" Text="全不選" 
                        Enabled="False" Width="78px" Height="22px" />
                    <asp:Button ID="Button1" runat="server" Text="批次更新" Enabled="False" 
                        Height="22px" />
                </td>
            </tr>
            <tr>
                <td align="center" class="style2" colspan="3">
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style2" colspan="3">
                                              <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                                                  AutoGenerateColumns="False" EnableModelValidation="True" PageSize="15" 
                                                  Width="368px" DataKeyNames="CustomerNo" Font-Size="Small">
                                                  <Columns>
                                                      <asp:TemplateField>
                                                          <ItemTemplate>
                                                              <asp:CheckBox ID="CheckBox2" runat="server"  Checked='<%# Bind("checkf") %>' 
                                                                  AutoPostBack="True" oncheckedchanged="CheckBox2_CheckedChanged" 
                                                                  />
                                                          </ItemTemplate>
                                                          <ItemStyle Width="20px" />
                                                      </asp:TemplateField>
                                                      <asp:BoundField DataField="CustomerNo" HeaderText="客戶編號">
                                                      <ItemStyle Width="100px" />
                                                      </asp:BoundField>
                                                      <asp:BoundField DataField="CustomerName" HeaderText="客戶名稱" />
                                                  </Columns>
                                                  <HeaderStyle BackColor="Teal" ForeColor="White" />
                                              </asp:GridView>
                                          <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                                                  SelectCommand="SELECT [EmployeeNo], [EmployeeNo]+'_'+[EmployeeName] as EmployeeName FROM [MMSEmployee]" >
                    </asp:SqlDataSource>
                                          <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
                                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                                                  SelectCommand="SELECT [EmployeeNo], [EmployeeNo]+'_'+[EmployeeName] as EmployeeName FROM [MMSEmployee]" >
                    </asp:SqlDataSource>
                                          <asp:SqlDataSource ID="SqlDataSource3" runat="server" 
                                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                                                  
                                                  SelectCommand="SELECT AreaCode, AreaCode+'-'+AreaName as  AreaName  FROM MMSArea where Effective ='Y'" >
                    </asp:SqlDataSource>
                </td>
            </tr>
            <tr>
                <td class="style5">
                    &nbsp;</td>
                <td class="style7">
                    &nbsp;</td>
                <td>
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="style5">
                    &nbsp;</td>
                <td class="style7">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
