<%@ Page Language="VB" AutoEventWireup="false" debug='true' CodeFile="InvoiceApply002.aspx.vb" Inherits="MMS_InvoiceApply002" %>

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
        .style80
        {}
        .style81
        {
            height: 17px;
            width: 255px;
        }
        .style82
        {
            height: 17px;
            width: 66px;
        }
        .style83
        {
            width: 386px;
        }
        .style84
        {
            width: 64px;
        }
        .style85
        {
            width: 64px;
            height: 17px;
        }
        </style>
</head>

<body onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form id="form1" runat="server">
          <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True">
                  <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
                  <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
       <uc:UsrCalendar ID="UsrCalendar" runat="server" /> 
          <br />
                          <table style="width: 800px;">
            <tr>
                <td align="center" 
                    style="height: 25px; text-align: center; background-color: #1c5e55">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" 
                        Text="發票開立一般申請" ForeColor="White"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style4">
                    </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Lable3" runat="server" Text="申請日期: " Font-Size="Small"></asp:Label>
                    &nbsp;<asp:Label ID="Label1" runat="server" Text="Label" Font-Size="Small"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Lable2" runat="server" Text="申請單位: " Font-Size="Small"></asp:Label>
                    <asp:DropDownList 
                        ID="DropDownList3" runat="server" Width="150px" 
                        AutoPostBack="True">
                    </asp:DropDownList>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Label ID="Lable1" runat="server" Text="業務員: " Font-Size="Small" 
                        Visible="False"></asp:Label>
                    <asp:DropDownList ID="DropDownList1" runat="server" Width="150px" 
                        AutoPostBack="True" Visible="False">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;&nbsp;
                    <table style="width:100%;">
                        <tr>
                            <td class="style81">
                    <asp:Label ID="Lable6" runat="server" Text="發票日期: " Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="txtDtTo" runat="server" Width="140px" ></asp:TextBox>
                      <asp:Image ID="Image2" runat="server" ImageUrl="../Images/date.gif"  
                        onclick="ShowCalendar('txtDtTo',true,false)" /></td>
                            <td class="style82">
                                發票類別:</td>
                            <td class="style80">
                                <asp:RadioButtonList ID="RadioButtonList1" runat="server" 
                                    RepeatDirection="Horizontal">
                                    <asp:ListItem Selected="True">收銀發票</asp:ListItem>
                                    <asp:ListItem >手開發票</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Lable4" runat="server" Text="客戶代碼: " Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="TextBox1" runat="server" Width="140px" AutoPostBack="True" 
                        style="height: 19px"></asp:TextBox>
                    <asp:Label ID="Label2" runat="server" Text="客戶名稱: " Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="TextBox2" runat="server" ReadOnly="True" Width="450px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style3">
                    <asp:Label ID="Lable5" runat="server" Text="統一編號: " Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="TextBox3" runat="server" Width="140px"></asp:TextBox>
                    <asp:Label ID="Label3" runat="server" Text="發票抬頭: " Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="TextBox4" runat="server" Width="450px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style3" valign="top">
                    <table style="width:100%;">
                        <tr>
                            <td class="style85">
                    <asp:Label ID="Lable7" runat="server" Text="備" Font-Size="Small"></asp:Label>
                    &nbsp;&nbsp;
                    <asp:Label ID="Lable8" runat="server" Text="註: " Font-Size="Small" Visible="False"></asp:Label>
                            </td>
                            <td class="style80" rowspan="3">
                    <asp:TextBox ID="TextBox11" runat="server" Width="692px" AutoPostBack="True" Height="84px" 
                                    TextMode="MultiLine"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="style84">
                                &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style84">
                                &nbsp;</td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Panel ID="Panel3" runat="server">
                        <table align="left" cellpadding="0" cellspacing="0" 
    style="width: 696px;">
                            <tr>
                                <td align="center" class="style83" style="background-color: skyblue">
                                    <span class="style79">*</span>品名</td>
                                <td align="center" class="style6" style="background-color: skyblue">
                                    <span class="style79">*</span>單價</td>
                                <td align="center" class="style7" style="background-color: skyblue">
                                    <span class="style79">*</span>數量</td>
                                <td align="center" style="background-color: skyblue">
                                    (含稅)金額</td>
                            </tr>
                            <tr>
                                <td align="left" class="style83" style="background-color: skyblue">
                                    <asp:TextBox ID="TextBox5" runat="server" Width="439px"></asp:TextBox>
                                </td>
                                <td align="left" class="style6" style="background-color: skyblue">
                                    <asp:TextBox ID="TextBox6" runat="server" AutoPostBack="True" Width="100px"></asp:TextBox>
                                </td>
                                <td align="left" class="style7" style="background-color: skyblue">
                                    <asp:TextBox ID="TextBox7" runat="server" AutoPostBack="True" Width="100px"></asp:TextBox>
                                </td>
                                <td align="left" style="background-color: skyblue">
                                    <asp:TextBox ID="TextBox8" runat="server" Width="100px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style83" style="background-color: skyblue">
                                    <asp:TextBox ID="TextBox10" runat="server" Visible="False" Width="123px"></asp:TextBox>
                                </td>
                                <td class="style6" style="background-color: skyblue">
                                    &nbsp;</td>
                                <td class="style7" style="background-color: skyblue">
                                    &nbsp;</td>
                                <td style="border-width: 0px; background-color: skyblue">
                                    <asp:Button ID="Button7" runat="server" Text="清除" Height="21px" />
                                    <asp:Button ID="Button5" runat="server" Text="確認" Height="21px" />
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Panel ID="Panel2" runat="server">
                    </asp:Panel>
                </td>
            </tr>
            <tr>
                <td>
                        <asp:GridView ID="GridView2" 
    runat="server" 
                                                  AutoGenerateColumns="False" EnableModelValidation="True" 
                                                  Width="781px" PageSize="5" 
                         DataKeyNames="Sn" Font-Size="Small">
                            <Columns>
                                <asp:BoundField DataField="Sn" ReadOnly="True">
                                <ItemStyle Width="20px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="ItemName" HeaderText="品名" >
                                <ControlStyle Width="240px" />
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="200px" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="單價" DataField="UnitPrice" >
                                <ControlStyle Width="70px" />
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="70px" HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="數量" DataField="Qty">
                                <ControlStyle Width="70px" />
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="70px" HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="金額" DataField="Amount" ReadOnly="True" >
                                <ControlStyle Width="100px" />
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle HorizontalAlign="Right" Width="100px" />
                                </asp:BoundField>
                                <asp:TemplateField ShowHeader="False">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" 
                                            CommandName="Delete" Text="刪除"></asp:LinkButton>
                                    </ItemTemplate>
                                    <EditItemTemplate>
                                        <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" 
                                            CommandName="Update" Text="更新"></asp:LinkButton>
                                        &nbsp;<asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" 
                                            CommandName="Cancel" Text="取消"></asp:LinkButton>
                                    </EditItemTemplate>
                                    <ItemStyle Width="60px" />
                                </asp:TemplateField>
                                <asp:CommandField ShowSelectButton="True" >
                                <ItemStyle Width="100px" />
                                </asp:CommandField>
                            </Columns>
                            <HeaderStyle BackColor="Teal" ForeColor="White" Font-Size="Small" />
                        </asp:GridView>
                    </td>
            </tr>
            <tr>
                <td>
                    <asp:Panel ID="Panel1" runat="server">
                        <asp:GridView ID="GridView1" 
    runat="server" AllowPaging="True" 
                                                  AutoGenerateColumns="False" EnableModelValidation="True" 
                                                  Width="781px" 
    DataKeyNames="tmp" PageSize="5" Font-Size="Small">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="CheckBox2" runat="server"  Checked='<%# Bind("checkf") %>' 
                                                                  />
                                    </ItemTemplate>
                                    <ItemStyle Width="20px" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="CustomerNo" HeaderText="客戶代號" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="70px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="CustomerName" HeaderText="(合約)客戶名稱" ReadOnly="True" >
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="200px" />
                                </asp:BoundField>
                                <asp:BoundField HeaderText="合約編號" DataField="ContractNo" ReadOnly="True" >
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="105px" />
                                </asp:BoundField>
                                
                                <asp:BoundField HeaderText="開立方式" DataField="InvoiceCycle" ReadOnly="True" >
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="60px" />
                                </asp:BoundField>
                                
                                <asp:BoundField HeaderText="開立期別" DataField="InvoicePeriod" ReadOnly="True" >
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="60px" />
                                </asp:BoundField>
                                
                                <asp:BoundField HeaderText="每期保養金額(含稅)" DataField="PeriodMaintenanceAmount" ReadOnly="True">
                                <HeaderStyle ForeColor="White" />
                                <ItemStyle Width="80px" HorizontalAlign="Right" />
                                </asp:BoundField>
                                
                                <asp:BoundField DataField="ItemName" HeaderText="品名" >
                                <HeaderStyle ForeColor="White" />
                                </asp:BoundField>
                                
                                <asp:BoundField DataField="tmp" HeaderText="tmp" Visible="False" />
                                <asp:CommandField ShowEditButton="True" />
                                
                            </Columns>
                            <HeaderStyle BackColor="Teal" ForeColor="White" Font-Size="Small" />
                        </asp:GridView>
                    </asp:Panel>
                </td>
            </tr>
            <tr>
                <td>
                                          <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
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
                <td align="center">
                    <asp:Button ID="Button8" runat="server" Text="修改" Visible="False" Width="100px" />
                    <asp:Button ID="Button9" runat="server" Text="刪除" Width="100px" />
                    <asp:Button ID="Button6" runat="server" Text="送出" Width="100px" ToolTip="修改" />
                    <asp:Button ID="Button2" runat="server" Text="全  選" Visible="False" 
                        Width="100px" />
                    <asp:Button ID="Button3" runat="server" style="height: 21px" Text="全不選" 
                        Visible="False" Width="100px" />
                    <asp:Button ID="Button1" runat="server" Text="送出" Width="100px" ToolTip="新增" />
                    <asp:Button ID="Button4" runat="server" Text="取消設定" Width="100px" />
                </td>
            </tr>
            <tr>
                <td class="style80">
         <uc:UsrMsgBox ID="UsrMsgBox1" runat="server" />  
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
            </tr>
        </table>

          <br />
      <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;">
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
      </div>
     
    <div>
    
    </div>
    </form>
</body>
</html>

