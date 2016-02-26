<%@ Page Language="VB" AutoEventWireup="false" debug= "false"  CodeFile="InvoiceApply001.aspx.vb" Inherits="MMS_InvoiceApply001" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
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
                        Text="發票開立常態申請" ForeColor="White"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style1">
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
                &nbsp;&nbsp;&nbsp;
                    <asp:Label ID="Label2" runat="server" Text="期        別:" Width="45px"></asp:Label>
                    <asp:TextBox ID="TextBox1" runat="server" Width="88px"></asp:TextBox>
                    <asp:Label ID="Label4" runat="server" Text="(Ex.2014/06)" Width="45px"></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBox ID="CheckBox3" runat="server" AutoPostBack="True" Text="業務員" />
                    <asp:DropDownList ID="DropDownList1" runat="server" Width="110px" 
                        Enabled="False">
                    </asp:DropDownList>
                    &nbsp;&nbsp;&nbsp;
                    <asp:Button ID="Button5" runat="server" Text="查詢" />
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Lable4" runat="server" Text="發票日期: " Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="txtDtTo" runat="server" Width="88px" ></asp:TextBox>
                      <img src="../Images/date.gif" alt="按此可點選日期" style="border:0;" 
                        onclick="ShowCalendar('txtDtTo',true,false)" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Label ID="Label3" runat="server" Text="客戶代號:" Width="68px" Visible="False"></asp:Label>
                    <asp:DropDownList ID="DropDownList5" runat="server" AutoPostBack="True" 
                        Visible="False" Width="118px">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td>
                                              <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                                                  AutoGenerateColumns="False" EnableModelValidation="True" 
                                                  Width="781px" DataKeyNames="pk" Font-Size="Small">
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
                                                      <ItemStyle Width="45px" />
                                                      </asp:BoundField>
                                                      <asp:BoundField DataField="CustomerName" HeaderText="客戶名稱" ReadOnly="True" >
                                                      <HeaderStyle ForeColor="White" />
                                                      <ItemStyle Width="190px" />
                                                      </asp:BoundField>
                                                      <asp:BoundField HeaderText="合約編號" DataField="ContractNo" ReadOnly="True" >
                                                      <HeaderStyle ForeColor="White" />
                                                      <ItemStyle Width="90px" />
                                                      </asp:BoundField>
                                                      <asp:BoundField HeaderText="開立方式" DataField="InvoiceCycle" ReadOnly="True" >
                                                      <HeaderStyle ForeColor="White" />
                                                      <ItemStyle Width="30px" />
                                                      </asp:BoundField>
                                                      <asp:BoundField DataField="InvoicePeriod" HeaderText="期別" ReadOnly="True" >
                                                      <HeaderStyle ForeColor="White" />
                                                      </asp:BoundField>
                                                      <asp:BoundField HeaderText="每期保養金額(含稅)" DataField="PeriodMaintenanceAmount" 
                                                          ReadOnly="True">
                                                      <HeaderStyle ForeColor="White" />
                                                      <ItemStyle Width="80px" HorizontalAlign="Right" />
                                                      </asp:BoundField>
                                                      <asp:BoundField DataField="ItemName" HeaderText="品名" >
                                                      <HeaderStyle ForeColor="White" />
                                                      </asp:BoundField>
                                                      <asp:BoundField DataField="pk" HeaderText="pk" Visible="False" />
                                                      <asp:BoundField DataField="Memo" HeaderText="備註">
                                                      <HeaderStyle ForeColor="White" />
                                                      </asp:BoundField>
                                                      <asp:CommandField ShowEditButton="True" />
                                                  </Columns>
                                                  <HeaderStyle BackColor="Teal" ForeColor="White" Font-Size="Small" />
                                              </asp:GridView>
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
                <td align="center" valign="top">
                    <asp:Button ID="Button2" runat="server" Text="    全選    " Height="25px" 
                        Width="100px" />
                    <asp:Button ID="Button3" runat="server" Text="  全不選  " Height="25px" 
                        Width="100px" />
                    <asp:Button ID="Button1" runat="server" Text="資料送出" Height="25px" 
                        Width="100px" />
                    <asp:Button ID="Button4" runat="server" Text="回主畫面" Height="25px" 
                        Width="100px" />
                </td>
            </tr>
            <tr>
                <td>
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
