<%@ Page Language="VB"  AutoEventWireup="false" CodeFile="Contract_001.aspx.vb" Inherits="MMS_Contract_001" %>

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
        .style82
        {
            width: 70px;
        }
        .style83
        {
            width: 86px;
        }
        .style84
        {
            width: 537px;
        }
        .style85
        {
            width: 447px;
        }
        .style86
        {
            width: 156px;
        }
        .style87
        {
            width: 144px;
        }
        .style90
        {
            width: 120px;
        }
        .style91
        {
            width: 151px;
        }
        .style93
        {
            width: 122px;
        }
        .style94
        {
            width: 150px;
        }
        .style95
        {
            width: 121px;
        }
        .style97
        {
            width: 10px;
        }
        .style98
        {
            width: 145px;
        }
        .style99
        {
            width: 153px;
        }
        .style100
        {
            width: 154px;
        }
        .style101
        {
            width: 61px;
        }
        .style102
        {
            width: 205px;
        }
        .style103
        {
            width: 79px;
        }
        .style105
        {
            width: 123px;
        }
        .style106
        {
            width: 89px;
        }
        .style107
        {
            width: 66px;
        }
        .style108
        {
            width: 115px;
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
                              style="height: 25px; text-align: center; background-color: #1c5e55" 
                              valign="middle">
                              <asp:Label ID="TitleLabel" runat="server" CssClass="titles2" Text="功能名稱"></asp:Label>
                          </td>
                      </tr>
                      <tr>
                          <td>
                              <table cellpadding="3" style="width: 800px">
                                  <tr>
                                      <td class="style9">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td class="style83">
                                                      <asp:Button ID="Button5" runat="server" Text="續  約" />
                                                  </td>
                                                  <td class="style103">
                                                      原合約編號:</td>
                                                  <td class="style34">
                                                      <asp:TextBox ID="OldContractNoTextBox" runat="server" 
                                                          Text='<%# Bind("OldContractNo") %>'  Width="120px" MaxLength="14" />
                                                  </td>
                                                  <td class="style57">
                                                      合約編號:</td>
                                                  <td class="style55">
                                                      <asp:TextBox ID="ContractNoTextBox" runat="server" 
                                                          Text='<%# Bind("ContractNo") %>' Width="130px" MaxLength="14" />
                                                       <asp:CheckBox ID="CheckBox1" runat="server" Text="年度檢查" Visible=false ></asp:CheckBox>    
                                                  </td>
                                                  <td align="right" class="style29">
                                                      <span class="style79">*</span>區域:</td>
                                                  <td width="121">
                                                      <asp:DropDownList ID="DropDownList5" runat="server" AutoPostBack="True" DataTextField="AreaName" 
                                                          DataValueField="AreaCode" Width="121px" Enabled="False">
                                                      </asp:DropDownList>
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td class="style9">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>客戶代號:</td>
                                                  <td class="style86">
                                                      <asp:TextBox ID="CustomerNoTextBox" runat="server" AutoPostBack="True" 
                                                          Text='<%# Bind("CustomerNo") %>' Width="110px" MaxLength="15" />
                                                      <asp:Button ID="Button4" runat="server" Font-Bold="True" 
                                                          Text="..." Width="20px" Visible="False" />
                                                  </td>
                                                  <td class="style24">
                                                      <asp:TextBox ID="CustomerNameTextBox" runat="server" 
                                                          Text='<%# Bind("CustomerName") %>' Width="100%" MaxLength="75" Enabled="False" />
                                                  </td>
                                                  <td align="right" class="style29" width="63px">
                                                      <span class="style79">*</span>業務員:</td>
                                                  <td width="121">
                                                      <asp:DropDownList ID="DropDownList6" runat="server" 
                                                          DataSourceID="SqlDataSource3" DataTextField="EmployeeName" 
                                                          DataValueField="EmployeeNo" Width="121px">
                                                      </asp:DropDownList>
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td class="style9">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      聯絡電話:</td>
                                                  <td class="style102">
                                                      <asp:TextBox ID="TelNoTextBox" runat="server" Text='<%# Bind("TelNo") %>' 
                                                          MaxLength="20" />
                                                  </td>
                                                  <td align="right" class="style101">
                                                      聯絡人:</td>
                                                  <td class="style64">
                                                      <asp:TextBox ID="ContactTextBox" runat="server" Text='<%# Bind("Contact") %>' 
                                                          MaxLength="25" />
                                                  </td>
                                                  <td align="right" class="style29">
                                                      <span class="style79">*</span>收款員:</td>
                                                  <td width="121">
                                                      <asp:DropDownList ID="DropDownList7" runat="server" 
                                                          DataSourceID="SqlDataSource4" DataTextField="EmployeeName" 
                                                          DataValueField="EmployeeNo" Width="121px">
                                                      </asp:DropDownList>
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td class="style9" style="background-color: skyblue">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>大樓名稱:</td>
                                                  <td class="style84">
                                                      <asp:TextBox ID="BuildingNameTextBox" runat="server" 
                                                          Text='<%# Bind("BuildingName") %>' Width="100%" MaxLength="75" />
                                                  </td>
                                                  <td align="right" class="style21">
                                                      <asp:RadioButtonList ID="RadioButtonList2" runat="server" 
                                                          RepeatDirection="Horizontal" >
                                                          <asp:ListItem Value="1">一般</asp:ListItem>
                                                          <asp:ListItem Value="2">全責</asp:ListItem>
                                                          <asp:ListItem Selected="True" Value="3">其他</asp:ListItem>
                                                      </asp:RadioButtonList>
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td class="style9" style="background-color: skyblue">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>地址:</td>
                                                  <td class="style21">
                                                      <asp:TextBox ID="AddressTextBox" runat="server" Text='<%# Bind("Address") %>' 
                                                          Width="99%" MaxLength="100" />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td class="style9" style="background-color: skyblue">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      規格/機號:</td>
                                                  <td class="style85">
                                                      <asp:TextBox ID="SpecificationNoTextBox" runat="server" 
                                                          Text='<%# Bind("SpecificationNo") %>' MaxLength="100" />
                                                  </td>
                                                  <td align="right" class="style47">
                                                      <span class="style79">*</span>台數:</td>
                                                  <td width="121">
                                                      <asp:TextBox ID="QuantityTextBox" runat="server" 
                                                          Text='<%# Bind("Quantity") %>' />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td class="style9" style="background-color: skyblue">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>合約期間:</td>
                                                  <td class="style93">
                                                      <asp:TextBox ID="StartDateCTextBox" runat="server" 
                                                          Text='<%# Bind("StartDateC") %>' Width="80px"  
                                                          onchange="__doPostBack('StartDateCTextBox','')" AutoPostBack="True" />
                                                      <asp:Image ID="Image1" runat="server" ImageUrl="../Images/date.gif" onclick="ShowCalendar('StartDateCTextBox',true,false,'EndDateCTextBox')"/>
                                                  </td>
                                                  <td class="style33">
                                                      <asp:Label ID="Label2" runat="server" Text="~"></asp:Label>
                                                  </td>
                                                  <td class="style105">
                                                      <asp:TextBox ID="EndDateCTextBox" runat="server" Text='<%# Bind("EndDateC") %>' 
                                                          Width="80px" AutoPostBack="True" />
                                                      <asp:Image ID="Image2" runat="server" ImageUrl="../Images/date.gif" 
                                                          onclick="ShowCalendar('EndDateCTextBox',true,false)" />
                                                  </td>
                                                  <td align="right" class="style65">
                                                      <asp:Button ID="Button6" runat="server" Text="計算" />
                                                      總月數:</td>
                                                  <td class="style106">
                                                      <asp:TextBox ID="TotalOfMonthsTextBox" runat="server" 
                                                          Text='<%# Bind("TotalOfMonths") %>' Width="67px" />
                                                  </td>
                                                  <td align="right" class="style47">
                                                      <span class="style79">*</span>總台數:</td>
                                                  <td width="121">
                                                      <asp:TextBox ID="TotalQuantityTextBox" runat="server" 
                                                          Text='<%# Bind("TotalQuantity") %>' />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td class="style9" style="background-color: skyblue">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      每月金額:</td>
                                                  <td class="style91">
                                                      <asp:TextBox ID="PricePerMonthTextBox" runat="server" 
                                                          Text='<%# Bind("PricePerMonth") %>' AutoPostBack="True" />
                                                  </td>
                                                  <td align="right" class="style95">
                                                      <span class="style79">*</span>每台每月金額:</td>
                                                  <td class="style86">
                                                      <asp:TextBox ID="UnitPricePerMonthTextBox" runat="server" 
                                                          Text='<%# Bind("UnitPricePerMonth") %>' />
                                                  </td>
                                                  <td align="right" class="style47">
                                                      <span class="style79">*</span>合約總金額:</td>
                                                  <td width="121">
                                                      <asp:TextBox ID="AmountOfContractTextBox" runat="server" 
                                                          Text='<%# Bind("AmountOfContract") %>' />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="Lime" class="style9">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>收款方式:</td>
                                                  <td class="style94">
                                                      <asp:DropDownList ID="DropDownList8" runat="server" Width="141px">
                                                          <asp:ListItem Value="1">支票</asp:ListItem>
                                                          <asp:ListItem Value="2">匯款</asp:ListItem>
                                                          <asp:ListItem Value="3">現金</asp:ListItem>
                                                      </asp:DropDownList>
                                                  </td>
                                                  <td align="right" class="style93">
                                                      <span class="style79">*</span>收款日數:</td>
                                                  <td class="style100">
                                                      <asp:TextBox ID="DaysAllowedTextBox" runat="server" 
                                                          Text='<%# Bind("DaysAllowed") %>' />
                                                  </td>
                                                  <td align="right" class="style47">
                                                      <span class="style79">*</span>每期保養金額:</td>
                                                  <td width="121">
                                                      <asp:TextBox ID="PeriodMaintenanceAmountTextBox" runat="server" 
                                                          Text='<%# Bind("PeriodMaintenanceAmount") %>' />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="Lime" class="style9">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>總保養金額:</td>
                                                  <td class="style94">
                                                      <asp:TextBox ID="MaintenanceAmountTextBox" runat="server" 
                                                          Text='<%# Bind("MaintenanceAmount") %>' />
                                                  </td>
                                                  <td align="right" class="style93">
                                                      調整後總保養金額:</td>
                                                  <td class="style28">
                                                      <asp:TextBox ID="AdjAmountOfContractTextBox" runat="server" 
                                                          Text='<%# Bind("AdjAmountOfContract") %>' />
                                                  </td>
                                                  <td class="style47">
                                                      &nbsp;</td>
                                                  <td class="style71">
                                                      &nbsp;</td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#FFCC66" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      統一編號:</td>
                                                  <td class="style91">
                                                      <asp:TextBox ID="GUINumberTextBox" runat="server" 
                                                          Text='<%# Bind("GUINumber") %>' AutoPostBack="True" MaxLength="8" />
                                                  </td>
                                                  <td align="right" class="style95">
                                                      發票抬頭:</td>
                                                  <td class="style28">
                                                      <asp:TextBox ID="TitleOfInvoiceTextBox" runat="server" 
                                                          Text='<%# Bind("TitleOfInvoice") %>' MaxLength="75" Width="400px" />
                                                  </td>
                                                  <td>
                                                      &nbsp;</td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#FFCC66" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      發票地址:</td>
                                                  <td class="style21">
                                                      <asp:TextBox ID="AddressOfInvoiceTextBox" runat="server" 
                                                          Text='<%# Bind("AddressOfInvoice") %>' Width="99%" MaxLength="100" />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#FFCC66" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>開立期間:</td>
                                                  <td class="style87">
                                                      <asp:TextBox ID="StartDateITextBox" runat="server" 
                                                          Text='<%# Bind("StartDateI") %>' Width="100px" />
                                                      <asp:Image ID="Image3" runat="server" ImageUrl="../Images/date.gif" 
                                                          onclick="ShowCalendar('StartDateITextBox',true,false,'EndDateITextBox')"/>
                                                  </td>
                                                  <td class="style97">
                                                      <asp:Label ID="Label5" runat="server" Text="~"></asp:Label>
                                                  </td>
                                                  <td class="style98">
                                                      <asp:TextBox ID="EndDateITextBox" runat="server" Text='<%# Bind("EndDateI") %>' 
                                                          Width="100px" />
                                                      <asp:Image ID="Image4" runat="server" ImageUrl="../Images/date.gif" 
                                                          onclick="ShowCalendar('EndDateITextBox',true,false)"/>
                                                  </td>
                                                  <td class="style82">
                                                      <span class="style79">*</span>開立方式:</td>
                                                  <td class="style107">
                                                      <asp:DropDownList ID="DropDownList4" runat="server" 
                                                          Width="60px">
                                                          <asp:ListItem Value="1">年</asp:ListItem>
                                                          <asp:ListItem Value="2">季</asp:ListItem>
                                                          <asp:ListItem Value="3">半年</asp:ListItem>
                                                          <asp:ListItem Value="4">二月</asp:ListItem>
                                                          <asp:ListItem Value="5">月</asp:ListItem>
                                                          <asp:ListItem Value="6">二年</asp:ListItem>
                                                      </asp:DropDownList>
                                                  </td>
                                                  <td class="style108">
                                                      &nbsp;</td>
                                                  <td>
                                                      &nbsp;</td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#FFCC66" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>發票品名:</td>
                                                  <td class="style78">
                                                      <asp:TextBox ID="ItemNameTextBox" runat="server" Text='<%# Bind("ItemName") %>' 
                                                          Width="423px" MaxLength="25" />
                                                  </td>
                                                  <td>
                                                      <asp:RadioButtonList ID="RadioButtonList3" runat="server" 
                                                          RepeatDirection="Horizontal" >
                                                          <asp:ListItem Value="1">品名+年月</asp:ListItem>
                                                          <asp:ListItem Value="2">年月+品名</asp:ListItem>
                                                          <asp:ListItem Value="3">品名</asp:ListItem>
                                                      </asp:RadioButtonList>
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#FFCC66" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      <span class="style79">*</span>付款人:</td>
                                                  <td class="style91">
                                                      <asp:TextBox ID="PayerTextBox" runat="server" Text='<%# Bind("Payer") %>' 
                                                          MaxLength="25" />
                                                  </td>
                                                  <td align="right" class="style93">
                                                      付款人電話:</td>
                                                  <td class="style28">
                                                      <asp:TextBox ID="PayerTelTextBox" runat="server" 
                                                          Text='<%# Bind("PayerTel") %>' MaxLength="20" />
                                                  </td>
                                                  <td class="style29">
                                                      &nbsp;</td>
                                                  <td>
                                                      &nbsp;</td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#FFCC66" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83">
                                                      付款人地址:</td>
                                                  <td class="style21">
                                                      <asp:TextBox ID="PayerAdderssTextBox" runat="server" 
                                                          Text='<%# Bind("PayerAdderss") %>' Width="99%" MaxLength="100" 
                                                          style="margin-bottom: 0px" />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#6699FF" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83" style="color: #FFFFFF">
                                                      <span class="style79">*</span>產品別:</td>
                                                  <td class="style99">
                                                      <asp:TextBox ID="ItemNoTextBox" runat="server" Text='<%# Bind("ItemNo") %>' 
                                                          MaxLength="75" />
                                                  </td>
                                                  <td align="right" class="style95" style="color: #FFFFFF">
                                                      合約歸檔日-機械:</td>
                                                  <td class="style28">
                                                      <asp:TextBox ID="ArchiveDateMTextBox" runat="server" 
                                                          Text='<%# Bind("ArchiveDateM") %>' Width="80px" />
                                                      <asp:Image ID="Image5" runat="server" ImageUrl="../Images/date.gif" 
                                                          onclick="ShowCalendar('ArchiveDateMTextBox',true,false)" />
                                                  </td>
                                                  <td class="style29">
                                                      &nbsp;</td>
                                                  <td>
                                                      &nbsp;</td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#6699FF" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83" style="color: #FFFFFF">
                                                      合約用印日:</td>
                                                  <td class="style100">
                                                      <asp:TextBox ID="SealDateTextBox" runat="server" 
                                                          Text='<%# Bind("SealDate") %>' Width="80px" />
                                                      <asp:Image ID="Image8" runat="server" ImageUrl="../Images/date.gif" 
                                                          onclick="ShowCalendar('SealDateTextBox',true,false)" />
                                                  </td>
                                                  <td align="right" class="style90" style="color: #FFFFFF">
                                                      合約歸檔日-會計:</td>
                                                  <td class="style28">
                                                      <asp:TextBox ID="ArchiveDateATextBox" runat="server" 
                                                          Text='<%# Bind("ArchiveDateA") %>' Width="80px" />
                                                      <asp:Image ID="Image6" runat="server" ImageUrl="../Images/date.gif" 
                                                          onclick="ShowCalendar('ArchiveDateATextBox',true,false)" />
                                                  </td>
                                                  <td align="right" class="style61" style="color: #FFFFFF">
                                                      合約取消確定日:</td>
                                                  <td>
                                                      <asp:TextBox ID="CancelDateTextBox" runat="server" 
                                                          Text='<%# Bind("CancelDate") %>' Width="80px" />
                                                      <asp:Image ID="Image7" runat="server" ImageUrl="../Images/date.gif" 
                                                         onclick="ShowCalendar('CancelDateTextBox',true,false)"  />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                                  <tr>
                                      <td bgcolor="#6699FF" class="style23">
                                          <table style="width: 795px;">
                                              <tr>
                                                  <td align="right" class="style83" style="color: #FFFFFF" valign="top">
                                                      備註:</td>
                                                  <td class="style21">
                                                      <asp:TextBox ID="MemoTextBox" runat="server" Height="80px" 
                                                          Text='<%# Bind("Memo") %>' Width="99%" MaxLength="100" 
                                                          TextMode="MultiLine" />
                                                  </td>
                                              </tr>
                                              
                                              <tr>
                                                  <td align="right" class="style83" style="color: #FFFFFF" valign="top">
                                                      發票列印備註:</td>
                                                  <td class="style21">
                                                      <asp:TextBox ID="Memo2TextBox" runat="server" Height="80px" 
                                                          Text='<%# Bind("Memo") %>' Width="99%" MaxLength="50" 
                                                          TextMode="MultiLine" />
                                                  </td>
                                              </tr>
                                          </table>
                                      </td>
                                  </tr>
                              </table>
                          </td>
                      </tr>
                      <tr>
                          <td align="center">
                              &nbsp;</td>
                      </tr>
                      <tr>
                          <td align="center">
                              <asp:Button ID="Button7" runat="server" Height="21px" Text="修改" Width="88px" />
                              <asp:Button ID="ButtonSave" runat="server" Text="資料送出" Height="21px" 
                                  ToolTip="新增" Width="88px" />
                              <asp:Button ID="Button2" runat="server" Text="資料送出" Height="21px" ToolTip="修改" 
                                  Width="89px" />
                              <asp:Button ID="Button8" runat="server" Height="21px" Text="清除重來" />
                              <asp:Button ID="Button3" runat="server" Text="回主畫面" Height="21px" 
                                  Width="83px"  />
                          </td>
                      </tr>
                      <tr>
                          <td>
                              <asp:TextBox ID="TextBox1" runat="server" Visible="False"></asp:TextBox>
                              <asp:TextBox ID="TextBox2" runat="server" Visible="False"></asp:TextBox>
                          </td>
                      </tr>
                      <tr>
                          <td>
                              <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                                  SelectCommand="SELECT * FROM [MMSContract]"></asp:SqlDataSource>
                              <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                                  
                                  
                                  SelectCommand="SELECT AreaCode, AreaCode+'_'+AreaName  as AreaName FROM MMSArea WHERE (Effective = 'Y')">
                              </asp:SqlDataSource>
                              <asp:SqlDataSource ID="SqlDataSource3" runat="server" 
                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" SelectCommand="SELECT EmployeeNo, EmployeeNo+'-'+EmployeeName  as EmployeeName  FROM MMSEmployee
where Effective = 'Y'"></asp:SqlDataSource>
                              <asp:SqlDataSource ID="SqlDataSource4" runat="server" 
                                  ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" SelectCommand="SELECT EmployeeNo, EmployeeNo+'-'+EmployeeName  as EmployeeName  FROM MMSEmployee
where Effective = 'Y'
"></asp:SqlDataSource>
                          </td>
                      </tr>
                      <tr>
                          <td>
                              <asp:Panel ID="pnlCmp" runat="server" BackColor="White" SkinID="selPanel" 
                                  Width="402px" Visible="False">
                                  <div class="selTitle" style="width:400px">
                                      選擇客戶</div>
   <br />
                                  <asp:UpdatePanel ID="updQryCmp" runat="server" RenderMode="Inline" 
                                      UpdateMode="Conditional">
                                      <ContentTemplate>
                                          代號：<asp:TextBox ID="txtQCmpID" runat="server" MaxLength="6" Width="50px"></asp:TextBox>
                                          &nbsp;名稱：<asp:TextBox ID="txtQSName" runat="server" Width="100px"></asp:TextBox>
                                          <asp:Button ID="btnSearch" runat="server"  
                                              style="width:50px" Text="查詢" />
                                          <asp:Button ID="btnCancel" runat="server" style="width:50px" Text="取消" />
       <br />
                                          <asp:Label ID="lblMsg" runat="server" CssClass="msg" Visible="false"></asp:Label>
                                          <asp:Panel ID="pnlSelect" runat="server" CssClass="fixedHeader" Height="475px">
                                              <asp:TextBox ID="txtRet" runat="server" style="display:none" Width="20px" />
                                              <asp:GridView ID="grdCmp" runat="server" AllowPaging="True" 
                                                  AutoGenerateColumns="False" DataKeyNames="CustomerNo" 
                                                  EnableModelValidation="True" 
                                                  PagerSettings-FirstPageImageUrl="~/Images/First.gif" 
                                                  PagerSettings-LastPageImageUrl="~/Images/Last.gif" 
                                                  PagerSettings-Mode="NextPreviousFirstLast" 
                                                  PagerSettings-NextPageImageUrl="~/Images/Next.gif" 
                                                  PagerSettings-PreviousPageImageUrl="~/Images/Previous.gif" PageSize="15" 
                                                  SkinID="selGrid" Width="368px">
                                                  <Columns>
                                                      <asp:TemplateField ShowHeader="False">
                                                          <ItemTemplate>
                                                              <asp:Button ID="Button1" runat="server" CausesValidation="False" 
                                                                  CommandName="Select" Text="選取" onclick="Button1_Click" />
                                                          </ItemTemplate>
                                                      </asp:TemplateField>
                                                      <asp:BoundField DataField="CustomerNo" HeaderText="客戶編號">
                                                      <ItemStyle Width="100px" />
                                                      </asp:BoundField>
                                                      <asp:BoundField DataField="CustomerName" HeaderText="客戶名稱">
                                                      <ItemStyle Width="200px" />
                                                      </asp:BoundField>
                                                  </Columns>
                                                  <HeaderStyle BackColor="Teal" ForeColor="White" />
                                                  <PagerSettings FirstPageImageUrl="~/Images/First.gif" 
                                                      LastPageImageUrl="~/Images/Last.gif" Mode="NextPreviousFirstLast" 
                                                      NextPageImageUrl="~/Images/Next.gif" 
                                                      PreviousPageImageUrl="~/Images/Previous.gif" />
                                              </asp:GridView>
                                          </asp:Panel>
                                      </ContentTemplate>
                                      <Triggers>
                                          <asp:AsyncPostBackTrigger ControlID="grdCmp" EventName="PageIndexChanging" />
                                      </Triggers>
                                  </asp:UpdatePanel>
                                  <ajaxToolkit:ModalPopupExtender ID="popCmp" runat="server" 
                                      BackgroundCssClass="modalBackground" PopupControlID="pnlCmp" 
                                      TargetControlID="Button4" OkControlID="Button4" 
                                       />
                                  <asp:SqlDataSource ID="dscCmp" runat="server" 
                                      ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>" 
                                      DataSourceMode="DataSet" EnableViewState="False"></asp:SqlDataSource>
                              </asp:Panel>
                          </td>
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
