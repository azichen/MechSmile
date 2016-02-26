<%@ Control Language="VB" AutoEventWireup="false" DEBUG="true" CodeFile="ATT_M_061U.ascx.vb" Inherits="ATT_M_061U" %>

<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />

<asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>" 
      DataSourceMode="DataReader">
</asp:SqlDataSource>     

<div style="width:10px;height:10px;position:absolute;left:-200px;top:-10px;" >
         <uc:UsrMsgBox ID="UsrMsgBox1" runat="server" />
</div>    

<table  class="table">
  <tr style=" height:5px">
     <td align="right" style="width:15%"> </td>
     <td style="width:85%" > </td>
  </tr>
  
  <tr >
     <td style="height: 25px; text-align:right"><span class="MustInput" >*</span>日期：</td>
     <td style="text-align:left">
       <asp:TextBox ID="txtDATEST" runat="server" Text='<%# Eval("HODDATE") %>' Width="88px" Font-Bold="True" Font-Size="Small" ></asp:TextBox>
     </td>
  </tr>
  
  <tr >
     <td style="height: 33px; text-align:right"><span class="MustInput" >*</span>時間：</td>
     <td style="text-align:left; height: 33px;">
       <asp:TextBox ID="txtHDTMST" runat="server" Text='<%# Eval("HDTMST") %>' Width="88px" Font-Bold="True" Font-Size="Small" ></asp:TextBox> ～ 
       <asp:TextBox ID="txtHDTMEN" runat="server" Text='<%# Eval("HDTMEN") %>' Width="88px" Font-Bold="True" Font-Size="Small" ></asp:TextBox>
        </td>
  </tr>
  
  <tr >
     <td style="height: 25px; text-align:right"><span class="MustInput" >*</span>時數：</td>
     <td style="text-align:left">
       <asp:TextBox ID="TxtHODHOUR" runat="server" Text='<%# Eval("HODHours") %>' Width="88px" Font-Size="Small" ></asp:TextBox>
     </td>
  </tr>
  
  <tr >
     <td style="height: 25px; text-align:right"><span class="MustInput" >*</span>說明：</td>
     <td style="text-align:left">
       <asp:TextBox ID="TxtMEMO" runat="server" Text='<%# Eval("HODMEMO") %>' Width="306px" ></asp:TextBox>
     </td>
  </tr>
  
  <tr >
     <td style="height: 25px; text-align:right"><span class="MustInput" >*</span>區域：</td>
     <td style="text-align:left">
         <asp:CheckBox ID="ChkC" runat="server" Text="基隆市" /><asp:CheckBox ID="ChkA" runat="server" Text="台北市" /><asp:CheckBox ID="ChkF" runat="server" Text="新北市" />
         <asp:CheckBox ID="ChkH" runat="server" Text="桃園市" /><asp:CheckBox ID="ChkO" runat="server" Text="新竹市" /><asp:CheckBox ID="ChkJ" runat="server" Text="新竹縣" />
         <asp:CheckBox ID="ChkG" runat="server" Text="宜蘭縣" />
         <asp:CheckBox ID="ChkK" runat="server" Text="苗栗縣" />
     </td>
  </tr>
  <tr >
     <td style="height: 25px; text-align:right"></td>
     <td style="text-align:left">
         <asp:CheckBox ID="ChkB" runat="server" Text="南投縣" /><asp:CheckBox ID="ChkN" runat="server" Text="彰化縣" />
         <asp:CheckBox ID="ChkP" runat="server" Text="雲林縣" /><asp:CheckBox ID="ChkI" runat="server" Text="嘉義市" /><asp:CheckBox ID="ChkQ" runat="server" Text="嘉義縣" />
         <asp:CheckBox ID="ChkD" runat="server" Text="台南市" /><asp:CheckBox ID="ChkE" runat="server" Text="高雄市" /><asp:CheckBox ID="ChkT" runat="server" Text="屏東縣" />
     </td>
  </tr>
  <tr >
     <td style="height: 25px; text-align:right"></td>
     <td style="text-align:left">
         <asp:CheckBox ID="ChkU" runat="server" Text="花蓮縣" /><asp:CheckBox ID="ChkV" runat="server" Text="台東縣" /><asp:CheckBox ID="ChkZ" runat="server" Text="連江縣" />
         <asp:CheckBox ID="ChkW" runat="server" Text="金門縣" /><asp:CheckBox ID="ChkX" runat="server" Text="澎湖縣" />         
     </td>
  </tr>

 </table>
 
 


  
  