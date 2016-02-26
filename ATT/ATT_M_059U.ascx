<%@ Control Language="VB" AutoEventWireup="false" DEBUG="true" CodeFile="ATT_M_059U.ascx.vb" Inherits="ATT_M_059U" %>

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
       <asp:TextBox ID="txtDATEST" runat="server" Text='<%# Eval("FWDS_DATEST") %>' Width="88px" ></asp:TextBox> ～ 
       <asp:TextBox ID="txtDATEEN" runat="server" Text='<%# Eval("FWDS_DATEEN") %>' Width="88px" ></asp:TextBox>
        </td>
  </tr>
  
  <tr>
     <td style="height: 25px; text-align:right">標準時數：</td>
     <td><asp:TextBox ID="txtSTNWH" runat="server" Text='<%# Eval("FWDS_STNWH") %>' Width="88px" ></asp:TextBox>
     </td>
  </tr>
  

 </table>
 
 


  
  