<%@ Control Language="VB" AutoEventWireup="false" DEBUG="true" CodeFile="ATT_M_056U.ascx.vb" Inherits="ATT_M_056U" %>

<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />

<asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:Smile_HQConnectionString %>" 
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
       <asp:TextBox ID="txtHDATE" runat="server" Text='<%# Eval("HDATE") %>' Width="88px" ></asp:TextBox>
       <img runat="server" id="imgHDATE" src="../Images/date.gif" alt="按此可點選日期" style="border:0;" />  </td>
  </tr>
  
  <tr>
     <td style="height: 25px; text-align:right">是否為假日：</td>
     <td><asp:CHECKBox ID="CKBHOLIDAY" runat="server"  Width="88px" Font-Size="12pt" />
         <asp:LABEL ID="LABHOLIDAY" runat="server"  Text='<%# Eval("HOLIDAY")  %>' Visible="false" /></td>
  </tr>
              
  

 </table>
 
 


  
  