<%@ Control Language="VB" AutoEventWireup="false" DEBUG="true" CodeFile="ATT_M_055U.ascx.vb" Inherits="ATT_M_055U" %>

<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
       
<div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" >      
     <uc:QryStn ID="QryStn" runat="server" UpdatePanelObj="udpStn" StnIDObj="txtStnID" StnNameObj="txtStnName" StnID_TextChanged="txtStnID_TextChanged"/>  
</div>

 <br /><br />
<table class="table"  >  

  <tr style="vertical-align:bottom">
     <td align="right" style="height: 30px"><span class="MustInput" >*</span>加油站號：</td>
     <td colspan="3">
       <asp:UpdatePanel ID="udpStn" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>     
            <asp:TextBox ID="txtStnID" runat="server"  Text='<%# Eval("STNID") %>' Width="60px" AutoPostBack="true" MaxLength="6" 
                   /><asp:TextBox ID="txtStnName" runat="server" Width="150px" />
                     <asp:Button ID="btnStnID" runat="server" Text="選擇單位" />
         </ContentTemplate>
       </asp:UpdatePanel></td>
  </tr>   

  <tr style="vertical-align:bottom">
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>晴天平日工時：</td>
     <td><asp:TextBox ID="txtHR11" runat="server" Text='<%# Eval("HR11") %>' Width="88px" Font-Size="12pt" /></td>
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>晴天假日工時：</td>
     <td><asp:TextBox ID="txtHR12" runat="server" Text='<%# Eval("HR12") %>' Width="88px" Font-Size="12pt" /></td>
  </tr>
  <tr style="vertical-align:bottom">
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>晴天漲價工時：</td>
     <td><asp:TextBox ID="txtHR13" runat="server" Text='<%# Eval("HR13") %>' Width="88px" Font-Size="12pt" /></td>
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>晴天跌價工時：</td>
     <td><asp:TextBox ID="txtHR14" runat="server" Text='<%# Eval("HR14") %>' Width="88px" Font-Size="12pt" /></td>
  </tr>
  <tr>
     <td align="right" style="width: 20%"></td>
     <td style="width: 30%">&nbsp;</td>
     <td align="right" style="width: 20%"></td>
     <td style="width: 30%"></td>
  </tr>
  
  <tr style="vertical-align:bottom">
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>陰天平日工時：</td>
     <td><asp:TextBox ID="txtHR21" runat="server" Text='<%# Eval("HR21") %>' Width="88px" Font-Size="12pt" /></td>
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>陰天假日工時：</td>
     <td><asp:TextBox ID="txtHR22" runat="server" Text='<%# Eval("HR22") %>' Width="88px" Font-Size="12pt" /></td>
  </tr>
  <tr style="vertical-align:bottom">
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>陰天漲價工時：</td>
     <td><asp:TextBox ID="txtHR23" runat="server" Text='<%# Eval("HR23") %>' Width="88px" Font-Size="12pt" /></td>
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>陰天跌價工時：</td>
     <td><asp:TextBox ID="txtHR24" runat="server" Text='<%# Eval("HR24") %>' Width="88px" Font-Size="12pt" /></td>
  </tr>
  <tr>
     <td align="right" style="width: 20%"></td>
     <td style="width: 30%">&nbsp;</td>
     <td align="right" style="width: 20%"></td>
     <td style="width: 30%"></td>
  </tr>
  
  <tr style="vertical-align:bottom">
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>雨天平日工時：</td>
     <td><asp:TextBox ID="txtHR31" runat="server" Text='<%# Eval("HR31") %>' Width="88px" Font-Size="12pt" /></td>
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>雨天假日工時：</td>
     <td><asp:TextBox ID="txtHR32" runat="server" Text='<%# Eval("HR32") %>' Width="88px" Font-Size="12pt" /></td>
  </tr>
  <tr style="vertical-align:bottom">
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>雨天漲價工時：</td>
     <td><asp:TextBox ID="txtHR33" runat="server" Text='<%# Eval("HR33") %>' Width="88px" Font-Size="12pt" /></td>
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>雨天跌價工時：</td>
     <td><asp:TextBox ID="txtHR34" runat="server" Text='<%# Eval("HR34") %>' Width="88px" Font-Size="12pt" /></td>
  </tr>
  <tr>
     <td align="right" style="width: 20%"></td>
     <td style="width: 30%">&nbsp;</td>
     <td align="right" style="width: 20%"></td>
     <td style="width: 30%"></td>
  </tr>
  
  <tr style="vertical-align:bottom">
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>誤差增加時數：</td>
     <td><asp:TextBox ID="txtDiffADD" runat="server" Text='<%# Eval("DIFFADD") %>' Width="88px" Font-Size="12pt" /></td>
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>誤差減少時數：</td>
     <td><asp:TextBox ID="txtDiffSUB" runat="server" Text='<%# Eval("DIFFSUB") %>' Width="88px" Font-Size="12pt" /></td>
  </tr>
  
  <tr style="vertical-align:bottom">
     <td align="right" style="height: 25px"><span class="MustInput" >*</span>是否檢查：</td>
     <td><asp:CHECKBox ID="CKBCHECKFLAG" runat="server"  Width="88px" Font-Size="12pt" />
         <asp:LABEL ID="LABCHECKFLAG" runat="server"  Text='<%# Eval("CHECKFLAG")  %>' Visible="false" /></td>
     <td align="right" style="width: 20%"></td>
     <td style="width: 30%"></td>
  </tr>
    
  <tr style="vertical-align:bottom">
     <td colspan="2" align="right" style="height: 20px; font-size:12pt"></td>
  </tr>   
</table>