<%@ Control Language="VB" DEBUG="false" AutoEventWireup="false" CodeFile="ATT_M_054U.ascx.vb" Inherits="ATT_M_054U" %>

<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />

<script type="text/javascript">
//**************************************************************************
//* 檢核時間之輸入格式
//**************************************************************************
function CheckTime(pTimeObj)
{ 
  var sFmt=/((0|1)(\d)|(2)([0-3]))([0-5])(\d)/;
  // (1)23:59=>/((0|1)(\d)|(2)([0-3]))(:)([0-5])(\d)/;  
  // (2)2359=>/((0|1)(\d)|(2)([0-3]))([0-5])(\d)/;  
  // (3)兩者皆可=> /(((0|1)(\d)|(2)([0-3]))(:)([0-5])(\d)|((0|1)(\d)|(2)([0-3]))([0-5])(\d))/;
  
  if((!sFmt.test(pTimeObj.value))&&!(String(pTimeObj.value).length==0)) {
      alert('輸入的時間格式錯誤(須4碼)！\n\n例：0000 ~ 2359');
      pTimeObj.select();pTimeObj.focus();return false;
  }
}
</script>

<asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>" 
      DataSourceMode="DataReader">
</asp:SqlDataSource>     

<div style="width:10px;height:10px;position:absolute;left:-200px;top:-10px;" >
         <uc:QryStn ID="QryStn" runat="server" UpdatePanelObj="updQry" StnIDObj="txtStnID" StnNameObj="txtStnName" />  
         <uc:UsrMsgBox ID="UsrMsgBox1" runat="server" />
</div>    

<table  class="table">
  <tr style=" height:5px">
     <td align="right" style="width:15%"> </td>
     <td style="width:85%" > </td>
  </tr>
  
  <tr>
    <td colspan="2" style="width:100%;">
          <table width="100%">
            <tr>
              <td align="right" style="width:25%;height: 25px"><span class="MustInput" >*</span>單位：</td>
              <td style="width:75%;height: 25px">
                 <asp:TextBox ID="txtSTNID" runat="server" Value='<%# Eval("STNID") %>' Width="100px" MaxLength="8" AutoPostBack="true" OnTextChanged="txtSTNID_TextChanged" /><asp:TextBox 
                      ID="txtStnName" runat="server" Value='<%# Eval("STNNAME") %>' Width="150px" ></asp:TextBox>
                      <asp:Button ID="btnStnID" runat="server" Text="選擇單位" OnClick="btnStnID_Click" />
              </td>
            </tr>
            
            <tr>
              <td align="right" style="height: 25px"><span class="MustInput" >*</span>排班代號：</td>
              <td><asp:TextBox ID="txtCLSID" runat="server" Text='<%# Eval("CLSID") %>'  Width="88px" ></asp:TextBox>
              </td>
            </tr>                
            
            <tr>
              <td align="right" style="height:25px;"><span class="MustInput" >*</span>排班時間：</td>
              <td><asp:TextBox ID="txtSTTIME" runat="server" Text='<%# Eval("STTIME") %>' MaxLength="4" Width="50px" /> 至 
              <asp:TextBox ID="txtEDTIME" runat="server" Text='<%# Eval("EDTIME") %>' MaxLength="4" Width="50px" />
             </td>
            </tr>
              
            <tr>
              <td align="right" style="height: 25px">正常班：</td>
              <td><asp:CheckBox ID="chkNflag" runat="server" Checked='<%# IIf(eval("NFlag")="Y",true,false) %>' Text="正常班" AutoPostBack="true" />
              </td>
            </tr> 
            
         </table>
    </td>
  </tr>   
 </table>



  
  