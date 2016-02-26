<%@ Control Language="VB" DEBUG="true" AutoEventWireup="false" CodeFile="ATT_M_053U.ascx.vb" Inherits="ATT_M_053U" %>

<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
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
         <uc:QryEmp id="QryEmp" runat="server" EmpIDObj="txtEmpID" EmpNameObj="txtEmpName" UpdatePanelObj="udpEmp" />
         <uc:UsrMsgBox ID="UsrMsgBox1" runat="server" />
</div>    

<table  class="table">
  <tr style=" height:5px">
     <td align="right" style="width:15%"> </td>
     <td style="width:85%" > </td>
  </tr>
  
  <tr>
    <td colspan="2" style="width:100%; height: 234px;">
          <table width="100%">
            <tr>
              <td align="right" style="width:25%;height: 25px"><span class="MustInput" >*</span>排班單位：</td>
              <td style="width:75%;height: 25px">
                 <asp:TextBox ID="txtSTNID" runat="server" Value='<%# Eval("STNID") %>' Width="100px" MaxLength="8" AutoPostBack="true" OnTextChanged="txtSTNID_TextChanged" /><asp:TextBox 
                      ID="txtStnName" runat="server" Value='<%# Eval("STNNAME") %>' Width="150px" ></asp:TextBox>
                      <asp:Button ID="btnStnID" runat="server" Text="選擇單位" OnClick="btnStnID_Click" />
              </td>
            </tr>
            <tr>
              <td align="right" style="width:25%;height: 25px"><span class="MustInput" >*</span>員工編號：</td>
              <td>
              <asp:UpdatePanel ID="udpEmp" runat="server" RenderMode="Inline" UpdateMode="Conditional">
                  <ContentTemplate>
                   <asp:TextBox ID="txtEmpID" runat="server" Text='<%# Eval("EmpID") %>' Width="80px" MaxLength="8" AutoPostBack="True" OnTextChanged="txtEmpID_TextChanged" /><asp:TextBox 
                                ID="txtEmpName" runat="server" Text='<%# Eval("EmpNAME") %>' Width="88px"></asp:TextBox>
                   <asp:Button ID="btnEmp" runat="server" Text="選擇員工" OnClick="btnEmp_Click"/> 
                  </ContentTemplate>
              </asp:UpdatePanel></td>
            </tr>
            
            <tr>
              <td align="right" style="height: 25px"><span class="MustInput" >*</span>補卡日期：</td>
              <td><asp:TextBox ID="txtFINDATE" runat="server" Text='<%# Eval("FINDATE") %>'  Width="88px" ></asp:TextBox>
                  </td>
            </tr>                
            
            <tr>
              <td align="right" style="height:25px;"><span class="MustInput" >*</span>補卡時間：</td>
              <td><asp:TextBox ID="txtFINTIME" runat="server" Text='<%# Eval("FINTIME") %>'  Width="50px" />
             </td>
            </tr>
            
            <tr>
              <td align="right" style="height:25px;"><span class="MustInput" >*</span>補登原因：</td>
              <td><asp:TextBox ID="TxtMEMO" runat="server" Text='<%# Eval("MEMO") %>' Width="200px" />
              (限10個中文字)
             </td>
            </tr>
            
            <tr>
              <td align="right" style="width:25%;height: 25px"><asp:Label ID="Label1" runat="server" Text="變更員工編號："></asp:Label></td>
              <td>
                   <asp:TextBox ID="txtEmpID2" runat="server" Width="80px" MaxLength="8" AutoPostBack="True" OnTextChanged="txtEmpID2_TextChanged" /><asp:TextBox 
                                ID="txtEmpName2" runat="server" Width="88px" Enabled="false"></asp:TextBox>
              </td>
            </tr>
              
         </table>
    </td>
  </tr>   
 </table>
 
 <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" > 
    <uc:QryEmp id="QryHrEmp" runat="server" EmpIDObj="txtEmpID" EmpNameObj="txtEmpName" UpdatePanelObj="udpEmp" />
 </div>


  
  