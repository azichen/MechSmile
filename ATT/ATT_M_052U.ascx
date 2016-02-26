<%@ Control Language="VB" AutoEventWireup="false" CodeFile="ATT_M_052U.ascx.vb" Inherits="ATT_M_052U" %>

<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
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

<br />

<asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>" 
      DataSourceMode="DataReader">
</asp:SqlDataSource>     

<div style="width:10px;height:10px;position:absolute;left:-200px;top:-10px;" >
         <uc:QryStn ID="QryStn" runat="server" UpdatePanelObj="updQry" StnIDObj="txtStnID" StnNameObj="txtStnName" />  
         <uc:QryEmp id="QryEmp" runat="server" EmpIDObj="txtEmpID" EmpNameObj="txtEmpName" UpdatePanelObj="udpEmp" />
         <uc:QryVcm id="QryVcm" runat="server" EmpIDObj="TxtVANMID" EmpNameObj="TxtVANMNAME" UpdatePanelObj="UpdVCM" />         
         <uc:QryVcm2 id="QryVcm2" runat="server" EmpIDObj="TxtVANMID" EmpNameObj="TxtVANMNAME" UpdatePanelObj="UpdVCM" /> 
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
              <td align="right" style="width:25%;height: 25px"><span class="MustInput" >*</span>員工編號：</td>
              <td>
              <asp:UpdatePanel ID="udpEmp" runat="server" RenderMode="Inline" UpdateMode="Conditional">
                  <ContentTemplate>
                   <asp:TextBox ID="txtEmpID" runat="server" Text='<%# Eval("VATM_EMPLID") %>' Width="80px" MaxLength="8" AutoPostBack="True" OnTextChanged="txtEmpID_TextChanged" /><asp:TextBox 
                                ID="txtEmpName" runat="server" Text='<%# Eval("Empl_NAME") %>' Width="88px"></asp:TextBox>
                   <asp:Button ID="btnEmp" runat="server" Text="選擇員工" OnClick="btnEmp_Click" Visible=false /> 
                  </ContentTemplate>
              </asp:UpdatePanel></td>
            </tr>
            
            <tr>
                  <td style="height: 25px; text-align:right;">請假別：</td>
                  <td style="text-align:left;">
                      <asp:UpdatePanel ID="UpdVCM" runat="server" RenderMode="Inline" UpdateMode="Conditional">
                      <ContentTemplate>
                      <asp:TextBox ID="TxtVANMID" runat="server" Text='<%# Eval("VATM_VANMID") %>' Width="40px" MaxLength="3" AutoPostBack="True" OnTextChanged="TxtVANMID_TextChanged" /><asp:TextBox 
                           ID="TxtVANMNAME" runat="server" Text='<%# Eval("VATM_NAME") %>' Width="100px"></asp:TextBox>
                      <asp:Button ID="BTNVCM" runat="server" Text="選擇假別" OnClick="btnVCM_Click"/> 
                      </ContentTemplate>
                      </asp:UpdatePanel></td>
               </tr>  
            
            <tr>
              <td align="right" style="height: 25px"><span class="MustInput" >*</span>請假日期(起)：</td>
              <td>
                  <asp:TextBox ID="txtVATMDATEST" runat="server" Text='<%# Eval("VATM_DATE_ST") %>'  Width="88px" > </asp:TextBox>
                  起始時間：
                  <asp:TextBox ID="txtVATMTIMEST" runat="server" Text='<%# Eval("VATM_TIME_ST") %>' MaxLength="4" Width="50px" onblur="CheckTime(this);" />
                  <asp:TextBox ID="txtSTDT_tmp" runat="server" Text='<%# Eval("VATM_DATE_ST") %>' Visible="False" Width="88px"></asp:TextBox>
                  <asp:Label ID="LEditmode" runat="server" Text="Label" Visible = "false" ></asp:Label></td>
            </tr>                
            
            <tr>
              <td align="right" style="height:25px;"><span class="MustInput" >*</span>請假日期(迄)：</td>
              <td><asp:TextBox ID="txtVATMDATEEN" runat="server" Text='<%# Eval("VATM_DATE_EN") %>'  Width="88px" ></asp:TextBox> 結束時間：
                  <asp:TextBox ID="txtVATMTIMEEN" runat="server" Text='<%# Eval("VATM_TIME_EN") %>' MaxLength="4" Width="50px" onblur="CheckTime(this);" />
                  <asp:TextBox ID="txtSTTM_tmp" runat="server" Text='<%# Eval("VATM_TIME_ST") %>' MaxLength="4" Width="50px" onblur="CheckTime(this);" Visible="False" />
                  <asp:TextBox ID="txtENTM_tmp" runat="server" Text='<%# Eval("VATM_TIME_ST") %>' MaxLength="4" Width="50px" onblur="CheckTime(this);" Visible="False" /></td>
            </tr>
               
              <tr>
                  <td align="right" style="height:25px;">請假時數：</td>
                  <td> <asp:TextBox ID="txtVATMHOUR" runat="server" Text='<%# Eval("VATM_HOURS") %>' Width="50px" /> 小時 
                  <asp:Button ID="BtnHourCal" runat="server" Text="時數計算" OnClick="BtnHourCal_Click"/> 
                  </td>
              </tr>
              
              
              <tr>
                  <td align="right" style="height:25px;">請假事由：</td>
                  <td> <asp:TextBox ID="txtREMARK" runat="server" Text='<%# Eval("VATM_REMARK") %>' Width="300px" />  
                  </td>
              </tr>
              
              <tr>
                  <td ></td> 
                  <td> <asp:TextBox ID="txtVATMSEQ" runat="server" Text='<%# Eval("VATM_SEQ") %>' Width="50px" Visible="false" />  
                  </td>
              </tr>
         </table>
    </td>
  </tr>   
 </table>
 
 <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" > 
    <uc:QryEmp id="QryHrEmp" runat="server" EmpIDObj="txtEmpID" EmpNameObj="txtEmpName" UpdatePanelObj="udpEmp" />
 </div>


  
  