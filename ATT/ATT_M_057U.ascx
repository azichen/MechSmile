<%@ Control Language="VB" DEBUG="TRUE" AutoEventWireup="false" CodeFile="ATT_M_057U.ascx.vb" Inherits="ATT_M_057U" %>

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

<asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>" 
      DataSourceMode="DataReader">
</asp:SqlDataSource>    

<div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" > 
    <uc:QryEmp id="QryHrEmp" runat="server" EmpIDObj="txtEmpID" EmpNameObj="txtEmpName" UpdatePanelObj="udpEmp" />
    <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
</div>
 
<uc:UsrCalendar ID="UsrCalendar" runat="server" /> 

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
    <td colspan="2" style="width:100%;">
          <table width="100%">
            <tr>
              <td align="right" style="width:25%;height: 25px"><span class="MustInput" >*</span>申請單位：</td>
              <td style="width:75%;height: 25px">
                 <asp:TextBox ID="txtSTNID" runat="server" Value='<%# Eval("HAP_STNID") %>' Width="100px" MaxLength="8" AutoPostBack="true" OnTextChanged="txtSTNID_TextChanged" /><asp:TextBox 
                      ID="txtStnName" runat="server" Value='<%# Eval("STNNAME") %>' Width="150px" ></asp:TextBox>
                      <asp:Button ID="btnStnID" runat="server" Text="選擇單位" OnClick="btnStnID_Click" />&nbsp;&nbsp;&nbsp;
                 <asp:TextBox ID="TxtSERID" runat="server" Value='<%# Eval("HAP_SERID") %>' Width="70px" MaxLength="8" AutoPostBack="true" /></td>
            </tr>
            <tr>
              <td align="right" style="width:25%;height: 25px"><span class="MustInput" >*</span>申請員工：</td>
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
              <td align="right" style="height:25px;"><span class="MustInput" >*</span>申請表單：</td>
              <td><asp:DROPDOWNLIST ID="txtFORMid" runat="server" Value='<%# Eval("FORMNAME") %>' Width="200px" AutoPostBack="True" /> 
                  <asp:TextBox ID="TxtFORMNAME" runat="server" Value='<%# Eval("FORMNAME") %>' Width="200px" VISIBLE="false" />
              </td>
            </tr>
                  
            <tr>
              <td align="right" style="height:25px;"><span class="MustInput" >*</span>連絡人：</td>
              <td><asp:TextBox ID="TxtContacts" runat="server" Value='<%# Eval("Contacts") %>' Width="200px" />
              </td>
            </tr>     

            <asp:UpdatePanel ID="UpdatePanel1" runat="server" RenderMode="Inline" UpdateMode="Conditional">
            <ContentTemplate>
            <tr>
              <td align="right" style="width:25%;height: 25px"><asp:Label ID="Label11" runat="server" Text="起迄年月："></asp:Label></td>
              <td>
                   <asp:TextBox ID="txtDtStart" runat="server" Value='<%# Eval("HAP_RefStrA") %>' Width="60px" MaxLength="7" ></asp:TextBox>
                   <img runat="server" id="imgDtStart" src="../Images/date.gif" alt="按此可點選月份" style="border:0;" onclick="ShowCalendar('txtDtStart',true,false)" />                  
                  <asp:Label ID="Label10" runat="server" Text=" 至 " ></asp:Label>
                  <asp:TextBox ID="txtDtEnd" runat="server" Value='<%# Eval("HAP_RefStrB") %>' Width="60px" MaxLength="7"  ></asp:TextBox>
                  <img runat="server" id="imgDtEnd" src="../Images/date.gif" alt="按此可點選月份" style="border:0;" onclick="ShowCalendar('txtDtEnd',true,false)" />
              </td>
            </tr>
            </ContentTemplate>
            </asp:UpdatePanel> 
              
            <tr>
              <td align="right" style="width:25%;height: 25px"><asp:Label ID="Label1" runat="server" Text="資料項："></asp:Label></td>
              <td>
                  <asp:Label ID="Label5" runat="server" Text="姓名:"></asp:Label>
                  <asp:TextBox ID="txtRefStrA" runat="server" Value='<%# Eval("HAP_RefStrA") %>' Width="120px" />
                  <asp:Label ID="Label6" runat="server" Text="身分字號:"></asp:Label>
                  <asp:TextBox ID="txtRefStrB" runat="server" Value='<%# Eval("HAP_RefStrB") %>' Width="120px" />
                   <asp:TextBox ID="TxtRefSTRA1" runat="server" Value='<%# Eval("HAP_RefStrA") %>' Width="200px" Visible="false" />
              </td>
            </tr>
            
            <tr>
              <td align="right" style="width:25%;height: 25px"><asp:Label ID="Label7" runat="server" Text="資料項："></asp:Label></td>
              <td><asp:Label ID="Label9" runat="server" Text="稱謂:"></asp:Label>
                  <asp:DROPDOWNLIST ID="txtRefSTRE" runat="server" Value='<%# Eval("HAP_RefStrE") %>' Width="200px" AutoPostBack="True" />                   
              </td>
            </tr>
            
            <tr>
                  <td align="right" style="height:27px;"><asp:Label ID="Label2" runat="server" Text="資料項："></asp:Label></td>
                  <td style="height: 27px">
                     <asp:Label ID="Label3" runat="server" Text="生日:"></asp:Label>
                     <asp:TextBox ID="txtRefStrC" runat="server" Value='<%# Eval("HAP_RefStrC") %>' Width="88px" />
                     &nbsp &nbsp;
                      <asp:Label ID="Label4" runat="server" Text="生效日:"></asp:Label>&nbsp
                     <asp:TextBox ID="txtRefStrD" runat="server" Value='<%# Eval("HAP_RefStrD") %>' Width="88px" />
                  </td>
            </tr>      
             <tr>
              <td align="right" style="width:25%;height: 25px"><asp:Label ID="Label8" runat="server" Text="處理狀態："></asp:Label></td>
              <td><asp:Label ID="Lbl_STATUS" runat="server" Value='<%# Eval("STATUS") %>' Font-Bold="True" Font-Size="Large" ForeColor="Blue"></asp:Label>
              <asp:TextBox ID="TxtStatus" runat="server" Value='<%# Eval("STATUS") %>' Width="88px" Font-Bold="True" ForeColor="Blue" />
              </td>
            </tr>
    </td>
  </tr>   
 </table>

 


  
  