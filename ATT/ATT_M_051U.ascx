<%@ Control Language="VB" DEBUG="false" AutoEventWireup="false" CodeFile="ATT_M_051U.ascx.vb" Inherits="ATT_M_051U" %>

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
    <td colspan="2" style="width:100%;">
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
              <td align="right" style="height: 25px"><span class="MustInput" >*</span>排班日期：</td>
              <td><asp:TextBox ID="txtSHSTDATE" runat="server" Text='<%# Eval("SHSTDATE") %>'  Width="88px" ></asp:TextBox>
              <asp:TextBox ID="txtSHEDDATE" runat="server" Text='<%# Eval("SHEDDATE") %>'  Width="88px" Visible="False" />
              <asp:TextBox ID="vSHSTDATE" runat="server" Width="88px" Visible="False"> </asp:TextBox>
              </td>
            </tr>                
            
            <tr>
              <td align="right" style="height:25px;">排班選擇：</td>
              <td><asp:DROPDOWNLIST ID="CLSTIME" runat="server" Width="200px" AutoPostBack="True" />  
             
             </td>
            </tr>
            
            <tr>
              <td align="right" style="height:25px;"><span class="MustInput" >*</span>排班時間：</td>
              <td><asp:TextBox ID="txtSHSTTIME" runat="server" Text='<%# Eval("SHSTTIME") %>' MaxLength="4" Width="50px" onblur="CheckTime(this);" /> 結束時間 
                  <asp:TextBox ID="txtSHEDTIME" runat="server" Text='<%# Eval("SHEDTIME") %>' MaxLength="4" Width="50px" onblur="CheckTime(this);" />&nbsp
                  <asp:TextBox ID="vSHSTTIME" runat="server" MaxLength="4" Width="50px" Visible="False"> </asp:TextBox> &nbsp
                  <asp:TextBox ID="vSHEDTIME" runat="server" MaxLength="4" Width="50px" Visible="False"> </asp:TextBox> &nbsp;
             </td>
            </tr>
               
               <tr>
                  <td align="right" style="height:25px;"><span class="MustInput" >*</span>午休時間：</td>
                  <td> <asp:CheckBox ID="chkNormalCls" runat="server" Checked='<%# IIf(eval("RtSTTIME")<>"",true,false) %>' Text="正常班" AutoPostBack="True" OnCheckedChanged="chkNormalCls_CheckedChanged" />
                     <asp:TextBox ID="txtRTSTTIME" runat="server" Text='<%# Eval("RTSTTIME") %>' Width="50px" /> 至 
                     <asp:TextBox ID="txtRTEDTIME" runat="server" Text='<%# Eval("RTEDTIME") %>' Width="50px" /> 時分
                  </td>
               </tr>
               
               <tr>
                  <td align="right" style="height:25px;">出勤時數：</td>
                  <td> <asp:TextBox ID="txtWORKHOUR" runat="server" Text='<%# Eval("WORKHOUR") %>' Width="50px" /> 小時 
                  </td>
              </tr>
         </table>
    </td>
  </tr>   
 </table>
 
 <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" > 
    <uc:QryEmp id="QryHrEmp" runat="server" EmpIDObj="txtEmpID" EmpNameObj="txtEmpName" UpdatePanelObj="udpEmp" />
 </div>


  
  