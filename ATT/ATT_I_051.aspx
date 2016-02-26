<%@ Page Language="VB" AutoEventWireup="false" DEBUG="false" CodeFile="ATT_I_051.aspx.vb" Inherits="ATT_I_051" Theme="ThemeCHG" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>排班資料維護</title>
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" /> 
</head>

<body onkeydown="if(event.keyCode==13)event.keyCode=9">

 <form runat="server" id="form1" >
  
 <asp:ScriptManager ID="ScriptManager1" runat="server">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
 </asp:ScriptManager>
      <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;">
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
      </div>
      
  <uc:UsrCalendar ID="UsrCalendar" runat="server" /> 
  <div> 
  <table width="800px">
         <tr>
               <td valign="middle" style="height: 25px; text-align: center;background-color:#1c5e55" align="center">
               <asp:Label ID="TitleLabel" runat="server" CssClass="titles2" Text="功能名稱"></asp:Label></td>
         </tr>
  
         <tr>
               <!--  <td style="background-color:#f7f7f7" align="center" >  -->
               <asp:Label ID="msgLabel2" runat="server" CssClass="msg"></asp:Label>
               <uc:UpdatePanelFix ID="updQry" runat="server" RenderMode="Inline" UpdateMode="Conditional">
                   <ContentTemplate> 
                        <table style="width: 100%">
                            <tr>
                                <td align="right" style="width: 15%">
                                        加油站：</td>
                                <td style="width: 50%" align="left">
                                            <asp:TextBox ID="txtSTNID_i" runat="server" MaxLength="6" Width="56px" Columns="6" AutoPostBack="true"
                                                 OnTextChanged="txtSTNID_i_TextChanged"
                                          /><asp:TextBox ID="txtSTNNAME_i" runat="server" Width="140px" AutoPostBack="true" />
                                </td>
                            </tr>
                            
                            <tr>
                                <td align="right" style="width: 15%">
                                        員工代號：</td>
                                <td style="width: 50%" align="left">
                                <asp:UpdatePanel ID="UpdatePanel_EMP" runat="server">
                                        <ContentTemplate>
                                            <asp:TextBox ID="txtEMPID_i" runat="server" MaxLength="8" Width="78px" Columns="8" AutoPostBack="true"
                                                 OnTextChanged="txtEMPID_i_TextChanged"
                                          /><asp:TextBox ID="txtEMPNAME_i" runat="server" Width="96px" AutoPostBack="true" />
                                            <asp:Button ID="btnEmp" runat="server" Text="選擇員工" OnClick="btnEmp_Click" />
                                            <uc:QryEmp id="QryEmp" runat="server" EmpIDObj="txtEMPID_i" EmpNameObj="txtEMPNAME_i" />
                                        </ContentTemplate>
                                     </asp:UpdatePanel>
                                </td>
                            </tr>
                            
                            <tr>
                                <td align="right" style="height:25px;"> 排班日期：</td>
                                <td>
                                <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                        <ContentTemplate>
                                         <asp:Button ID="btnPreDate" runat="server" Width="80px" Text="前一天" OnClick="btnPreDate_Click" />&nbsp;&nbsp;
                                         <asp:TextBox ID="txtSchDt" runat="server" Width="88px" AutoPostBack="True" OnTextChanged="txtSchDt_TextChanged" />
                                         <img runat="server" id="imgSchDt" src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtSchDt',true,false)" />
                                         &nbsp;&nbsp;<asp:Button ID="btnPosDate" runat="server" Width="80px" Text="後一天" OnClick="btnPosDate_Click" />
                                         &nbsp;&nbsp;<asp:Label ID="DUTYHOUR" runat="server" Width="200px" Text="" />
                                       </ContentTemplate>
                                     </asp:UpdatePanel>
                                </td>
                             </tr>
                             
                             <tr>   
                                <td align="right" style="width: 15%"> 排班代碼：</td>
                                
                                <td align="LEFT" style="width: 50%">
                                <asp:UpdatePanel ID="UpdatePanel3" runat="server">
                                        <ContentTemplate>
                                            <asp:DropDownList ID="CLSTIME" runat="server" AutoPostBack="True"  OnTextChanged="CLSTIME_i_TextChanged"></asp:DropDownList> 
                                            </ContentTemplate>
                                     </asp:UpdatePanel>
                                </td>
                            </tr>
                            
                            <tr> </tr>
                         </table>
               
                   </ContentTemplate>
               </uc:UpdatePanelFix>
            
               
               
               <asp:FormView ID="FormView1" runat="server" DataSourceID="SqlDataSource1" Width="100%">
                    <InsertItemTemplate>
                        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                           <ContentTemplate>
                                <table style="width: 100%">
                                    <tr>
                                        <!-- <td align="center" style="color: black; height: 40px; background-color: #ffcccc">
                                            排班日期</td> -->
                                        <td align="center" style="color: black; height: 40px; background-color: #ffcccc">
                                            起始時間</td>
                                        <td align="center" style="color: black; height: 40px; background-color: #ffcccc">
                                            結束時間</td>
                                        <td align="center" style="color: black; height: 40px; background-color: #ffcccc">
                                            午休起始</td>
                                        <td align="center" style="color: black; height: 40px; background-color: #ffcccc">
                                            午休結束</td>
                                    </tr>
                 
                                    <tr>
                                        <!-- <td style="background-color: #ffffcc" align="center">
                                                <asp:TextBox ID="txtSHSTDATE_i" runat="server" ></asp:TextBox></td>       -->
                                        <td style="background-color: #ffffcc" align="center">
                                            <asp:TextBox ID="txtSHSTTIME_i" runat="server" Columns="8" MaxLength="8"></asp:TextBox></td>
                                        <td style="background-color: #ffffcc" align="center">
                                            <asp:TextBox ID="txtSHEDTIME_i" runat="server" Columns="8" MaxLength="8"></asp:TextBox></td>
                                        <td style="background-color: #ffffcc" align="center">
                                            <asp:TextBox ID="txtRTSTTIME_i" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>    
                                        <td style="background-color: #ffffcc" align="center">
                                            <asp:TextBox ID="txtRTEDTIME_i" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>    
                                    </tr>
                                </table>
            
                                <table style="width: 100%">
                                    <tr>
                                        <td align="left" colspan="12" style="height: 2px">
                                            <asp:Button ID="b1" runat="server" Text="新增一筆" OnClick="b1_Click"/>
                                            <asp:Button ID="btnClear" runat="server" OnClick="btnClear_Click" Text="清除重設" /></td>
                                   </tr>
                                    <tr>
                                        <td align="left" colspan="12" style="height: 15px">
                                            <asp:Label ID="msgLabel" runat="server" CssClass="msg"></asp:Label></td>
                                    </tr>
                                 </table>
                                <asp:GridView ID="GridView1" runat="server" OnRowDeleting="GridView1_RowDeleting" OnRowEditing="GridView1_RowEditing" CellSpacing="2">
                                </asp:GridView>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                        
                        <asp:Button ID="SendButton" runat="server" OnClick="SendButton_Click" Text="資料送出" />
                        <asp:Button ID="ClearButton" runat="server" CausesValidation="False" OnClick="ClearButton_Click"
                                Text="清除重來" UseSubmitBehavior="False" />
                        <asp:Button ID="Button_return" runat="server" OnClick="Button_return_Click" Text="回上一頁" />
                           <uc:QryStn ID="QryStn" runat="server" StnIDObj="txtSTNID_i" StnNameObj="txtSTNNAME_i" />
                           
                    </InsertItemTemplate>
                    
                </asp:FormView>  
                
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:MECHPAConnectionString %>"
                 SelectCommand="SELECT S.StnID,M.StnName,S.EMPID,EMPL_NAME,SHSTDATE,SHSTTIME,SHEDDATE,SHEDTIME,RTSTTIME,RTEDTIME,WORKHOUR FROM SCHEDM S INNER JOIN MECHSTNM M ON S.STNID=M.STNID INNER JOIN [MP_HR].DBO.EMPLOYEE E ON S.EMPID=E.EMPL_ID WHERE STNID='441008' and empid='09112801' and shstdate>='20130721'"
                 InsertCommand="INSERT INTO [SCHEDM] ([STNID], [EMPID], [SHSTDATE], [SHSTTIME], [SHEDDATE], [RTSTTIME], [RTEDTIME], [WORKHOUR], [FACFLAG], [FACWH], [NITEWH], [INUSER], [INDATE], [INTIME]) VALUES ( @STNID, @EMPID, @SHSTDATE, @SHSTTIME, @SHEDDATE, @SHEDTIME, @RTSTTIME, @RTEDTIME,@WORKHOUR,'N',0,0, @INUSER , @INDATE, @INTIME)" >
                            <SelectParameters>
                                <asp:QueryStringParameter Name="STNID" QueryStringField="STNID" Type="String" />
                                <asp:QueryStringParameter Name="EMPID" QueryStringField="EMPID" Type="String" />
                                <asp:QueryStringParameter Name="SHDATEST" QueryStringField="SHDATEST" Type="String" />
                                <asp:QueryStringParameter Name="SHDATEEN" QueryStringField="SHDATEEN" Type="String" />
                            </SelectParameters>
                            <InsertParameters>
                                <asp:Parameter Name="STNID" Type="String" />
                                <asp:Parameter Name="EMPID" Type="String"/>
                                <asp:Parameter Name="SHSTDATE" Type="String"/>
                                <asp:Parameter Name="SHSTTIME" Type="String"/>
                                <asp:Parameter Name="SHEDDATE" Type="String"/>
                                <asp:Parameter Name="SHEDTIME" Type="String"/>
                                <asp:Parameter Name="RTSTTIME" Type="String"/>
                                <asp:Parameter Name="RTEDTIME" Type="String"/>
                                <asp:Parameter Name="WORKHOUR" Type="DOUBLE"/>
                                <asp:Parameter Name="INUSER" Type="String" />
                                <asp:Parameter Name="INDATE" Type="String" />
                                <asp:Parameter Name="INTIME" Type="String" />
                            </InsertParameters>
                </asp:SqlDataSource>
    
           </tr>
     </table>
    </div> 
  </form>
</body>
</html>
