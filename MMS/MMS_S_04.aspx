<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MMS_S_04.aspx.vb" Inherits="MMS_MMS_S_04" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
    <style type="text/css">
        .style14
        {
            height: 7px;
        }
        .style20
        {
        }
        .style21
        {
            width: 50px;
        }
        .style22
        {
            width: 88px;
        }
    </style>
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form id="form1" runat="server">
    <div>
          <asp:ScriptManager ID="ScriptManager1" runat="server">
         <Scripts><asp:ScriptReference Path="~/Script/HighLight.js" /></Scripts>
         <Scripts><asp:ScriptReference Path="~/Script/Util.js" /></Scripts>
      </asp:ScriptManager>
      <uc:UsrCalendar ID="UsrCalendar" runat="server" />      
      <div style="width: 10px; height: 10px; position: absolute; left: -200px; top: -10px;">
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
      </div>
          
        <table style="width: 800px;">
            <tr>
                <td align="center" valign="bottom">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="發票開立狀況查詢"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style14">
                    </td>
            </tr>
            <tr>
                <td>
      <asp:UpdatePanel ID="udpGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td align="left">
                      <table style="width: 800px;">
                          <tr>
                              <td class="style21" align="right">
                                  &nbsp;</td>
                              <td align="right" class="style22">
                                  <asp:Label ID="Label1" runat="server" ForeColor="Red" Text="*"></asp:Label>
                                  年度：</td>
                              <td align="left">
                                  &nbsp;<asp:TextBox ID="txtDtFrom0" runat="server" Width="100px" MaxLength="4"></asp:TextBox>
                              </td>
                          </tr>
                          <tr>
                              <td align="right" class="style21">
                                  &nbsp;</td>
                              <td align="right" class="style22">
                                  客戶代號：</td>
                              <td align="left">
                                  &nbsp;<asp:TextBox ID="txtDtFrom1" runat="server" AutoPostBack="True" 
                                      Width="100px" OnTextChanged="txtDtFrom1_TextChanged"></asp:TextBox>
                                  &nbsp;至
                                  <asp:TextBox ID="txtDtFrom2" runat="server" Width="100px"></asp:TextBox>
                              </td>
                          </tr>
                           <tr>
                              <td align="right" class="style21">
                                  &nbsp;</td>
                              <td align="right" class="style22">
                                  發票截止日：</td>
                              <td align="left">
                                  &nbsp;<asp:TextBox ID="txtDtFrom" runat="server" Width="88px" />
                     <img runat="server" id="imgDtFrom" src="../Images/date.gif" alt="按此可點選日期" style="border:0;" onclick="ShowCalendar('txtDtFrom',true,false)" />
                              </td>
                          </tr>
                          <tr>
                              <td align="right" class="style21">
                                  &nbsp;</td>
                              <td align="right" class="style22">
                                  <asp:CheckBox ID="CheckBox2" runat="server" AutoPostBack="True" Text="業務員" />
                                  ：</td>
                              <td>
                                  &nbsp;<asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" 
                                      Enabled="False" Width="113px">
                                  </asp:DropDownList>
                                  &nbsp;&nbsp;&nbsp;<asp:DropDownList ID="DropDownList2" runat="server" AutoPostBack="True" 
                                      Enabled="False" Width="100px">
                                  </asp:DropDownList>
                                  &nbsp;&nbsp;&nbsp;&nbsp;</td>
                          </tr>
                          <tr>
                              <td class="style20" align="center">
                                  &nbsp;</td>
                              <td align="center" class="style22">
                                  &nbsp;</td>
                              <td class="style20">
                                  <asp:Button ID="QueryButton" runat="server" Text="    查    詢    " Height="21px" 
                                      Width="94px" />
                                  <asp:Button ID="ClearButton" runat="server" Text="清除重來" Height="21px" 
                                      Width="94px" />
                                  <asp:Button ID="ExcelButton" runat="server" Enabled="False" Text="EXCEL匯出" 
                                      Height="21px" Width="94px" />
                              </td>
                          </tr>
                      </table>
                   </td>
               </tr>
               <tr>
                  <td align="left">
                      <asp:Label ID="MsgLabel" runat="server" CssClass="msg"></asp:Label>
                      <asp:TextBox ID="txtSql" runat="server" Visible="False" Width="272px"></asp:TextBox>
                   </td>
               </tr>
                <tr>
                    <td align="left">
                        <asp:Panel ID="Panel1" runat="server" ScrollBars="Horizontal" Height="350px" 
                            Width="800px">
                            <asp:GridView ID="grdList" runat="server" />
                        </asp:Panel>
                        <asp:SqlDataSource ID="item_defList" runat="server" 
                            ConnectionString="Data Source=127.0.0.1;Initial Catalog=smile_HQ;Persist Security Info=True;User ID=sa;Password=vispark"></asp:SqlDataSource>
                    </td>
                </tr>
            </table>
         </ContentTemplate>
      </asp:UpdatePanel>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
    
    </div>

    </form>
</body>
</html>
