<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Cust.aspx.vb" Inherits="MMS_Cust" EnableViewState="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
    <style type="text/css">
        .style14
        {
            height: 7px;
        }
        .style15
        {
        }
        .style16
        {
        }
        .style18
        {
            width: 28px;
        }
        .style19
        {
            height: 4px;
        }
        .style22
        {
            width: 560px;
            height: 4px;
        }
        .style23
        {
            height: 4px;
        }
        .style24
        {
            width: 73px;
        }
        .style25
        {
            width: 560px;
        }
        .style26
        {
            width: 60px;
        }
        .style27
        {
            height: 4px;
            width: 50px;
        }
        .style28
        {
            width: 50px;
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
      <div style="width: 10px; height: 10px; position: absolute; left: -200px; top: -10px;">
         <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
      </div>
          
        <table style="width: 800px;">
            <tr>
                <td align="center" valign="bottom">
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="客戶基本資料維護"></asp:Label></td>
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
                              <td align="right" class="style28">
                                  &nbsp;</td>
                              <td align="right" class="style24">
                                  客戶代號:</td>
                              <td>
                                  <asp:TextBox ID="TextBox1" runat="server" Width="121px" AutoPostBack="True" 
                                      MaxLength="15"></asp:TextBox>
                              </td>
                              <td align="center" class="style18">
                                  至</td>
                              <td class="style25">
                                  <asp:TextBox ID="TextBox2" runat="server" Width="121px" MaxLength="15"></asp:TextBox>
                              </td>
                              <td>
                                  &nbsp;</td>
                          </tr>
                          <tr>
                              <td align="right" class="style28">
                                  &nbsp;</td>
                              <td align="right" class="style24">
                                  客戶名稱:</td>
                              <td class="style15" colspan="3">
                                  <asp:TextBox ID="TextBox3" runat="server" Width="286px" MaxLength="150"></asp:TextBox>
                              </td>
                              <td>
                                  &nbsp;</td>
                          </tr>
                          <tr>
                              <td align="right" class="style28">
                                  &nbsp;</td>
                              <td align="right" class="style24">
                                  統一編號:</td>
                              <td>
                                  <asp:TextBox ID="TextBox4" runat="server" Width="121px" MaxLength="8"></asp:TextBox>
                              </td>
                              <td class="style18">
                                  &nbsp;</td>
                              <td class="style25">
                                  &nbsp;</td>
                              <td>
                                  &nbsp;</td>
                          </tr>
                          <tr>
                              <td class="style27">
                                  &nbsp;</td>
                              <td class="style19" colspan="3">
                                  <asp:CheckBox ID="CheckBox1" runat="server" Text="顯示包含已停用之客戶資料" />
                              </td>
                              <td class="style22">
                              </td>
                              <td class="style23">
                              </td>
                          </tr>
                          <tr>
                              <td class="style28">
                                  &nbsp;</td>
                              <td class="style16" colspan="5">
                                  <asp:Button ID="QueryButton" runat="server" Text="    查    詢    " />
                                  <asp:Button ID="ClearButton" runat="server" Text="  清除重來  " />
                                  <asp:Button ID="ExcelButton" runat="server" Enabled="False" Text="EXCEL匯出" />
                                  <img alt="" style="border:0" src="../Images/teacher_icon.gif"
                   />
                                  <asp:LinkButton ID="btnNew" runat="server" style="font-size:11pt" 
                                      Text="新增一筆資料" />
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
                        <asp:GridView ID="grdList" runat="server" />
                        <asp:SqlDataSource ID="item_defList" runat="server"></asp:SqlDataSource>
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
