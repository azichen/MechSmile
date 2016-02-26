<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Area.aspx.vb" Inherits="MMS_Areas" Theme="ThemeCHG"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />    
       <script type="text/javascript">
           //--------------------------------* 
           function setValue(obj1, obj2) {
               o1 = $get(obj1).value;
               o2 = $get(obj2).value;
               if (o2 == "") {
                   $get(obj2).value = o1;
               }
           }
          
    </script>
    <style type="text/css">
        .style1
        {
        }
        .style3
        {
        }
        .style4
        {
            width: 91px;
        }
        .style5
        {
            width: 50px;
        }
        .style6
        {
            height: 25px;
            width: 50px
            ;
        }
    </style>
</head>
<body  onkeydown="if(event.keyCode==13)event.keyCode=9">
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
                     <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="區域資料維護"></asp:Label></td>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
      <asp:UpdatePanel ID="udpGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
         <ContentTemplate>
            <table class="table">
               <tr>
                  <td style="height:25px;" colspan="2">
                      <table style="width: 100%;">
                          <tr>
                              <td class="style5">
                                  &nbsp;</td>
                              <td class="style4">
                                  區域代號起:</td>
                              <td class="style3" valign="middle">
                                  <asp:TextBox ID="TextBox1" runat="server" AutoPostBack="true" MaxLength="10" ></asp:TextBox>
                                  &nbsp;至 
                                  <asp:TextBox ID="TextBox2" runat="server" MaxLength="10"></asp:TextBox>
                              </td>
                              <td>
                                  &nbsp;</td>
                          </tr>
                          <tr>
                              <td class="style5">
                                  &nbsp;</td>
                              <td class="style1" colspan="3">
                                  <asp:CheckBox ID="CheckBox1" runat="server" Text="顯示包含已停用資料" />
                              </td>
                          </tr>
                          <tr>
                              <td class="style5">
                                  &nbsp;</td>
                              <td class="style1" colspan="3">
                                  <asp:Button ID="QueryButton" runat="server" Text="查    詢" />
                                  <asp:Button ID="ClearButton" runat="server" Text="清除重來" />
                                  <asp:Button ID="ExcelButton" runat="server" Enabled="False" Text="EXCEL匯出" />
                                  <img alt="" style="border:0" src="../Images/teacher_icon.gif"
                   />
                                  <asp:LinkButton ID="btnNew" runat="server" style="font-size: 11pt" 
                                      Text="新增一筆資料" />
                                  <asp:TextBox ID="txtSql" runat="server" Visible="False" Width="272px"></asp:TextBox>
                              </td>
                          </tr>
                      </table>
                   </td>
               </tr>
                <tr>
                    <td>
                    </td>
                    <td style="width: 85%">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td colspan="2" style="height: 25px;">
                        <asp:Label ID="MsgLabel" runat="server" CssClass="msg"></asp:Label>
                        <asp:GridView ID="grdList" runat="server" />
                        <asp:SqlDataSource ID="item_defList" runat="server"></asp:SqlDataSource>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="height: 25px;">
                        &nbsp;</td>
                </tr>
               <tr>
                  <td colspan="2" style="height: 25px;">
                      &nbsp;</td>
               </tr>
               <tr>
                  <td align="left" colspan="2">
                      &nbsp;</td>
               </tr>
               <tr>
                  <td align="left" colspan="2">
                      &nbsp;</td>
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
