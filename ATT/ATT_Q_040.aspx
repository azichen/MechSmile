<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ATT_Q_040.aspx.vb" Inherits="ATT_ATT_Q_040"  Theme="ThemeCHG"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<!-- 版次：2014/01/24(VER1.00)：新開發 -->

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>加油站天氣預報查詢</title>
    <link href="../StyleSheet.css"rel="stylesheet" type="text/css" />    
                              <script type="text/javascript">
    //--------------------------------* 
    function setValue(obj1,obj2) {
     o1 =$get(obj1).value;
     o2=$get(obj2).value;
     if (o2 == "") {
        $get(obj2).value = o1;
        }
    }
          
 
          </script> 

    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
</head>
<body  onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form id="form1" runat="server">
        <div style="text-align: center">
            <div style="text-align: center">
                <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True">
                    <Scripts>
                        <asp:ScriptReference Path="~/Script/Util.js" />
                    </Scripts>
                </asp:ScriptManager>
                <table class="table">
                    <tr>
                        <td colspan="2" align="center">
                            <asp:Label ID="TitleLabel" runat="server" CssClass="titles" Text="功能名稱"></asp:Label></td>
                    </tr>
                    <tr>

                                    <td style="width: 15%" align="right">
                            區別：
                                    </td>
                                    <td style="width: 85%">
                                        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
                                            <ContentTemplate>
                                        <asp:TextBox ID="txtAreaFrom" runat="server" AutoPostBack="true" MaxLength="5" Width="64px"></asp:TextBox>
                                                <asp:TextBox ID="AreaName" runat="server" CssClass="readonly" ReadOnly="true" Width="100px"> </asp:TextBox>
                                                <asp:Button ID="btnArea1" runat="server" Text="選擇區別" /> 至
                            <asp:TextBox ID="txtAreaTo" runat="server" AutoPostBack="true" MaxLength="5" Width="64px"></asp:TextBox>
                                                <asp:TextBox ID="AreaName2" runat="server" CssClass="readonly" ReadOnly="true" Width="100px"> </asp:TextBox>
                                                <asp:Button ID="btnArea2" runat="server" Text="選擇區別" />
                                            </ContentTemplate>
                                        </asp:UpdatePanel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right" valign="top">
                                        加油站：</td>
                                    <td>
                                    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                                    <ContentTemplate>
                                       <asp:TextBox ID="txtStnFrom" runat="server" MaxLength="6" Width="64px" AutoPostBack="true"></asp:TextBox>
                                        <asp:TextBox   ID="txtStnName" runat="server" Width="100px" CssClass="readonly" ReadOnly="true" /> 
                                        <asp:Button ID="btnStnId" runat="server" Text="選擇站別" />
                                        至
                                        <asp:TextBox ID="txtStnTo" runat="server" AutoPostBack="true" MaxLength="6" Width="64px"></asp:TextBox>
                                        <asp:TextBox ID="txtStnName2" runat="server" CssClass="readonly" ReadOnly="true"
                                            Width="100px">
                                        </asp:TextBox>
                                        <asp:Button ID="btnStnId2" runat="server" Text="選擇站別" />
                                    </ContentTemplate>
                                </asp:UpdatePanel>
                                        
                                     </td>
                                </tr>
                                <tr>
                                    <td align="right">
                                        <span class="MustInput" >*</span>日期：</td>
                                    <td>
                                        <asp:TextBox ID="DateSTextBox" runat="server" Columns="10" CssClass="readonlyDate"></asp:TextBox>
                                       <img runat="server" id="imgDtFrom" src="../Images/date.gif" alt="按此可點選日期" style="border:0;" />
                                        ~ <asp:TextBox ID="DateETextBox" runat="server" Columns="10" CssClass="readonlyDate"></asp:TextBox>
                                        <img runat="server" id="imgDtTo" src="../Images/date.gif" alt="按此可點選日期" style="border:0;" />
                                        <asp:TextBox ID="txtDtMax" runat="server" style="display:none" />
                                        
                                        </td>
                                </tr>
                                <tr>
                                    <td align="right">
                                        分類：</td>
                                    <td>
                                        &nbsp;
                                        <asp:CheckBoxList ID="WetCodeChk" runat="server" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">晴</asp:ListItem>
                                            <asp:ListItem Value="2">陰</asp:ListItem>
                                            <asp:ListItem Value="3">雨</asp:ListItem>
                                        </asp:CheckBoxList></td>
                                </tr>
                                <tr><asp:HiddenField ID="txtToDay" runat="server" /></tr>
                                <tr>
                                    <td align="right">
                                    </td>
                                    <td>
                            
                            <asp:Button ID="QueryButton" runat="server" Text="開始查詢" /><asp:Button ID="ClearButton" runat="server"
                                Text="清除重來" /><asp:Button ID="ExcelButton" runat="server" Text="匯出EXCEL" Enabled="False" /></td>
                             
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Label ID="MsgLabel" runat="server" CssClass="msg"></asp:Label></td>
                    </tr>
                    <tr>
                        <td align="left" colspan="2">
                            <asp:GridView ID="grdList" runat="server">
                            </asp:GridView>
                        <!--GridView開始 -->
                        
                        <!--Gridview結束-->
                            <asp:SqlDataSource ID="item_defList" runat="server"></asp:SqlDataSource>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
 
<div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" >
       <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
       </div>
       <uc:UsrCalendar ID="UsrCalendar" runat="server" /> 
                      <uc:QryStn ID="QryStn" runat="server" StnIDObj="txtStnFrom" StnNameObj="txtStnName"  StnID_TextChanged="txtStnFrom_TextChanged"/>                  
        <uc:QryStn ID="QryStn2" runat="server" StnIDObj="txtStnTo" StnNameObj="txtStnName2"/>                  
        <uc:QryArea ID="QryArea1" runat="server"  AreID_TextChanged="txtAreaFrom_TextChanged"   AreIDObj="txtAreaFrom"  AreDescObj="AreaName" />
        <uc:QryArea ID="QryArea2" runat="server"   AreIDObj="txtAreaTo"  AreDescObj="AreaName2" />

    </form>
</body>
</html>

