<%@ Page Language="VB" AutoEventWireup="True" CodeFile="VoucherTr.aspx.vb" Inherits="VoucherTr" EnableEventValidation = "false" Theme="ThemeCHG" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>會計傳票匯入作業</title>
<link href="../StyleSheet.css" rel="stylesheet" type="text/css" />  
</head>
<body>
    <form id="form1" runat="server">
       <asp:ScriptManager ID="ScriptManager1" runat="server">
          <Scripts> <asp:ScriptReference Path="~/Script/HighLight.js" /> </Scripts>       
       </asp:ScriptManager>
        
       <% '*----------- 用 UpLoad 元件之區域，就不可使用 UpdatePanel，否則無法取得上傳檔案 ----------- %> 
        <table class="table">
            <tr style=" height:5px">
                <td align="right" style="width:15%"> </td>
                <td style="width:85%" > </td>
            </tr>        
            <tr>
                <td align="center" style="height: 21px" colspan="2">
                    <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="會計傳票匯入作業"></asp:Label></td>
            </tr>
            <tr style=" height:5px">
                <td > </td> <td > </td>
            </tr>             
            <tr>
                <td align="right" style="height: 21px;"><span class="MustInput" >*</span>檔案位置：</td>
                <td><asp:FileUpload ID="fileUp" runat="server"  />              
                </td>
            </tr>  
          <tr>
            <td style="height: 21px">&nbsp</td>
            <td>
               <asp:Button ID="btnImport" runat="server" Text="預覽"  /><asp:Button 
                           ID="btnClear" runat="server" Text="清除重來" />
                <asp:Button 
                           ID="btnSave" runat="server" Text="匯入" Enabled="false" /></td>
            </tr>
          </table>
            
            
      <uc:UpdatePanelFix ID="updGrid" runat="server" RenderMode="Inline" UpdateMode="Conditional">
        <ContentTemplate>
          <table class="table">
          <tr>
             <td colspan="2" align="left">
                 <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label></td>
          </tr>           
         </table>
         
         <asp:Panel ID="pnlQry" runat="server" CssClass="fixedHeader" Width="800px" Height="470px" ScrollBars="Vertical">           
         <asp:GridView ID="grdList" runat="server" AutoGenerateColumns="False"  >
           <EmptyDataTemplate> </EmptyDataTemplate>
             <Columns>
                 <asp:TemplateField HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-ForeColor="white" >
                   <ItemTemplate><%#Container.DataItemIndex + 1%></ItemTemplate>
                 </asp:TemplateField>
             </Columns>
         </asp:GridView>
         </asp:Panel>
       </ContentTemplate>
     </uc:UpdatePanelFix>
       
       <asp:SqlDataSource ID="dscList" runat="server" ConnectionString="<%$ ConnectionStrings:Smile_HQConnectionString %>"></asp:SqlDataSource> 
                   
        <uc:QryCmp id="QryCmp" runat="server" CmpIDObj="txtCmpID" FNameObj="txtSName" UpdatePanelObj="updQry" />
        
        <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
   
    </form>
</body>
</html>
