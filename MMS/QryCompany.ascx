<%@ Control Language="VB" AutoEventWireup="false" CodeFile="QryCompany.ascx.vb" Inherits="Common_QryCompany" %>
   <link href="../StyleSheet.css" rel="stylesheet" type="text/css" />
 
 <asp:TextBox ID="txtTarget" runat="server" style=" display:none" />
 <asp:Panel ID="pnlCmp" runat="server" SkinID="selPanel" Width="402px" 
    BackColor="White"> 
   <div class="selTitle" style="width:400px" >選擇客戶</div>
   <br />
   <asp:UpdatePanel ID="updQryCmp" runat="server" RenderMode="Inline" UpdateMode="Conditional">
       <ContentTemplate>
                代號：<asp:TextBox ID="txtQCmpID" runat="server" MaxLength="6" Width="50px"></asp:TextBox>
                &nbsp;名稱：<asp:TextBox ID="txtQSName" runat="server" Width="100px"></asp:TextBox>
        <asp:Button ID="btnSearch" runat="server" Text="查詢" OnClick="btnSearch_Click" style="width:50px" /><asp:Button 
                         ID="btnCancel" runat="server" Text="取消" style="width:50px" />
       <br /><asp:Label ID="lblMsg" runat="server" CssClass="msg" Visible="false"></asp:Label>

       <asp:Panel ID="pnlSelect" runat="server"  CssClass="fixedHeader"  Height="475px" >
        <asp:TextBox ID="txtRet" runat="server" Width="20px" style="display:none" />
        <asp:GridView ID="grdCmp" runat="server" AutoGenerateColumns="False" Width="368px" 
                      SkinID="selGrid" DataKeyNames="CustomerNo"  AllowPaging="True" PageSize="15"
                      PagerSettings-Mode="NextPreviousFirstLast" 
                      PagerSettings-FirstPageImageUrl="~/Images/First.gif" 
                      PagerSettings-PreviousPageImageUrl="~/Images/Previous.gif"
                      PagerSettings-NextPageImageUrl="~/Images/Next.gif" 
                      PagerSettings-LastPageImageUrl="~/Images/Last.gif" 
               EnableModelValidation="True"   > 
            <Columns>               
                 <asp:TemplateField HeaderText="選取" ItemStyle-HorizontalAlign="Center" >
                   <ItemTemplate>
                      <asp:Button runat="server" ID="btnSel" Text="選取" UseSubmitBehavior="false" 
                           style="width:45px" onclick="btnSel_Click" />
                   </ItemTemplate>
                   <ItemStyle Width="50px" />
                 </asp:TemplateField>
                
                <asp:BoundField DataField="CustomerNo" HeaderText="客戶編號" >
                    <ItemStyle Width="100px" />
                </asp:BoundField>
                <asp:BoundField DataField="CustomerName" HeaderText="客戶名稱" >
                    <ItemStyle Width="200px" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="Teal" ForeColor="White" />
            <PagerSettings FirstPageImageUrl="~/Images/First.gif" 
                LastPageImageUrl="~/Images/Last.gif" Mode="NextPreviousFirstLast" 
                NextPageImageUrl="~/Images/Next.gif" 
                PreviousPageImageUrl="~/Images/Previous.gif" />
        </asp:GridView>
        </asp:Panel>
       </ContentTemplate>
       <Triggers>
           <asp:AsyncPostBackTrigger ControlID="grdCmp" EventName="PageIndexChanging" />
       </Triggers>
    </asp:UpdatePanel>
        
    <ajaxToolkit:ModalPopupExtender ID="popCmp" runat="server" BackgroundCssClass="modalBackground"
               PopupControlID="pnlCmp" TargetControlID="txtTarget" />
        
    <asp:SqlDataSource ID="dscCmp" runat="server" DataSourceMode="DataSet"
             ConnectionString="<%$ ConnectionStrings:smile_HQConnectionString %>"  EnableViewState="False">
    </asp:SqlDataSource>
               
     
 </asp:Panel>
