<%@ Page Language="VB" AutoEventWireup="false" Debug="true" CodeFile="ATT_M_055.aspx.vb" Inherits="ATT_M_055"  Theme="ThemeCHG" %>
<%@ Register TagPrefix="ucCHG" TagName="ATT_M_055U"  Src="~/ATT/ATT_M_055U.ascx" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>加油站合理工時維護</title>
    <link href="../StyleSheet.css" rel="stylesheet" type="text/css" /> 
</head>
<body onkeydown="if(event.keyCode==13)event.keyCode=9">
    <form runat="server" id="form1" >
    <div>
     <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True" EnableScriptLocalization="True">
       <Scripts> <asp:ScriptReference Path="~/Script/Util.js" /> </Scripts>   
     </asp:ScriptManager>
     <table class="table" >
      <tr align="center" >
        <td style="height: 10px">
          <asp:Label ID="lblTitle" runat="server" CssClass="titles" Text="功能名稱"></asp:Label></td>
      </tr>
      <tr>
        <td align="left" style="height: 25px; background-color:#fffbd6">
          <asp:Label ID="lblMsg" runat="server" CssClass="msg"></asp:Label><br />
          <asp:FormView ID="fmvEdit" runat="server" DataKeyNames="STNID" DataSourceID="dscMain" >
            <ItemTemplate>
               <ucCHG:ATT_M_055U ID="ATT_M_055U" runat="server"/>
               <table class="table">
                 <tr>
                   <td align="center">
                    <asp:Button ID="btnEdit" runat="server" Text="修改" CommandName="Edit" />
                    <asp:Button ID="btnDel" runat="server" Text="刪除" OnClick="btnFunc_Click" OnClientClick="return confirm('確定要刪除此筆資料嗎?')" />
                    <asp:Button ID="btnPrn" runat="server" Text="列印" OnClientClick="window.print()" />
                    <asp:Button ID="btnBack" runat="server" Text="回主畫面"  OnClick="btnFunc_Click" />                    
                   </td>
                 </tr>
               </table>
            </ItemTemplate>
            
            <EditItemTemplate>
               <ucCHG:ATT_M_055U ID="ATT_M_055U" runat="server"/>
               <table class="table">
                 <tr>
                   <td align="center">
                     <asp:Button ID="btnMod" runat="server" Text="資料送出" CommandName="Update" />
                     <asp:Button ID="btnCancel" runat="server"  Text="取消設定" CommandName="Cancel" />       
                   </td>
                 </tr>
               </table>
            </EditItemTemplate>      
                  
            <InsertItemTemplate>
               <ucCHG:ATT_M_055U ID="ATT_M_055U" runat="server"/>
               <table class="table">
                 <tr>
                   <td align="center">
                    <asp:Button ID="btnIns" runat="server" Text="資料送出" CommandName="Insert" />
                    <asp:Button ID="btnClear" runat="server" Text="清除重來" OnClick="btnFunc_Click" />
                    <asp:Button ID="btnBack" runat="server" Text="回主畫面" OnClick="btnFunc_Click" />                            
                   </td>
                 </tr>
               </table>
            </InsertItemTemplate>
        </asp:FormView>
          
            <asp:SqlDataSource ID="dscMain" runat="server" ConnectionString="<%$ ConnectionStrings:Smile_HQConnectionString %>"
                DeleteCommand="DELETE FROM [DUTY_STNREAL] WHERE [STNID] = @STNID "
                InsertCommand="INSERT INTO [DUTY_STNREAL] ([STNID], [HR11], [HR12], [HR13], [HR14], [HR21], [HR22],[HR23],[HR24],[HR31],[HR32],[HR33],[HR34],[DIFFADD],[DIFFSUB],[CHECKFLAG],[DtCreate],[CreateMan],[DtUpdate],[UpdateMan]) VALUES (@STNID, @HR11, @HR12, @HR13, @HR14, @HR21, @HR22, @HR23,@HR24, @HR31,@HR32,@HR33,@HR34,@DIFFADD,@DIFFSUB,@CHECKFLAG,GETDATE(),@CreateMan,GETDATE(),@UpdateMan)"
                SelectCommand="SELECT * FROM [DUTY_STNREAL] WHERE [STNID] = @STNID "
                UpdateCommand="UPDATE [DUTY_STNREAL] SET [HR11] = @HR11, [HR12] = @HR12, [HR13] = @HR13, [HR14] = @HR14, [HR21] = @HR21, [HR22] = @HR22, [HR23] = @HR23, [HR24] = @HR24, [HR31] = @HR31, [HR32] = @HR32, [HR33] = @HR33, [HR34] = @HR34,[DIFFADD] = @DIFFADD,[DIFFSUB] = @DIFFSUB,[CHECKFLAG] = @CHECKFLAG, [DtUpdate]=GETDATE(),[UpdateMan]=@UpdateMan WHERE [STNID] = @STNID ">
                <DeleteParameters>
                    <asp:Parameter Name="STNID" Type="String" />
                </DeleteParameters>
                <SelectParameters>
                   <asp:QueryStringParameter Name="STNID" QueryStringField="STNID" Type="String" />
                </SelectParameters>
                <UpdateParameters>
                    <asp:Parameter Name="STNID" Type="String" />
                    <asp:Parameter Name="HR11" Type="int32" />
                    <asp:Parameter Name="HR12" Type="int32" />
                    <asp:Parameter Name="HR13" Type="int32" />
                    <asp:Parameter Name="HR14" Type="int32" />
                    <asp:Parameter Name="HR21" Type="int32" />
                    <asp:Parameter Name="HR22" Type="int32" />
                    <asp:Parameter Name="HR23" Type="int32" />
                    <asp:Parameter Name="HR24" Type="int32" />
                    <asp:Parameter Name="HR31" Type="int32" />
                    <asp:Parameter Name="HR32" Type="int32" />
                    <asp:Parameter Name="HR33" Type="int32" />
                    <asp:Parameter Name="HR34" Type="int32" />
                    <asp:Parameter Name="DIFFADD" Type="int32" />
                    <asp:Parameter Name="DIFFSUB" Type="int32" />
                    <asp:Parameter Name="ChECKFLAG" Type="String" />
                    <asp:Parameter Name="UpdateMan" Type="String" />
                </UpdateParameters>
                <InsertParameters>
                    <asp:Parameter Name="STNID" Type="String" />
                    <asp:Parameter Name="HR11" Type="int32" />
                    <asp:Parameter Name="HR12" Type="int32" />
                    <asp:Parameter Name="HR13" Type="int32" />
                    <asp:Parameter Name="HR14" Type="int32" />
                    <asp:Parameter Name="HR21" Type="int32" />
                    <asp:Parameter Name="HR22" Type="int32" />
                    <asp:Parameter Name="HR23" Type="int32" />
                    <asp:Parameter Name="HR24" Type="int32" />
                    <asp:Parameter Name="HR31" Type="int32" />
                    <asp:Parameter Name="HR32" Type="int32" />
                    <asp:Parameter Name="HR33" Type="int32" />
                    <asp:Parameter Name="HR34" Type="int32" />
                    <asp:Parameter Name="DIFFADD" Type="int32" />
                    <asp:Parameter Name="DIFFSUB" Type="int32" />
                    <asp:Parameter Name="CHECKFLAG" Type="String" />
                    <asp:Parameter Name="UpdateMan" Type="String" />
                    <asp:Parameter Name="CreateMan" Type="String" />
                </InsertParameters>
            </asp:SqlDataSource>
        </td>
      </tr>
      </table>
    
    </div>
    
    <div style="width:10px; height:10px;position:absolute; left:-200px; top:-10px;" >  
      <uc:UsrMsgBox ID="UsrMsgBox" runat="server" />
    </div>
    
    </form>
</body>
</html>
