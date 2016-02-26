'******************************************************************************************************
'* 程式：QryCompany 選取公司代號
'* 作成：NEC 杜志揚
'* 版次：2007/12/26(VER1.00)：新開發
'* 版次：2008/02/22(VER1.01)：(NEC杜)加入 ModalPopupExtender，母程式以Call Show()顯示本功能畫面
'******************************************************************************************************

Imports AjaxControlToolkit

Partial Class Common_QryCompany
    Inherits System.Web.UI.UserControl

    Private gCmpID As String = ""
    Private gFName As String = ""
    Private gCmpID_TextChanged As String = ""
    Private gUpdatePanel As String = "" '* For UpdateMode="Conditional" 且不指定Trigger時之更新


    '******************************************************************************************************
    '* 回傳查詢結果，並關閉視窗
    '******************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '--------------------------------------------------* 設定欄位屬性
            Me.txtQCmpID.Attributes.Add("onkeypress", "return CheckKeyAZ09()")

        ElseIf TypeOf sender Is Common_QryCompany Then
            If Me.txtRet.Text.Trim <> "" Then
                Try
                    Dim splitStr As String() = Me.txtRet.Text.Split(",")
                    CType(Me.Parent.FindControl("CustomerNoTextBox"), TextBox).Text = splitStr(0)
                    CType(Me.Parent.FindControl("CustomerNameTextBox"), TextBox).Text = splitStr(1)
                    Me.txtRet.Text = ""

                    'If gCmpID_TextChanged <> "" Then
                    '    If TypeOf Me.Parent Is HtmlForm Then
                    '        'CallByName(Me.Page, gCmpID_TextChanged, CallType.Method, Me, Nothing)
                    '    Else

                    'CallByName(Me.Parent, "Contract_001_CustomerNoTextBox_TextChanged", CallType.Method, Me, Nothing)
                    'End If
                    'End If

                    ''If gUpdatePanel <> "" Then
                    'modUtil.UpdateScreen(Me.Parent, "UpdatePanel1") '* 手動更新
                    'End If

                    Me.popCmp.Hide()

                Catch ex As Exception
                    'MsgBox("grdCmp_SelectedIndexChanged Err:" & ex.Message, MsgBoxStyle.SystemModal Or MsgBoxStyle.Critical)
                    Dim ss As String = ""
                End Try

            Else '----* For 分頁
                Me.dscCmp.SelectCommand = Me.ViewState("SQL")
            End If
        End If
    End Sub

    '******************************************************************************************************
    '* 設定 回傳值Script
    '******************************************************************************************************
    Protected Sub grdCmp_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdCmp.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim sBtn As Button = e.Row.FindControl("btnSel")
            Dim sStr As String
            sStr = "javascript: var sObj = '" & Me.ClientID & "_txtRet'; " _
                 & "$get(sObj).value = '" & e.Row.DataItem.Item(0) & "," & e.Row.DataItem.Item(1) & "' ;"
            sBtn.Attributes.Add("onclick", sStr)
            'DataBinder.Eval(e.Row.DataItem,  "CmpID")
        End If
    End Sub

    '******************************************************************************************************
    '* 開始查詢
    '******************************************************************************************************
    Protected Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim sSql As String
        Me.txtQCmpID.Text = Me.txtQCmpID.Text.Trim
        Me.txtQSName.Text = Me.txtQSName.Text.Trim

        Me.lblMsg.Text = ""
        If Me.txtQCmpID.Text & Me.txtQSName.Text = "" Then
            Me.lblMsg.Text = "至少須輸入一個查詢條件！"
        ElseIf Me.txtQSName.Text = "" And Me.txtQCmpID.Text.Length < 2 Then
            Me.lblMsg.Text = "代號至少須輸入兩碼！"
        End If
        If Me.lblMsg.Text <> "" Then Me.lblMsg.Visible = True : Exit Sub
        Me.lblMsg.Visible = False

        '--------------------------------------------------* 設定SqlDataSource連線及Select命令
        sSql = "SELECT [CustomerNo],[CustomerName] FROM [MMSCustomers] where Effective='Y'"
        If Me.txtQCmpID.Text.Trim <> "" Then
            sSql += " and CustomerNo like '%" + Me.txtQCmpID.Text.Trim + "%'"
        End If
        If Me.txtQSName.Text.Trim <> "" Then
            sSql += " and CustomerName like '%" + Me.txtQSName.Text.Trim + "%'"
        End If
        Me.dscCmp.SelectCommand = sSql

        '--------------------------------------------------* 設定GridView資料來源ID
        Me.grdCmp.DataSourceID = Me.dscCmp.ID
        Me.grdCmp.DataBind()

        If Me.grdCmp.Rows.Count = 0 Then
            Me.lblMsg.Text = "查無資料！"
            Me.lblMsg.Visible = True
        Else
            Me.ViewState("SQL") = sSql
        End If
    End Sub

    ''******************************************************************************************************
    ''* 回傳查詢結果，並關閉視窗
    ''******************************************************************************************************
    'Protected Sub btnSel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        CType(Me.Parent.FindControl(gCmpID), TextBox).Text = Mid(sender.CommandArgument, 1, 6)
    '        CType(Me.Parent.FindControl(gFName), TextBox).Text = Mid(sender.CommandArgument, 7)

    '        If gCmpID_TextChanged <> "" Then
    '            If TypeOf Me.Parent Is HtmlForm Then
    '                CallByName(Me.Page, gCmpID_TextChanged, CallType.Method, Me, Nothing)
    '            Else
    '                CallByName(Me.Parent, gCmpID_TextChanged, CallType.Method, Me, Nothing)
    '            End If
    '        End If
    '        If gUpdatePanel <> "" Then modUtil.UpdateScreen(Me.Parent, gUpdatePanel) '* 手動更新　

    '        'Me.txtQCmpID.Text = "" : Me.txtQFName.Text = ""
    '        Me.popCmp.Hide()

    '    Catch ex As Exception
    '        'MsgBox("grdCmp_SelectedIndexChanged Err:" & ex.Message, MsgBoxStyle.SystemModal Or MsgBoxStyle.Critical)
    '    End Try
    'End Sub

    '******************************************************************************************************
    '* 取消 => 關閉PopUp視窗
    '******************************************************************************************************
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.txtQCmpID.Text = "" : Me.txtQSName.Text = ""
        Me.popCmp.Hide()
    End Sub

    '******************************************************************************************************
    '* 設定/取得 公司代號物件名稱(TextBox)
    '******************************************************************************************************
    Public Property CmpIDObj() As String
        Get
            Return gCmpID
        End Get
        Set(ByVal pValue As String)
            gCmpID = pValue
        End Set
    End Property

    '******************************************************************************************************
    '* 設定/取得 公司名稱物件名稱(TextBox)
    '******************************************************************************************************
    Public Property FNameObj() As String
        Get
            Return gFName
        End Get
        Set(ByVal pValue As String)
            gFName = pValue
        End Set
    End Property

    '******************************************************************************************************
    '* 設定/取得 公司TextChanged 事件名稱 => 有設定時，會啟用該事件
    '******************************************************************************************************
    Public Property CmpID_TextChanged() As String
        Get
            Return gCmpID_TextChanged
        End Get
        Set(ByVal pValue As String)
            gCmpID_TextChanged = pValue
        End Set
    End Property

    '******************************************************************************************************
    '* 設定/取得 UpdatePanel物件名稱
    '******************************************************************************************************
    Public Property UpdatePanelObj() As String
        Get
            Return gUpdatePanel
        End Get
        Set(ByVal pValue As String)
            gUpdatePanel = pValue
        End Set
    End Property

    '******************************************************************************************************
    '* 顯示本功能畫面
    '******************************************************************************************************
    Public Sub Show()
        Me.popCmp.Show()
    End Sub

    Protected Sub btnSel_Click(sender As Object, e As System.EventArgs)

    End Sub
End Class
