'******************************************************************************************************
'* 程式：ATT_M_061 颱風假設定維護
'* 作成：
'* 版次：2015/10/28(VER1.01)：新開發
'******************************************************************************************************
Imports modUnset

Partial Class ATT_M_061
    Inherits System.Web.UI.Page

    Private cUpDateOK As Boolean = False '* DB Ins/Upadte/Del 是否成功

    '******************************************************************************************************
    '* 初始化及權限 設定 
    '******************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '--------------------------------------------------* 檢查是否有登錄User.Identity.Name
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else '---------------------------------------------* 檢查是否有讀取權限
                ViewState("ROL") = modUtil.GetRolData(Request)
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
            End If

            Me.ViewState("QryField") = Session("QryField") '* 紀錄原查詢條件(供返回時更新)
            Session("QryField") = Nothing

            If Request("formMode") = "add" Then '* 是否有新增權限
                If ViewState("ROL").Substring(1, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                Me.fmvEdit.ChangeMode(FormViewMode.Insert)
                Me.lblTitle.Text = "颱風假設定維護 - 新增表單"
            Else
                Me.fmvEdit.ChangeMode(FormViewMode.ReadOnly)
                Me.lblTitle.Text = "颱風假設定維護 - 瀏覽資料表單"
            End If
        End If
        Me.lblMsg.Text = Request("msg")
    End Sub

    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Protected Sub fmvEdit_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles fmvEdit.DataBound
        'modDB.InsertSignRecord("AziTest", "ATTM059 INIT-1!", My.User.Name) 'AZITEST
        If Me.fmvEdit.FindControl("ATT_M_061U") IsNot Nothing Then
            CType(Me.fmvEdit.FindControl("ATT_M_061U"), ATT_M_061U).InitScreen(fmvEdit.CurrentMode)
            'modDB.InsertSignRecord("AziTest", "ATTM059 INIT-2!", My.User.Name) 'AZITEST
            If Me.fmvEdit.CurrentMode = FormViewMode.ReadOnly Then
                '----------------------* 判斷是否有修改/刪除權限
                If ViewState("ROL").Substring(3, 1) = "N" Then CType(Me.fmvEdit.FindControl("btnDel"), Button).Enabled = False
            End If
        End If
    End Sub

    '******************************************************************************************************
    '* 返回查詢 or 清除輸入畫面
    '******************************************************************************************************
    Protected Sub btnFunc_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case sender.ID
            Case "btnBack"
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_Q_061.aspx")

            Case "btnClear"
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_M_061.aspx?FormMode=add")
            Case "btnDel"
                Me.fmvEdit.DeleteItem()
        End Select
    End Sub

    '******************************************************************************************************
    '* DB 新增/修改 前之處理
    '******************************************************************************************************
    Protected Sub dscMain_Updating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) _
              Handles dscMain.Inserting, dscMain.Updating, dscMain.Deleting
        CType(Me.fmvEdit.FindControl("ATT_M_061U"), ATT_M_061U).UpdateDB(fmvEdit.CurrentMode, e) '* 主檔資料填入處理 
    End Sub

    '******************************************************************************************************
    '* DB 新增/修改 後之處理
    '******************************************************************************************************
    Protected Sub dscMain_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) _
              Handles dscMain.Inserted, dscMain.Updated, dscMain.Deleted
        If e.AffectedRows > 0 Then cUpDateOK = True
    End Sub

    '******************************************************************************************************
    '* FormView ItemInserted /ItemUpdated /ItemDeleted 狀態之處理
    '******************************************************************************************************
    Protected Sub fmvEdit_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewInsertedEventArgs) Handles fmvEdit.ItemInserted
        Call DBUpdated("新增", e)
    End Sub

    Protected Sub fmvEdit_ItemUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewUpdatedEventArgs) Handles fmvEdit.ItemUpdated
        Call DBUpdated("修改", e)
    End Sub

    Protected Sub fmvEdit_ItemDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewDeletedEventArgs) Handles fmvEdit.ItemDeleted
        'modDB.InsertSignRecord("AziTest", "ok-0", My.User.Name)
        Call DBUpdated("刪除", e)
    End Sub

    '******************************************************************************************************
    '* 更新 DB 後之處理(單檔建檔)
    '******************************************************************************************************
    Private Sub DBUpdated(ByVal pMode As String, ByVal e As Object)
        '-----------------* 執行 DB 更新
        If Not cUpDateOK Then
            e.ExceptionHandled = True
            Me.lblMsg.Text = "資料" & pMode & "失敗： "
            If e.Exception IsNot Nothing Then
                If Trim(e.Exception.Message()).IndexOf("PRIMARY KEY") > 0 Then
                    Me.lblMsg.Text = Me.lblMsg.Text & "資料重複建檔！"
                Else
                    Me.lblMsg.Text = Me.lblMsg.Text & e.Exception.Message()
                End If
            End If
            If pMode = "新增" Then
                e.KeepInInsertMode = True
            ElseIf pMode = "修改" Then
                e.KeepInEditMode = True
            End If
            Exit Sub
        End If

        '-----------------------------* 更新成功時，紀錄Log
        With fmvEdit.FindControl("ATT_M_061U")
            Me.lblMsg.Text = "資料" & pMode & "成功！"
            If pMode = "新增" Then
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_Q_061.aspx?msg=" & Me.lblMsg.Text)
            ElseIf pMode = "修改" Then
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_Q_061.aspx?msg=" & Me.lblMsg.Text)
            ElseIf pMode = "刪除" Then
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_Q_061.aspx?msg=" & Me.lblMsg.Text)
            End If
        End With

    End Sub

End Class
