'******************************************************************************************************
'* 程式：ATT_M_055 加油站合理工時維護
'* 作成：
'* 版次：2013/12/31(VER1.01)：新開發
'******************************************************************************************************

Partial Class ATT_M_055

    Inherits System.Web.UI.Page

    Private cUpDateOK As Boolean
    '******************************************************************************************************
    '* 初始化設定
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
            Me.ViewState("iCurPage") = Session("iCurPage") '* 紀錄原瀏覽頁次(供返回時更新)
            Session("QryField") = Nothing
            Session("iCurPage") = Nothing

            If Request("formMode") = "add" Then '* 是否有新增權限
                If ViewState("ROL").Substring(1, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                Me.fmvEdit.ChangeMode(FormViewMode.Insert)
                Me.lblTitle.Text = "加油站合理工時維護 - 新增表單"
            Else
                Me.fmvEdit.ChangeMode(FormViewMode.ReadOnly)
                Me.lblTitle.Text = "加油站合理工時維護 - 瀏覽更新表單"
            End If
        End If
        Me.lblMsg.Text = ""
    End Sub

    '******************************************************************************************************
    '* FormView_DataBound => 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Protected Sub fmvEdit_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles fmvEdit.DataBound
        If Me.fmvEdit.FindControl("ATT_M_055U") IsNot Nothing Then
            CType(Me.fmvEdit.FindControl("ATT_M_055U"), ATT_M_055U).InitScreen(fmvEdit.CurrentMode)
            If Me.fmvEdit.CurrentMode = FormViewMode.ReadOnly Then
                '----------------------* 判斷是否有修改 /刪除權限
                If ViewState("ROL").Substring(2, 1) = "N" Then CType(Me.fmvEdit.FindControl("btnEdit"), Button).Enabled = False
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
                Session("iCurPage") = Me.ViewState("iCurPage")
                Response.Redirect("ATT_Q_055.aspx")
            Case "btnClear"
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_M_055.aspx?FormMode=add")
            Case "btnDel" '--------------------* 檢核 活動是否仍在進行中
                Me.fmvEdit.DeleteItem()
        End Select
    End Sub

    '******************************************************************************************************
    '* DB 新增/修改 前之處理
    '******************************************************************************************************
    Protected Sub dscMain_Updating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) _
              Handles dscMain.Inserting, dscMain.Updating, dscMain.Deleting
        CType(Me.fmvEdit.FindControl("ATT_M_055U"), ATT_M_055U).UpdateDB(fmvEdit.CurrentMode, e) '* 主檔資料填入處理 
        'End If
    End Sub

    '******************************************************************************************************
    '* DB 新增/修改 後之處理
    '******************************************************************************************************
    Protected Sub dscMain_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) _
              Handles dscMain.Inserted, dscMain.Updated, dscMain.Deleted
        If e.AffectedRows > 0 Then
            e.Command.Transaction.Commit()
            cUpDateOK = True
        Else
            e.Command.Transaction.Rollback() '* 失敗時須RollBack
        End If
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
        Call DBUpdated("刪除", e)
    End Sub

    Private Sub DBUpdated(ByVal pMode As String, ByVal e As Object)
        If Not cUpDateOK Then

            e.ExceptionHandled = True
            Me.lblMsg.Text = "資料" & pMode & "失敗： "
            If e.Exception IsNot Nothing Then Me.lblMsg.Text = Me.lblMsg.Text & e.Exception.Message()
            If pMode = "新增" Then
                e.KeepInInsertMode = True
            ElseIf pMode = "修改" Then
                e.KeepInEditMode = True
            End If
            Exit Sub
        End If

        '-----------------------------* 更新成功時，紀錄Log
        Dim sStr As String

        With fmvEdit.FindControl("ATT_M_055U")
            sStr = "STNID:" & CType(.FindControl("txtSTNID"), TextBox).Text
            modDB.InsertSignRecord("加油站合理工時維護" & pMode, sStr, My.User.Name)
            Me.lblMsg.Text = "資料" & pMode & "成功！"
            If pMode = "新增" Then
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_Q_055.aspx?msg=" & Me.lblMsg.Text)
            ElseIf pMode = "刪除" Then
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_Q_055.aspx?msg=" & Me.lblMsg.Text)
            End If
        End With

    End Sub

End Class
