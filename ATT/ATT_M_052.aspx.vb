'******************************************************************************************************
'* 程式：ATT_M_052 公出資料維護 2/24 改->請假
'* 作成：陳盈志
'******************************************************************************************************
Imports modUnset

Partial Class ATT_M_052
    Inherits System.Web.UI.Page

    Private cUpDateOK As Boolean = False '* DB Ins/Upadte/Del 是否成功

    '******************************************************************************************************
    '* 初始化及權限 設定 
    '******************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            'modDB.InsertSignRecord("AziTest", "M052:Page_Load", My.User.Name) 'AZITEST
            '--------------------------------------------------* 檢查是否有登錄User.Identity.Name
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else '---------------------------------------------* 檢查是否有讀取權限
                ViewState("ROL") = modUtil.GetRolData(Request)
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
            End If

            Me.ViewState("QryField") = Session("QryField") '* 紀錄原查詢條件(供返回時更新)
            Session("QryField") = Nothing
            '
            'modDB.InsertSignRecord("AziTest", "M052:ok-1", My.User.Name) 'AZITEST
            If Session("ATT052MODE") = "1" Then
                'modDB.InsertSignRecord("AziTest", "ATTM052:Mode=1", My.User.Name) '//azitest
                Me.lblTitle.Text = "請假資料 -"
            Else
                Me.lblTitle.Text = "出勤請假 -"
                Me.ViewState("QryField_052B") = Session("QryField_052B") '* 紀錄原查詢條件(供返回時更新)
                Session("QryField_052B") = Nothing
            End If
            '
            'modDB.InsertSignRecord("AziTest", "ATTM052_Page_Load_ATT052MODE=" & Session("ATT052MODE"), My.User.Name) 'AZITEST
            Me.ViewState("ATT052MODE") = Session("ATT052MODE")
            'Session("ATT052MODE") = Nothing
            '
            If Request("formMode") = "add" Then '* 是否有新增權限
                If ViewState("ROL").Substring(1, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                Me.fmvEdit.ChangeMode(FormViewMode.Insert)
                Me.lblTitle.Text = Me.lblTitle.Text + " 新增表單"
            ElseIf Request("formMode") = "edit" Then '* 修改
                Me.fmvEdit.ChangeMode(FormViewMode.Edit)
                Me.lblTitle.Text = Me.lblTitle.Text + " 瀏覽更新表單"
            Else
                Me.fmvEdit.ChangeMode(FormViewMode.ReadOnly)
                Me.lblTitle.Text = Me.lblTitle.Text + " 瀏覽更新表單"
            End If
            '
            Me.lblMsg.Text = Request("msg")
        End If
    End Sub

    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Protected Sub fmvEdit_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles fmvEdit.DataBound
        If Me.fmvEdit.FindControl("ATT_M_052U") IsNot Nothing Then
            CType(Me.fmvEdit.FindControl("ATT_M_052U"), ATT_M_052U).InitScreen(fmvEdit.CurrentMode)

            If Me.fmvEdit.CurrentMode = FormViewMode.ReadOnly Then
                '----------------------* 判斷是否有修改/刪除權限
                If ViewState("ROL").Substring(2, 1) = "N" Then CType(Me.fmvEdit.FindControl("btnEdit"), Button).Enabled = False
                If ViewState("ROL").Substring(3, 1) = "N" Then CType(Me.fmvEdit.FindControl("btnDel"), Button).Enabled = False
                '
                'Dim LCKDT As String = GET_LOCKYM("1", Request.Cookies("STNID").Value) & "20"
                Dim LCKDT As String = GET_LOCKDT(Request.Cookies("STNID").Value)
                Dim VATMDATE As String = Session("VATMDATE").ToString
                If VATMDATE <= LCKDT Then
                    Me.lblMsg.Text = "資料月份已鎖或過帳，不可異動!"
                    CType(Me.fmvEdit.FindControl("btnEdit"), Button).Enabled = False
                    CType(Me.fmvEdit.FindControl("btnDel"), Button).Enabled = False
                End If
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
                'Session("ATT052MODE") = Me.ViewState("ATT052MODE")
                If Request("ATT052MODE") = "1" Then 'Me.ViewState("ATT052MODE")
                    Response.Redirect("ATT_Q_052.aspx")
                Else
                    Session("QryField_052B") = Me.ViewState("QryField_052B")
                    Response.Redirect("ATT_Q_052B.aspx")
                End If
            Case "btnClear"
                Session("QryField") = Me.ViewState("QryField")
                Response.Redirect("ATT_M_052.aspx?FormMode=add")

            Case "btnDel"
                Me.fmvEdit.DeleteItem()
        End Select
    End Sub

    '******************************************************************************************************
    '* DB 新增/修改 前之處理
    '******************************************************************************************************
    Protected Sub dscMain_Updating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) _
              Handles dscMain.Inserting, dscMain.Updating, dscMain.Deleting
        CType(Me.fmvEdit.FindControl("ATT_M_052U"), ATT_M_052U).UpdateDB(fmvEdit.CurrentMode, e) '* 主檔資料填入處理 
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
        With fmvEdit.FindControl("ATT_M_052U")
            Me.lblMsg.Text = "資料" & pMode & "成功！"

            If Request("ATT052MODE") = "1" Then
                If pMode = "新增" Then
                    Session("QryField") = Me.ViewState("QryField")
                    Response.Redirect("ATT_Q_052.aspx?msg=" & Me.lblMsg.Text)
                ElseIf pMode = "刪除" Then
                    Session("QryField") = Me.ViewState("QryField")
                    Response.Redirect("ATT_Q_052.aspx?msg=" & Me.lblMsg.Text)
                ElseIf pMode = "修改" Then
                    Session("QryField") = Me.ViewState("QryField")
                    Response.Redirect("ATT_Q_052.aspx?msg=" & Me.lblMsg.Text)
                End If
            Else
                Session("QryField_052B") = Me.ViewState("QryField_052B")
                If pMode = "新增" Then
                    Session("QryField") = Me.ViewState("QryField")
                    Response.Redirect("ATT_Q_052B.aspx?msg=" & Me.lblMsg.Text)
                ElseIf pMode = "刪除" Then
                    Session("QryField") = Me.ViewState("QryField")
                    Response.Redirect("ATT_Q_052B.aspx?msg=" & Me.lblMsg.Text)
                ElseIf pMode = "修改" Then
                    Session("QryField") = Me.ViewState("QryField")
                    Response.Redirect("ATT_Q_052B.aspx?msg=" & Me.lblMsg.Text)
                End If
            End If
        End With

    End Sub

End Class
