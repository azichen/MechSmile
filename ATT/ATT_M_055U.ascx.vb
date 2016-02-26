'******************************************************************************************************
'* 程式：ATT_M_055U 加油站合理工時設定維護
'* 版次：2013/12/31(VER1.01)：新開發
'******************************************************************************************************

Imports modUtil

Partial Class ATT_M_055U

    Inherits System.Web.UI.UserControl

    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Public Sub InitScreen(ByVal pMode As FormViewMode)
        Dim sSql As String
        Me.ViewState("FormViewMode") = pMode

        '-------------------------------------------* 設定 ReadOnly
        btnStnID.Visible = False
        If pMode = FormViewMode.ReadOnly Then
            SetAllReadOnly(Me)
            sSql = "Select STNNAME From MECHSTNM Where STNID = '" & Me.txtStnID.Text & "'"
            Me.txtStnName.Text = modDB.GetCodeName("STNNAME", sSql)
            '
            'modDB.InsertSignRecord("AziTest", "是否檢查-1:" & Me.LABCHECKFLAG.Text & "=" & sSql, My.User.Name) 'AZITEST
            If LABCHECKFLAG.Text = "Y" Then
                CKBCHECKFLAG.Checked = True
            Else
                CKBCHECKFLAG.Checked = False
            End If
        Else
            '-----------------------------------------* 設定欄位屬性
            Me.txtHR11.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR12.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR13.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR14.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR21.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR22.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR23.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR24.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR31.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR32.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR33.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtHR34.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtDiffADD.Attributes.Add("onkeypress", "return CheckKeyNumber()")
            Me.txtDiffSUB.Attributes.Add("onkeypress", "return CheckKeyNumber()")

            SetObjReadOnly(Me, "txtStnNAME")
            If pMode = FormViewMode.Insert Then
                Me.CKBCHECKFLAG.Checked = False
                btnStnID.Visible = True
                Me.txtStnID.Focus()
            Else '-------------------------------------* 顯示名稱
                'Call ShowIdName()
                sSql = "Select STNNAME From MECHSTNM Where STNID = '" & Me.txtStnID.Text & "'"
                Me.txtStnName.Text = modDB.GetCodeName("STNNAME", sSql)
                SetObjReadOnly(Me, "txtStnID")
                'Call SetObjFocus(Me.Page, Me.txtDtStart)
                'modDB.InsertSignRecord("AziTest", "是否檢查-2:" & Me.LABCHECKFLAG.Text & "=" & sSql, My.User.Name) 'AZITEST
                If LABCHECKFLAG.Text = "Y" Then
                    CKBCHECKFLAG.Checked = True
                Else
                    CKBCHECKFLAG.Checked = False
                End If
                Me.txtHR11.Focus()
            End If
            '
        End If
    End Sub

#Region "Button Event / Obj Changed"


    '******************************************************************************************************
    '* 顯示站名稱
    '******************************************************************************************************
    Public Sub txtSTNID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStnID.TextChanged
        Dim sSql As String
        If Me.txtStnID.Text.Trim <> "" Then
            sSql = "Select STNID From DUTY_STNREAL Where STNID = '" & Me.txtStnID.Text & "'"
            sSql = modDB.GetCodeName("STNID", sSql)
            If sSql <> "" Then
                modUtil.showMsg(Me.Page, "訊息", "此站之合理工時已建立：" & Me.txtStnID.Text)
                Me.txtStnID.Text = ""
                Me.txtStnName.Text = ""
                Exit Sub
            Else
                sSql = "Select STNNAME From MECHSTNM Where STNID = '" & Me.txtStnID.Text & "'"
                Me.txtStnName.Text = modDB.GetCodeName("STNNAME", sSql)
                If Me.txtStnName.Text = "" Then
                    modUtil.showMsg(Me.Page, "訊息", "查無此站資料：" & Me.txtStnID.Text)
                    Me.txtStnID.Text = ""
                    Exit Sub
                End If
            End If
        End If
    End Sub

#End Region

    '******************************************************************************************************
    '* DB 更新前置檢核處理
    '******************************************************************************************************
    Private Function CheckData(ByVal pMode As FormViewMode) As Boolean
        Dim sValidOK As Boolean = True, I As Integer
        Dim sMsg As String = "", sWeek As String = ""

        Try
            If pMode = FormViewMode.Insert Then '* 新增模式
                sValidOK = sValidOK And CheckNotEmpty(Me, "txtSTNID", sMsg, "站代號")
            End If

        Catch ex As Exception
            sValidOK = False
            modUtil.showMsg(Me.Page, "錯誤訊息(CheckData)", ex.Message)
        End Try
        Return sValidOK
    End Function

    '******************************************************************************************************
    '* DB 更新處理(活動檔期檔)
    '******************************************************************************************************
    Public Sub UpdateDB(ByVal pMode As FormViewMode, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs)
        Try
            If Not CheckData(pMode) Then '* 取消新增/修改
                e.Cancel = True
                Exit Sub
            End If

            '---------------------------------* 建立 Transaction 機制 (用於須同時更新不同檔案時)
            Dim sConn As Data.Common.DbConnection = e.Command.Connection
            If sConn.State = Data.ConnectionState.Closed Then sConn.Open()
            e.Command.Transaction = sConn.BeginTransaction()

            With e.Command
                .Parameters("@STNID").Value = Me.txtStnID.Text
                '
                If (Me.ViewState("FormViewMode") = FormViewMode.Insert) Or (Me.ViewState("FormViewMode") = FormViewMode.Edit) Then

                    If Me.txtHR11.Text.Trim = "" Then
                        .Parameters("@HR11").Value = 0
                    Else
                        .Parameters("@HR11").Value = Convert.ToInt32(Me.txtHR11.Text)
                    End If
                    '
                    If Me.txtHR12.Text.Trim = "" Then
                        .Parameters("@HR12").Value = 0
                    Else
                        .Parameters("@HR12").Value = Convert.ToInt32(Me.txtHR12.Text)
                    End If
                    '
                    If Me.txtHR13.Text.Trim = "" Then
                        .Parameters("@HR13").Value = 0
                    Else
                        .Parameters("@HR13").Value = Convert.ToInt32(Me.txtHR13.Text)
                    End If
                    '
                    If Me.txtHR14.Text.Trim = "" Then
                        .Parameters("@HR14").Value = 0
                    Else
                        .Parameters("@HR14").Value = Convert.ToInt32(Me.txtHR14.Text)
                    End If
                    '
                    If Me.txtHR21.Text.Trim = "" Then
                        .Parameters("@HR21").Value = 0
                    Else
                        .Parameters("@HR21").Value = Convert.ToInt32(Me.txtHR21.Text)
                    End If
                    '
                    If Me.txtHR22.Text.Trim = "" Then
                        .Parameters("@HR22").Value = 0
                    Else
                        .Parameters("@HR22").Value = Convert.ToInt32(Me.txtHR22.Text)
                    End If
                    '
                    If Me.txtHR23.Text.Trim = "" Then
                        .Parameters("@HR23").Value = 0
                    Else
                        .Parameters("@HR23").Value = Convert.ToInt32(Me.txtHR23.Text)
                    End If
                    '
                    If Me.txtHR24.Text.Trim = "" Then
                        .Parameters("@HR24").Value = 0
                    Else
                        .Parameters("@HR24").Value = Convert.ToInt32(Me.txtHR24.Text)
                    End If
                    '
                    If Me.txtHR31.Text.Trim = "" Then
                        .Parameters("@HR31").Value = 0
                    Else
                        .Parameters("@HR31").Value = Convert.ToInt32(Me.txtHR31.Text)
                    End If
                    '
                    If Me.txtHR32.Text.Trim = "" Then
                        .Parameters("@HR32").Value = 0
                    Else
                        .Parameters("@HR32").Value = Convert.ToInt32(Me.txtHR32.Text)
                    End If
                    '
                    If Me.txtHR33.Text.Trim = "" Then
                        .Parameters("@HR33").Value = 0
                    Else
                        .Parameters("@HR33").Value = Convert.ToInt32(Me.txtHR33.Text)
                    End If
                    '
                    If Me.txtHR34.Text.Trim = "" Then
                        .Parameters("@HR34").Value = 0
                    Else
                        .Parameters("@HR34").Value = Convert.ToInt32(Me.txtHR34.Text)
                    End If
                    '
                    If Me.txtDiffADD.Text.Trim = "" Then
                        .Parameters("@DIFFADD").Value = 0
                    Else
                        .Parameters("@DIFFADD").Value = Convert.ToInt32(Me.txtDiffADD.Text)
                    End If
                    '
                    If Me.txtDiffSUB.Text.Trim = "" Then
                        .Parameters("@DIFFSUB").Value = 0
                    Else
                        .Parameters("@DIFFSUB").Value = Convert.ToInt32(Me.txtDiffSUB.Text)
                    End If
                    .Parameters("@UpdateMan").Value = Me.Page.User.Identity.Name
                    If CKBCHECKFLAG.Checked Then
                        .Parameters("@CHECKFLAG").Value = "Y"
                    Else
                        .Parameters("@CHECKFLAG").Value = "N"
                    End If
                    '
                    If Me.ViewState("FormViewMode") = FormViewMode.Insert Then
                        '.Parameters("@DtAppend").Value = Now
                        .Parameters("@CreateMan").Value = Me.Page.User.Identity.Name
                    End If
                End If

                '---------------------------------* 更新 相關檔案
                Dim sSql As String = e.Command.CommandText
                e.Command.CommandText = ""
                'If Not UpdateDB_Detail(pMode, e.Command) Then
                '    e.Command.Transaction.Rollback() '* 失敗時須RollBack
                '    e.Cancel = True
                'End If
                e.Command.CommandText = sSql

            End With
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(UpdateDB)", ex.Message)
        End Try
    End Sub

    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStnID.Click
        'Me.QryStn.Show()
        Me.QryStn.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
    End Sub

End Class
