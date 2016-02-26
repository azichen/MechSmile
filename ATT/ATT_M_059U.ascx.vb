'******************************************************************************************************
'* 程式：ATT_M_056U 站假日設定資料維護
'* //版次：2013/05/31(VER1.01)：新開發
'******************************************************************************************************

Imports modUtil
Imports modUnset
Imports System.Data


Partial Class ATT_M_059U
    Inherits System.Web.UI.UserControl

    'Private cWriteLog As Boolean = False '* 上線時須設成 False
    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Public Sub InitScreen(ByVal pMode As FormViewMode)
        'modDB.InsertSignRecord("AziTest", "ATTM059u INIT SCREEN!", My.User.Name) 'AZITEST
        If pMode = FormViewMode.ReadOnly Then
            'modDB.InsertSignRecord("AziTest", "ATTM059u INIT-1!", My.User.Name) 'AZITEST
            SetAllReadOnly(Me)
            'modDB.InsertSignRecord("AziTest", "ATTM056u INIT SCREEN IS IN Readonly_MODE!", My.User.Name) 'AZITEST
        Else
            'modDB.InsertSignRecord("AziTest", "ATTM059u INIT-2!", My.User.Name) 'AZITEST
            'If pMode = FormViewMode.Insert Then
            '    'modDB.InsertSignRecord("AziTest", "ATTM056u INIT SCREEN IS IN INSERT_MODE!", My.User.Name) 'AZITEST
            '    modUtil.SetDateObj(Me.txtHDATE, False, Nothing, False)
            '    modUtil.SetDateImgObj(Me.imgHDATE, Me.txtHDATE, True, False, Nothing)

            '    Me.CKBHOLIDAY.Checked = False
            '    Me.txtHDATE.Focus()
            'Else '-------------------------------------* 顯示名稱

            SetObjReadOnly(Me, "txtDATEST")
            SetObjReadOnly(Me, "txtDATEEN")
            'End If
        End If
        '
    End Sub

    '******************************************************************************************************
    '* DB 更新前置檢核處理
    '******************************************************************************************************
    Private Function CheckData(ByVal pMode As FormViewMode) As Boolean
        Dim sValidOK As Boolean = True
        Dim sMsg As String = ""
        Try
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtSTNWH", sMsg, "標準時數")
            '
            If Not sValidOK Then showMsg(Me.Page, "訊息(請修正錯誤欄位)", sMsg)
        Catch ex As Exception
            sValidOK = False
            modUtil.showMsg(Me.Page, "錯誤訊息(CheckData)", ex.Message)
        End Try
        Return sValidOK
    End Function

    '******************************************************************************************************
    '* DB 更新處理(主檔)
    '******************************************************************************************************
    Public Sub UpdateDB(ByVal pMode As FormViewMode, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs)
        Try
            If Not CheckData(pMode) Then '* 取消新增/修改
                e.Cancel = True
                Exit Sub
            End If
            '---------------------------------* 建立 Transaction 機制 (用於須同時更新不同檔案時)
            'Dim sConn As Data.Common.DbConnection = e.Command.Connection
            'If sConn.State = Data.ConnectionState.Closed Then sConn.Open()
            'e.Command.Transaction = sConn.BeginTransaction()
            With e.Command
                .Parameters("@FWDS_DATEST").Value = Me.txtDATEST.Text
                '.Parameters("@FWDS_STNWH").Value = Me.txtSTNWH.Text
                .Parameters("@FWDS_STNWH").Value = Convert.ToDouble(Me.txtSTNWH.Text)
                '---------------------------------* 更新 相關檔案
            End With
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(UpdateDB)", ex.Message)
        End Try
    End Sub

End Class
