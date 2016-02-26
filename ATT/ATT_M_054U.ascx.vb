'******************************************************************************************************
'* 程式：ATT_M_053U 打卡資料維護
'* //版次：2013/05/31(VER1.01)：新開發
'******************************************************************************************************

Imports modUtil
Imports modUnset
Imports System.Data


Partial Class ATT_M_054U
    Inherits System.Web.UI.UserControl

    'Private cWriteLog As Boolean = False '* 上線時須設成 False
    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Public Sub InitScreen(ByVal pMode As FormViewMode)
        Me.btnStnID.Visible = False '暫不開放

        If pMode = FormViewMode.ReadOnly Then
            SetAllReadOnly(Me)
            Me.btnStnID.Visible = False
        Else
            SetObjReadOnly(Me, "txtStnID")
            SetObjReadOnly(Me, "txtStnName")
            If pMode = FormViewMode.Insert Then
                '只考慮新增模式()
                Me.txtSTNID.Text = HttpUtility.UrlDecode(Request.Cookies("STNID").Value)
                Me.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies("STNNAME").Value)
                'Me.txtSTNID.Text = Session("VSTNID").ToString
                'Me.txtStnName.Text = Session("VSTNNAME").ToString
            End If
        End If
    End Sub

    '******************************************************************************************************
    '* DB 更新前置檢核處理
    '******************************************************************************************************
    Private Function CheckData(ByVal pMode As FormViewMode) As Boolean
        Dim sValidOK As Boolean = True
        Dim sMsg As String = ""

        Try
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtCLSID", sMsg, "起始日期")
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtSTTIME", sMsg, "起始時間")
            sValidOK = sValidOK And CheckNotEmpty(Me, "txtEDTIME", sMsg, "結束時間")
            '
            If sValidOK And (pMode = FormViewMode.Insert) Then
                Dim sConn As SqlClient.SqlConnection = Nothing
                Dim sCmd As SqlClient.SqlCommand = Nothing
                Dim sSql As String

                sConn = New SqlClient.SqlConnection
                sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
                sConn.Open()

                sCmd = New SqlClient.SqlCommand()
                sCmd.Connection = sConn

                Dim sReader As SqlClient.SqlDataReader = Nothing
                sCmd = New SqlClient.SqlCommand()
                sCmd.Connection = sConn

                '-------------------------------------* 找出是否有重複資料
                sSql = "SELECT CLSID FROM CLSTIME " _
                     & " WHERE STNID='" & txtSTNID.Text & "'" _
                     & "   AND CLSID='" & txtCLSID.Text & "'"
                sCmd.CommandText = sSql
                sReader = sCmd.ExecuteReader()
                sReader.Read()

                If (sReader.HasRows) Then
                    sValidOK = False
                    sMsg = "已有資料重複!!"
                End If
            End If

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
        'modUtil.showMsg(Me.Page, "鎖檔日期", LCKDT)
        If pMode = FormViewMode.Insert Then
            'If Me.txtSHSTTIME.Text > Me.txtSHEDTIME.Text Then
            'txtSHEDDATE.Text = Format(CDate(Format(CInt(Me.txtSHSTDATE.Text), "0000/00/00")).AddDays(1), "yyyyMMdd")
            'End If
        End If

        Try
            If Not CheckData(pMode) Then '* 取消新增/修改
                e.Cancel = True
                Exit Sub
            End If

            With e.Command
                .Parameters("@STNID").Value = Me.txtSTNID.Text
                .Parameters("@CLSID").Value = Me.txtCLSID.Text
                '-------------------------------------* 刪除舊資料
                If pMode = FormViewMode.Insert Then
                    .Parameters("@STTIME").Value = Me.txtSTTIME.Text
                    .Parameters("@EDTIME").Value = Me.txtEDTIME.Text
                    .Parameters("@NFLAG").Value = IIf(Me.chkNflag.Checked, "Y", "N")
                End If 'END OF INSERT
            End With
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(UpdateDB)", ex.Message)
        End Try

    End Sub

    Protected Sub btnStnID_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.QryStn.Show(Me.ViewState("RolType"), Me.ViewState("RolDeptID"))
    End Sub

    Protected Sub txtStnID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'modCharge.GetStnName(Me.txtSTNID, Me.txtStnName, Me.ViewState("RolDeptID"))
    End Sub

End Class
