'******************************************************************************************************
'* 程式：ATT_M_061U 颱風假設定維護
'* 版次：2015/10/28(VER1.01)：新開發
'******************************************************************************************************

Imports modUtil
Imports modUnset
Imports System.Data


Partial Class ATT_M_061U
    Inherits System.Web.UI.UserControl

    'Private cWriteLog As Boolean = False '* 上線時須設成 False
    '******************************************************************************************************
    '* 依編輯模式 初始化輸入畫面
    '******************************************************************************************************
    Public Sub InitScreen(ByVal pMode As FormViewMode)
        'modDB.InsertSignRecord("AziTest", "ATTM059u INIT SCREEN!", My.User.Name) 'AZITEST
        If pMode = FormViewMode.ReadOnly Then
            SetAllReadOnly(Me)
            SHOWCHK()
        Else
            'modDB.InsertSignRecord("AziTest", "ATTM061u INIT-1!", My.User.Name) 'AZITEST
            If pMode = FormViewMode.Insert Then
                'modDB.InsertSignRecord("AziTest", "ATTM056u INIT SCREEN IS IN INSERT_MODE!", My.User.Name) 'AZITEST
                'modUtil.SetDateObj(Me.txtHDATE, False, Nothing, False)
                'modUtil.SetDateImgObj(Me.imgHDATE, Me.txtHDATE, True, False, Nothing)

                'Me.CKBHOLIDAY.Checked = False
                'Me.txtHDATE.Focus()
            Else '-------------------------------------* 顯示名稱

                SetObjReadOnly(Me, "txtDATEST")
                SetObjReadOnly(Me, "txtTimeST")
                SetObjReadOnly(Me, "txtTimeEN")
                SetObjReadOnly(Me, "TxtHODHOUR")
                SetObjReadOnly(Me, "TxtMEMO")
                SetObjReadOnly(Me, "ChkA")
                SetObjReadOnly(Me, "ChkB")
                SetObjReadOnly(Me, "ChkC")
                SetObjReadOnly(Me, "ChkD")
                SetObjReadOnly(Me, "ChkE")
                SetObjReadOnly(Me, "ChkF")
                SetObjReadOnly(Me, "ChkG")
                SetObjReadOnly(Me, "ChkH")
                SetObjReadOnly(Me, "ChkI")
                SetObjReadOnly(Me, "ChkJ")
                SetObjReadOnly(Me, "ChkK")
                SetObjReadOnly(Me, "ChkN")
                SetObjReadOnly(Me, "ChkO")
                SetObjReadOnly(Me, "ChkT")
                SetObjReadOnly(Me, "ChkU")
                SetObjReadOnly(Me, "ChkV")
                SetObjReadOnly(Me, "ChkW")
                SetObjReadOnly(Me, "ChkX")
                SetObjReadOnly(Me, "ChkZ")
                '
            End If
        End If
        '
    End Sub

    Private Sub SHOWCHK()
        'modDB.InsertSignRecord("AziTest", "ATTM061u SHOWCHK()", My.User.Name) 'AZITEST
        Dim sConn As SqlClient.SqlConnection = Nothing
        Dim sCmd As SqlClient.SqlCommand = Nothing
        Dim sSql As String
        Dim sReader As SqlClient.SqlDataReader = Nothing
        sConn = New SqlClient.SqlConnection
        sConn.ConnectionString = Web.Configuration.WebConfigurationManager.ConnectionStrings("MECHPAConnectionString").ToString
        sConn.Open()
        sCmd = New SqlClient.SqlCommand()
        sCmd.Connection = sConn
        '
        sSql = "SELECT CITYSTR FROM HOLYSETM" _
                     & " WHERE HODDATE='" & txtDATEST.Text & "' AND HDTMST= '" & txtHDTMST.Text & "' AND HDTMEN= '" & txtHDTMEN.Text & "' AND STNFLAG='1' AND DELFLAG<>'Y'" _
                     & " ORDER BY HODDATE "
        sCmd.CommandText = sSql
        sReader = sCmd.ExecuteReader()
        sReader.Read()
        If Not IsDBNull(sReader.GetValue(0)) Then
            Dim VCITYSTR As String = sReader.GetString(0)
            'modDB.InsertSignRecord("AziTest", "ATTM061u VCITYSTR=" + VCITYSTR, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "ATTM061u VCITYSTR.IndexOf(A)=" + VCITYSTR.IndexOf("A").ToString, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "ATTM061u VCITYSTR.IndexOf(B)=" + VCITYSTR.IndexOf("B").ToString, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "ATTM061u VCITYSTR.IndexOf(E)=" + VCITYSTR.IndexOf("E").ToString, My.User.Name) 'AZITEST

            If VCITYSTR.IndexOf("A") >= 0 Then
                ChkA.Checked = True
            End If
            If VCITYSTR.IndexOf("B") >= 0 Then
                ChkB.Checked = True
            End If
            If VCITYSTR.IndexOf("C") >= 0 Then
                ChkC.Checked = True
            End If
            If VCITYSTR.IndexOf("D") >= 0 Then
                ChkD.Checked = True
            End If
            If VCITYSTR.IndexOf("E") >= 0 Then
                ChkE.Checked = True
            End If
            If VCITYSTR.IndexOf("F") >= 0 Then
                ChkF.Checked = True
            End If
            If VCITYSTR.IndexOf("G") >= 0 Then
                ChkG.Checked = True
            End If
            If VCITYSTR.IndexOf("H") >= 0 Then
                ChkH.Checked = True
            End If
            If VCITYSTR.IndexOf("I") >= 0 Then
                ChkI.Checked = True
            End If
            If VCITYSTR.IndexOf("J") >= 0 Then
                ChkJ.Checked = True
            End If
            If VCITYSTR.IndexOf("K") >= 0 Then
                ChkK.Checked = True
            End If
            If VCITYSTR.IndexOf("N") >= 0 Then
                ChkN.Checked = True
            End If
            If VCITYSTR.IndexOf("O") >= 0 Then
                ChkO.Checked = True
            End If
            If VCITYSTR.IndexOf("P") >= 0 Then
                ChkP.Checked = True
            End If
            If VCITYSTR.IndexOf("Q") >= 0 Then
                ChkQ.Checked = True
            End If
            If VCITYSTR.IndexOf("T") >= 0 Then
                ChkT.Checked = True
            End If
            If VCITYSTR.IndexOf("U") >= 0 Then
                ChkU.Checked = True
            End If
            If VCITYSTR.IndexOf("V") >= 0 Then
                ChkV.Checked = True
            End If
            If VCITYSTR.IndexOf("X") >= 0 Then
                ChkX.Checked = True
            End If
            If VCITYSTR.IndexOf("Z") >= 0 Then
                ChkZ.Checked = True
            End If
        End If
        sReader.Close()
        sConn.Close()
        sCmd.Dispose() '手動釋放資源
        sConn.Dispose()
        sConn = Nothing '移除指標
    End Sub

    '******************************************************************************************************
    '* DB 更新前置檢核處理
    '******************************************************************************************************
    Private Function CheckData(ByVal pMode As FormViewMode) As Boolean
        Dim sValidOK As Boolean = True
        Dim sMsg As String = ""
        Dim CHKRSTR As String
        '
        Try
            'sValidOK = sValidOK And CheckNotEmpty(Me, "txtSTNWH", sMsg, "標準時數")
            '排班日期檢查
            If sValidOK Then
                CHKRSTR = modUnset.PADATECHK(txtDATEST.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "日期錯誤!請檢查!"
                Else
                    txtDATEST.Text = CHKRSTR.Substring(1, 8)
                End If
            End If

            '起始時間檢查
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(txtHDTMST.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "起始時間錯誤!請檢查!"
                Else
                    txtHDTMST.Text = CHKRSTR.Substring(1, 4)
                End If
            End If

            '結束時間檢查
            If sValidOK Then
                CHKRSTR = modUnset.PATIMECHK(txtHDTMEN.Text)
                If CHKRSTR.Substring(0, 1) <> "Y" Then
                    sValidOK = False : sMsg = "結束時間錯誤!請檢查!"
                Else
                    txtHDTMEN.Text = CHKRSTR.Substring(1, 4)
                End If
            End If
            '
            If sValidOK Then '再作鎖檔日期檢核
                Dim LCKDT As String = GET_LOCKDT("000000")
                If txtDATEST.Text <= LCKDT Then
                    sValidOK = False
                    sMsg = "資料週期已鎖或過帳，不可異動!"
                End If
            End If
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
            'modDB.InsertSignRecord("AziTest", "ATTM061u UpdateDB-1!", My.User.Name) 'AZITEST
            If Not CheckData(pMode) Then '* 取消新增/修改
                e.Cancel = True
                Exit Sub
            End If
            '
            Dim vCITYSTR = ""
            If ChkA.Checked Then
                vCITYSTR = vCITYSTR & "A"
            End If
            If ChkB.Checked Then
                vCITYSTR = vCITYSTR & "B"
            End If
            If ChkC.Checked Then
                vCITYSTR = vCITYSTR & "C"
            End If
            If ChkD.Checked Then
                vCITYSTR = vCITYSTR & "D"
            End If
            If ChkE.Checked Then
                vCITYSTR = vCITYSTR & "E"
            End If
            If ChkF.Checked Then
                vCITYSTR = vCITYSTR & "F"
            End If
            If ChkG.Checked Then
                vCITYSTR = vCITYSTR & "G"
            End If
            If ChkH.Checked Then
                vCITYSTR = vCITYSTR & "H"
            End If
            If ChkI.Checked Then
                vCITYSTR = vCITYSTR & "I"
            End If
            If ChkJ.Checked Then
                vCITYSTR = vCITYSTR & "J"
            End If
            If ChkK.Checked Then
                vCITYSTR = vCITYSTR & "K"
            End If
            If ChkN.Checked Then
                vCITYSTR = vCITYSTR & "N"
            End If
            If ChkO.Checked Then
                vCITYSTR = vCITYSTR & "O"
            End If
            If ChkP.Checked Then
                vCITYSTR = vCITYSTR & "P"
            End If
            If ChkQ.Checked Then
                vCITYSTR = vCITYSTR & "Q"
            End If
            If ChkT.Checked Then
                vCITYSTR = vCITYSTR & "T"
            End If
            If ChkU.Checked Then
                vCITYSTR = vCITYSTR & "U"
            End If
            If ChkV.Checked Then
                vCITYSTR = vCITYSTR & "V"
            End If
            If ChkX.Checked Then
                vCITYSTR = vCITYSTR & "X"
            End If
            If ChkZ.Checked Then
                vCITYSTR = vCITYSTR & "Z"
            End If

            '---------------------------------* 建立 Transaction 機制 (用於須同時更新不同檔案時)
            'Dim sConn As Data.Common.DbConnection = e.Command.Connection
            'If sConn.State = Data.ConnectionState.Closed Then sConn.Open()
            'e.Command.Transaction = sConn.BeginTransaction()
            'modDB.InsertSignRecord("AziTest", "ATTM061u UpdateDB-2!", My.User.Name) 'AZITEST
            With e.Command
                .Parameters("@HODDATE").Value = Me.txtDATEST.Text
                .Parameters("@HDTMST").Value = Me.txtHDTMST.Text
                .Parameters("@HDTMEN").Value = Me.txtHDTMEN.Text

                '-------------------------------------* 刪除舊資料
                '<asp:Parameter Name="VATM_SEQ"     Type="Int32"  />
                If (pMode = FormViewMode.Insert) Or (pMode = FormViewMode.Edit) Then
                    'AZITEMP 20151030
                    .Parameters("@HODHOURS").Value = Convert.ToDouble(Me.TxtHODHOUR.Text)
                    .Parameters("@HODMEMO").Value = Me.TxtMEMO.Text
                    .Parameters("@CITYSTR").Value = vCITYSTR
                    .Parameters("@MODUSER").Value = Me.Page.User.Identity.Name
                    'modDB.InsertSignRecord("AziTest", "ATTM061u UpdateDB-3!", My.User.Name) 'AZITEST
                End If 'END OF INSERT
            End With
        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(UpdateDB)", ex.Message)
        End Try
    End Sub

End Class
