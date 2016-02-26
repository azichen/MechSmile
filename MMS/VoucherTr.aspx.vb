'******************************************************************************************************
'* 程式：CHG_Q_016 簽帳會計傳票匯入作業
'* 作成：NEC 杜志揚
'* 版次：2008/01/07(VER1.01)：新開發
'* 版次：2008/07/08(VER1.02)：簽帳額度之恢復變更為依金額之正負值沖轉，原依會計科目
'* 版次：2008/08/12(VER1.03)：1.取消更新發票已全數沖轉FLAG(改由收款維護更新) 2.將[因異常不處理]變更為[OK但不處理]
'******************************************************************************************************

Imports system.Data
Imports system.Data.OleDb
Imports system.Data.SqlClient

Partial Class VoucherTr
    Inherits System.Web.UI.Page

    Private cWriteLog As Boolean = True, cErr As Boolean
    Private cConn As SqlConnection = Nothing
    Private cCmd As SqlCommand = Nothing
    Private cTrans As SqlTransaction = Nothing
    Private cSql, cConnStr, cLogFile As String


#Region "Page Event"
    '******************************************************************************************************
    '* Page_Load => 在此頁面初始化設定
    '******************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '--------------------------------------------------* 檢查是否有登錄User.Identity.Name
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else '---------------------------------------------* 檢查是否有讀取權限(須財務人員)
                ViewState("ROL") = modUtil.GetRolData(Request)
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
                If (Mid(ViewState("ROL"), 5, 2) <> "03") And (Mid(ViewState("ROL"), 5, 2) <> "05") Then FormsAuthentication.RedirectToLoginPage()
            End If

            Me.lblTitle.Text = "會計傳票匯入作業"   '* 功能標題
            modDB.SetGridViewStyle(Me.grdList)          '* 套用gridview樣式
            Me.grdList.Width = 775
            '--------------------------------------------------* 加入欄位並標題顯示中文
            modDB.SetFields("CmpID", "公司編號", grdList, False)
            modDB.SetFields("AccNo", "傳票編號", grdList, False)
            modDB.SetFields("DtAcc", "傳票日期", grdList, False)
            modDB.SetFields("AccItem", "會計科目", grdList, False)
            modDB.SetFields("AccAmt", "金額", grdList, HorizontalAlign.Right, "{0:N0}", False)
            modDB.SetFields("InvoiceNo", "發票號碼", grdList, False)
            modDB.SetFields("DtInvo", "發票日期", grdList, False)

            Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 7, Me.ViewState("EmptyTable"))
            Me.grdList.AllowPaging = False
            Me.grdList.RowStyle.Height = 18
            'Me.fileUp.Focus()
        End If

        'Me.grdList.DataSource = Me.ViewState("tblExcel")
    End Sub

#End Region

#Region "GridView Event"
    '******************************************************************************************************
    '* GridView_DataBound => 顯示頁碼
    '******************************************************************************************************
    Protected Sub grdList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdList.DataBound
        'modDB.SetGridPageNum(Me.grdList)
    End Sub

    '******************************************************************************************************
    '* GridView_RowDataBound => 游標移動時，顯示光筆效果
    '******************************************************************************************************
    Protected Sub grdList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdList.RowDataBound
        modDB.SetGridLightPen(e) '* 游標移動時，顯示光筆效果
    End Sub
#End Region

#Region "Button Event"

    '******************************************************************************************************
    '* 清除
    '******************************************************************************************************
    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("QryField") = Nothing
        Me.lblMsg.Text = "" : Me.btnSave.Enabled = False
        Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 8, Me.ViewState("EmptyTable"))
        Me.updGrid.Update()
    End Sub

    '******************************************************************************************************
    '* Excel 匯出
    '******************************************************************************************************


    '******************************************************************************************************
    '* 匯入
    '******************************************************************************************************
    Protected Sub btnImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnImport.Click
        Dim sConn As OleDbConnection = Nothing
        Dim sCmd As OleDbCommand = Nothing
        Dim sAdapter As OleDbDataAdapter = Nothing
        Dim sTable As DataTable = Nothing

        Try
            If Not Me.fileUp.HasFile Then
                modUtil.showMsg(Me.Page, "訊息", "尚未指定檔案！")
                Me.btnSave.Enabled = False
                Exit Sub

            Else
                If Not (Me.fileUp.PostedFile Is Nothing) Then
                    '--------------------------------* 上傳 Excel 檔案至 Server 暫存區
                    Dim sFileName As String = Server.MapPath("~\App_Data\") & "M016.xls"
                    If IO.File.Exists(sFileName) Then IO.File.Delete(sFileName)
                    Me.fileUp.SaveAs(sFileName)

                    Try
                        '--------------------------------* 讀取 Excel 資料
                        Dim sConnStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFileName _
                                               & ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
                        sConn = New OleDbConnection(sConnStr)
                        sConn.Open()
                        sCmd = New OleDbCommand("SELECT * FROM [Sheet1$]", sConn)
                        sAdapter = New OleDbDataAdapter(sCmd)
                        sTable = New DataTable()
                        sTable.Columns.Add(New DataColumn("客戶代號", GetType(String)))
                        sTable.Columns.Add(New DataColumn("傳票編號", GetType(String)))
                        sTable.Columns.Add(New DataColumn("傳票日期", GetType(String)))
                        sTable.Columns.Add(New DataColumn("會計科目", GetType(String)))
                        sTable.Columns.Add(New DataColumn("金額", GetType(Int32)))
                        sTable.Columns.Add(New DataColumn("發票號碼", GetType(String)))
                        sTable.Columns.Add(New DataColumn("發票日期", GetType(String)))

                        sAdapter.Fill(sTable)

                        'sTable.Columns(0).ColumnName = "Status"    '* 處理狀態
                        sTable.Columns(0).ColumnName = "CmpID"     '* 客戶代號(6)
                        sTable.Columns(1).ColumnName = "AccNo"     '* 傳票編號(12)
                        sTable.Columns(2).ColumnName = "DtAcc"     '* 傳票日期(8):YYYYMMDD
                        sTable.Columns(3).ColumnName = "AccItem"   '* 會計科目(12)
                        sTable.Columns(4).ColumnName = "AccAmt"    '* 金額(9)
                        sTable.Columns(5).ColumnName = "InvoiceNo" '* 發票號碼(10)
                        sTable.Columns(6).ColumnName = "DtInvo"    '* 發票日期(8):YYYYMMDD


                    Catch ex As Exception
                        modUtil.showMsg(Me.Page, "錯誤訊息(匯入)", ex.Message)
                        Exit Sub
                    Finally
                        '--------------------------------* 釋放資源 並刪除檔案
                        sAdapter.Dispose() : sAdapter = Nothing
                        sCmd.Dispose() : sCmd = Nothing
                        sConn.Close() : sConn.Dispose() : sConn = Nothing
                        IO.File.Delete(sFileName)
                    End Try

                    If sTable.Rows.Count > 0 Then
                        Me.btnSave.Enabled = True
                        '---------------------------------------* Step0. 將Null值補成空白 
                        Dim I, J As Integer
                        For I = 0 To sTable.Rows.Count - 1
                            For J = 0 To sTable.Columns.Count - 1
                                If sTable.Rows(I).IsNull(J) Then
                                    If J = 5 Then
                                        sTable.Rows(I).Item(J) = 0
                                    Else
                                        sTable.Rows(I).Item(J) = ""
                                    End If
                                Else
                                    sTable.Rows(I).Item(J) = Trim(sTable.Rows(I).Item(J))
                                    If (J = 3 Or J = 7) And Len(sTable.Rows(I).Item(J)) = 8 Then
                                        sTable.Rows(I).Item(J) = modUtil.GetDateES(sTable.Rows(I).Item(J))
                                    End If
                                End If
                            Next
                        Next

                        Me.ViewState("tblExcel") = sTable
                        Me.grdList.DataSource = sTable
                        Me.grdList.DataBind()

                    Else
                        Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 6, Me.ViewState("EmptyTable"))
                        modUtil.showMsg(Me.Page, "訊息", "查無會計傳票資料！")
                    End If
                    Me.updGrid.Update()
                End If
            End If

        Catch ex As Exception
            modUtil.showMsg(Me.Page, "錯誤訊息(匯入)", ex.Message)
        End Try
    End Sub

    '******************************************************************************************************
    '* 匯入
    '******************************************************************************************************
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim i As Integer
        Dim tmpsql As String
        For i = 0 To Me.grdList.Rows.Count - 1
            tmpsql = "insert into MMSVoucherNumber([CustomerNo],[DocNo],[DocDate],[Accounts],[Amount],[InvoiceNo],[InvoiceDate]) "
            tmpsql += "values("
            tmpsql += "'" + Me.grdList.Rows(i).Cells(1).Text + "',"
            tmpsql += "'" + Me.grdList.Rows(i).Cells(2).Text + "',"
            tmpsql += "'" + Me.grdList.Rows(i).Cells(3).Text + "',"
            tmpsql += "'" + Me.grdList.Rows(i).Cells(4).Text + "',"
            tmpsql += Me.grdList.Rows(i).Cells(5).Text.Replace(",", "") + ","
            tmpsql += "'" + Me.grdList.Rows(i).Cells(6).Text + "',"
            tmpsql += "'" + Me.grdList.Rows(i).Cells(7).Text + "'"
            tmpsql += ")"
            'modDB.InsertSignRecord("AziTest", "Cells(1)=" & Me.grdList.Rows(i).Cells(1).Text, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "Cells(2)=" & Me.grdList.Rows(i).Cells(2).Text, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "Cells(3)=" & Me.grdList.Rows(i).Cells(3).Text, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "Cells(4)=" & Me.grdList.Rows(i).Cells(4).Text, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "Cells(5)=" & Me.grdList.Rows(i).Cells(5).Text, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "Cells(6)=" & Me.grdList.Rows(i).Cells(6).Text, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "Cells(7)=" & Me.grdList.Rows(i).Cells(7).Text, My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql.Substring(0, 90), My.User.Name) 'AZITEST
            'modDB.InsertSignRecord("AziTest", "tmpsql=" & tmpsql.Substring(90, 90), My.User.Name) 'AZITEST
            Me.EXE_SQL(tmpsql)
        Next
        modUtil.showMsg(Me.Page, "訊息", "資料匯入完畢！")
        Session("QryField") = Nothing
        Me.lblMsg.Text = "" : Me.btnSave.Enabled = False
        Call modDB.ShowEmptyDataGridHeader(Me.grdList, 0, 7, Me.ViewState("EmptyTable"))
        'Me.updGrid.Update()
    End Sub

    '******************************************************************************************************
    '* Update DB 
    '******************************************************************************************************
    Private Function UpdateDB() As Boolean
        Try
            If cWriteLog Then modUtil.WriteLog(cSql, "Q016")
            cCmd.CommandType = CommandType.Text
            cCmd.CommandText = cSql
            cCmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            If cWriteLog Then modUtil.WriteLog("WriteDB Err:" & ex.Message, "Q016")
            cErr = True
            Return False
        End Try
    End Function


#End Region



    '執行quary不回傳值
    Public Function EXE_SQL(ByVal TMPSQL1 As String) As Boolean
        Dim AA As Integer
        Dim CONNSTR1 As String = modDB.GetDataSource.ConnectionString
        Dim CONN1 As New System.Data.SqlClient.SqlConnection(CONNSTR1)
        Dim SQLCOMMAND1 As New System.Data.SqlClient.SqlCommand
        Try
            SQLCOMMAND1.CommandText = TMPSQL1
            If CONN1.State.Closed = ConnectionState.Open Then CONN1.Close()
            SQLCOMMAND1.Connection = CONN1
            CONN1.Open()
            SQLCOMMAND1.ExecuteNonQuery()
            CONN1.Close()
        Catch ex As Exception
            EXE_SQL = False
            Exit Function
        End Try
        EXE_SQL = True
    End Function


End Class
