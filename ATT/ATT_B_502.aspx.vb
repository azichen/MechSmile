'******************************************************************************************************
'* 程式：ATT_B_502 加油站合理工時檢查電文
'* 作成：SMILE 劉遇龍
'* 版次：2014/02/11(VER1.00)：新開發
'* 版次：2014/03/06(VER1.01)：憲哥:只判斷超排即可，短排不提示
'* 版次：2014/03/07(VER1.02)：志誠:超短排差異分成兩欄，以此判斷
'* 版次：2014/03/10(VER1.03)：修正帶入日期格式問題
'* 版次：2014/03/14(VER1.04)：志誠:已開始的上班日無須驗證，因此帶入檢核日期改以系統日期為準
'******************************************************************************************************

Imports System.Data.SqlClient

Partial Class ATT_ATT_B_502
    Inherits System.Web.UI.Page

    Private cWriteLog As Boolean = True '* 正式上線時須設成 False
    Private cStart, cEnd, cLens As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.ContentEncoding = System.Text.Encoding.Default '* 避免中文亂碼
        Call GetTradeInfo()
    End Sub

    '******************************************************************************************************
    '* 解析 交易欄位資料
    '******************************************************************************************************
    Private Function GetFieldValue(ByRef pStatus As String, ByVal pSrc As String, ByRef pField As String, ByVal pFieldName As String, ByVal pChkEmpty As Boolean) As Boolean
        Try
            cLens = Len("<" & pFieldName & ">")
            cStart = InStr(pSrc, "<" & pFieldName & ">")
            cEnd = InStr(pSrc, "</" & pFieldName & ">")
            pField = Trim(Mid(pSrc, cStart + cLens, cEnd - cStart - cLens))
            If pChkEmpty And pField = "" Then
                pStatus = "3" : Return False '* 電文欄位錯誤
            Else
                Return True
            End If
        Catch ex As Exception
            pStatus = "3" : Return False '* 電文欄位錯誤
        End Try
    End Function

    '******************************************************************************************************
    '* 取得指定日期站所的合理工時
    '******************************************************************************************************
    Private Function ChkHR(ByVal CutStn As String, ByVal CutDate As String)
        Dim sConn As SqlConnection = Nothing
        Dim sCmd As SqlCommand = Nothing
        Dim sReader As SqlDataReader = Nothing
        Dim sSql, wCode, oCode, wCodeStr, oCodeStr, AdjDate As String
        Dim AdjAmt, WORKHOUR As Single
        Dim getHR, getDiffAdd, getDiffSub As Integer
        sSql = "" : wCode = "" : oCode = "" : wCodeStr = "" : oCodeStr = "" : AdjDate = "" : AdjAmt = 0 : WORKHOUR = 0 : getHR = 0 : getDiffAdd = 0 : getDiffSub = 0
        Try
            sConn = modDB.GetSqlConnection()
            sCmd = New SqlCommand()
            sCmd.Connection = sConn
            '--取得加油站當日氣候類別
            sSql = "select AttCode from StationWeather_View where STNID='" & CutStn & "' AND CutDate='" & CutDate & "'"
            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()
            If sReader.Read Then
                wCode = sReader.Item("AttCode")
            Else
                wCode = "1"
            End If
'Return sSql
            sReader.Close()
            '--取得當日有無漲跌價（以九五油品，且差額在0.4元以上為準）
            sSql = "select DtAdjust, AdjustAmt from OilPrice " _
            & " where OilID='05' and abs(AdjustAmt)>=0.4 " _
            & " and Convert(varchar,DateAdd(day,0,'" & CutDate & "'),111) in (Convert(varchar,DateAdd(day,-1,DtAdjust),111),DtAdjust)"
            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()
            If sReader.Read Then
                AdjDate = sReader.Item("DtAdjust")
                AdjAmt = sReader.Item("AdjustAmt")
                '依取得漲跌價結果判斷
                If AdjAmt > 0 Then '漲價
                    If AdjDate = CutDate Then '漲價當日
                        oCode = "3"
                    Else '漲價前一日以降價計算
                        oCode = "4"
                    End If
                ElseIf AdjAmt < 0 Then
                    If AdjDate = CutDate Then '降價當日
                        oCode = "4"
                    Else '降價前一日以漲價計算
                        oCode = "3"
                    End If
                End If
            End If
            sReader.Close()
            '--若油價漲跌無結果，便以平假日為判斷依據
            If oCode = "" Then
                sSql = "select HOLIDAY from Duty_holiday where HDATE='" & CutDate & "'"
                sCmd.CommandText = sSql
                sReader = sCmd.ExecuteReader()
                If sReader.Read Then '假日檔有特別設定
                    If sReader.Item("HOLIDAY") = "Y" Then
                        oCode = "2" '假日
                    Else
                        oCode = "1" '平日
                    End If
                Else '無特別設定
                    If DatePart(DateInterval.Weekday, CDate(CutDate)) = 7 Or DatePart(DateInterval.Weekday, CDate(CutDate)) = 1 Then
                        oCode = "2" '週末=假日
                    Else
                        oCode = "1" '平日
                    End If
                End If
                sReader.Close()
            End If
            '--依前述條件，取得合理工時設定
            sSql = "select hr" & wCode & oCode & ",DiffAdd,DiffSub from Duty_stnreal where CheckFlag='Y' and STNID= '" & CutStn & "'"
            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()
            If sReader.Read Then
                getHR = sReader.Item("hr" & wCode & oCode)
                getDiffAdd = sReader.Item("DiffAdd")
                getDiffSub = sReader.Item("DiffSub")
            Else
                Return "尚未設定合理工時"
            End If
            sReader.Close()
            '--依取得條件判斷顯示文字
            Select Case wCode
                Case "1" : wCodeStr = "晴天"
                Case "2" : wCodeStr = "陰天"
                Case "3" : wCodeStr = "雨天"
            End Select
            Select Case oCode
                Case "1" : oCodeStr = "平日"
                Case "2" : oCodeStr = "假日"
                Case "3" : oCodeStr = "漲價"
                Case "4" : oCodeStr = "降價"
            End Select
            '--取得當日實際工時合計
            sSql = "select ISNULL(sum(WORKHOUR),0) AS WORKHOUR from [192.168.1.1].[MECHPA].dbo.SCHEDM with(nolock)" _
            & "where STNID='" & CutStn & "' and SHSTDATE=Convert(varchar,DateAdd(day,0,'" & CutDate & "'),112)"
            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()
            '比對有否超、短排合理工時
            If sReader.Read Then
                WORKHOUR = sReader.Item("WORKHOUR")
                If WORKHOUR = 0 Then
                    Return CutDate & "尚未排定工時"
                ElseIf WORKHOUR > (getHR + getDiffAdd) Then '超排
                    Return CutDate & "超排工時" & "(已排:" & WORKHOUR.ToString & " ＞ " & wCodeStr & oCodeStr & ":" & getHR.ToString & ")"
                ElseIf WORKHOUR < (getHR - getDiffSub) Then '短排
                    Return CutDate & "短排工時" & "(已排:" & WORKHOUR.ToString & " ＜ " & wCodeStr & oCodeStr & ":" & getHR.ToString & ")"
                Else
                    Return ""
                End If
            Else
                Return CutDate & "尚未排定工時"
            End If
        Catch ex As Exception
            Return CutDate & "工時判斷異常" & ex.Message.ToString
        End Try
    End Function
    '******************************************************************************************************
    '* 取得上傳之交易資料，並下傳相關資訊電文
    '******************************************************************************************************
    Private Sub GetTradeInfo()
        Dim sConn As SqlConnection = Nothing
        Dim sCmd As SqlCommand = Nothing
        Dim sReader As SqlDataReader = Nothing
        Dim sTrans As SqlTransaction = Nothing
        Dim sSql, sTradeData, sStatus, sMsg, ChkResult As String
        Dim I, CutDays As Integer

        CutDays = 3 '檢查天數(預設後推三天)
        sTradeData = "" : sStatus = "" : sMsg = "" : ChkResult = ""
        Try
            sTradeData = Request("XMLData")
            If cWriteLog Then modUtil.WriteLog(sTradeData, "ATT_B_502")
            If sTradeData = "" Then sStatus = "3" : Exit Sub
            'sTradeData = "<STNID>420999</STNID><CutDate>2014/02/10</CutDate>"
            '完整電文Sample：http://10.1.1.111/ATT/ATT_B_502.aspx?XMLData=<STNID>421001</STNID><CutDate>2014/03/14</CutDate>

            sConn = modDB.GetSqlConnection()
            sCmd = New SqlCommand()
            sCmd.Connection = sConn

            '--------------------------------------------------------------* Step1. 解析欄位
            Dim sSTNID, sCutDate As String
            sSTNID = "" : sCutDate = "" 
            If Not GetFieldValue(sStatus, sTradeData, sSTNID, "STNID", True) Then sStatus = "3" : Exit Sub '* 站別
            If Not GetFieldValue(sStatus, sTradeData, sCutDate, "CutDate", True) Then sStatus = "3" : Exit Sub '* 計算日期
            'If Mid(sDtAcc, 3, 1) = "/" Then '* (2008/07/12)修正POS日期之格式(POS:MM/dd/yyyy)
            '    sDtAcc = Mid(sDtAcc, 7, 4) & "/" & Mid(sDtAcc, 1, 5)
            'End If

            '--------------------------------------------------------------* Step2. 先確認此站是否須檢查預排工時(ChkFlag='Y')
            sSql = "select CheckFlag from Duty_stnreal where STNID= '" & sSTNID & "'"
            sCmd.CommandText = sSql
            sReader = sCmd.ExecuteReader()
            If sReader.Read Then
                If sReader.Item("CheckFlag").ToString = "N" Then
                    sStatus = "1"  '* 此站免檢查，沿用檢查成功的設定
                Else
                    '志誠:已開始的上班日無須驗證，因此帶入檢核日期改以系統日期為準 by EricLiu@20140313
                    sCutDate = Format(Today, "yyyy/MM/dd")
                    '須檢查，自結帳日後推三天
                    For I = 0 To CutDays - 1
                        '逐日檢測是否超時(Function ChkHR)
                        ChkResult = ChkHR(sSTNID, Format(CDate(sCutDate).AddDays(I + 1), "yyyy/MM/dd"))
                        '依檢測結果組合顯示文字
                        If ChkResult <> "" Then
                            'If sMsg <> "" Then
                            '    sMsg = sMsg & "," & ChkResult
                            'Else
                            sMsg = sMsg & ChkResult
                            Exit For
                            'End If
                        End If
                    Next
                    '若有訊息文字，才需要顯示
                    If sMsg <> "" Then
                        sStatus = "2" '有異常需提醒
                    Else
                        sStatus = "1" '檢查無異常
                    End If
                End If
            Else '若查無此站資料
                sStatus = "5"  '* 查無資料
                sMsg = "尚未設定合理工時，請叫修!"
            End If
            sReader.Close()

        Catch ex As Exception
            If sStatus <> "3" Then sStatus = "4" : sMsg = ex.Message.ToString '* 系統處理錯誤
            If sTrans IsNot Nothing Then sTrans.Rollback()
            If cWriteLog Then modUtil.WriteLog("Err=>" & ex.Message, "ATT_B_502")

        Finally
            If sReader IsNot Nothing Then sReader.Close() : sReader = Nothing
            If cWriteLog Then modUtil.WriteLog("處理結果=>" & sStatus, "ATT_B_502")
            Response.Clear()

            Select Case sStatus
                Case "0" '志誠:備用代碼，目前尚未定義用途
                    Response.Write("<CODE>" & sStatus & "</CODE>")
                    sStatus = "S" : Response.Write("<DESC>成功</DESC>")
                Case "1" '* 檢查成功
                    Response.Write("<CODE>" & sStatus & "</CODE>")
                    sStatus = "S" : Response.Write("<DESC>檢查成功</DESC>")
                Case "2" '* 檢查異常
                    Response.Write("<CODE>" & sStatus & "</CODE>")
                    sStatus = "F" : Response.Write("<DESC>" & sMsg & "</DESC>")
                Case "3" '* 欄位錯誤
                    modUtil.WriteLog("ATT_B_502 欄位錯誤=>" & sTradeData, "ATT_B_502")
                    Response.Write("<CODE>" & sStatus & "</CODE>")
                    sStatus = "F" : Response.Write("<DESC>電文欄位錯誤</DESC>")
                Case "4" '* 系統錯誤
                    modUtil.WriteLog("ATT_B_502 系統錯誤=>" & sTradeData, "ATT_B_502")
                    Response.Write("<CODE>" & sStatus & "</CODE>")
                    sStatus = "F" : Response.Write("<DESC>" & sMsg & "</DESC>")
                Case "5" '* 主檔未設定
                    Response.Write("<CODE>" & sStatus & "</CODE>")
                    sStatus = "F" : Response.Write("<DESC>" & sMsg & "</DESC>")
            End Select

            'Try
            '    sSql = "INSERT ChargeTeleGram (Actions,Result,TeleGram) VALUES('CHG_B_502','" & sStatus & "','" & sTradeData & "') "
            '    If sCmd IsNot Nothing Then
            '        sCmd.CommandText = sSql
            '        sCmd.ExecuteNonQuery()
            '    End If
            'Catch ex2 As Exception

            'End Try
            If sCmd IsNot Nothing Then sCmd.Dispose() : sCmd = Nothing
            If sConn IsNot Nothing Then sConn.Close() : sConn = Nothing
            Response.End()
        End Try
    End Sub



End Class
