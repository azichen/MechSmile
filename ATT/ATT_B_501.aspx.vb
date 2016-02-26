'******************************************************************************************************
'* 程式：ATT_B_501 氣候資訊電文
'* 作成：SMILE 劉遇龍
'* 版次：2014/01/17(VER1.01)：ASP.NET版新開發
'******************************************************************************************************

Imports System.Data.SqlClient

Partial Class Charge_ATT_B_501
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
                pStatus = "1" : Return False '* 電文欄位錯誤
            Else
                Return True
            End If
        Catch ex As Exception
            pStatus = "1" : Return False '* 電文欄位錯誤
        End Try
    End Function

    '******************************************************************************************************
    '* 取得上傳之交易資料，並下傳車籍相關資訊電文
    '******************************************************************************************************
    Private Sub GetTradeInfo()
        Dim sConn As SqlConnection = Nothing
        Dim sCmd As SqlCommand = Nothing
        Dim sReader As SqlDataReader = Nothing
        Dim sTrans As SqlTransaction = Nothing
        Dim sSql, sStatus, sTradeData As String

        sTradeData = "" : sStatus = ""
        Try
            sTradeData = Request("XMLData")
            'If cWriteLog Then modUtil.WriteLog(sTradeData, "B502")
            If sTradeData = "" Then sStatus = "1" : Exit Sub
            'sTradeData = "<WeatherCity><City>臺北市</City><Date>2014/01/23</Date><Code>8</Code><Text>晴時多雲</Text><City>基隆市</City><Date>2014/01/24</Date><Code>8</Code><Text>晴時多雲</Text></WeatherCity>"
			'sTradeData = "<WeatherCode><Code>2</Code><Text>多雲</Text><Code>8</Code><Text>晴時多雲</Text></WeatherCity>"

            sConn = modDB.GetSqlConnection()
            sCmd = New SqlCommand()
            sCmd.Connection = sConn

            '--------------------------------------------------------------* Step1. 解析欄位
            Dim sWeatherCity, sWeatherCode As String
            sWeatherCity = "" : sWeatherCode = ""
            GetFieldValue(sStatus, sTradeData, sWeatherCity, "WeatherCity", False)  '* 城市氣候
            GetFieldValue(sStatus, sTradeData, sWeatherCode, "WeatherCode", False)  '* 氣候代碼

            If sWeatherCity = "" And sWeatherCode = "" Then sStatus = "0" : Exit Sub '* 無資料匯入

            'If Mid(sDtAcc, 3, 1) = "/" Then '* (2008/07/12)修正POS日期之格式(POS:MM/dd/yyyy)
            '    sDtAcc = Mid(sDtAcc, 7, 4) & "/" & Mid(sDtAcc, 1, 5)
            'End If
            sTrans = sConn.BeginTransaction
            sCmd.Transaction = sTrans
            '--------------------------------------------------------------* Step2. 新增城市氣候(AttWeatherCity)
            'If Not GetFieldValue(sStatus, sTradeData, sRemain, "Remain", False) Then Exit Sub
            If sWeatherCity <> "" Then
                Dim sCity, sDate, sCode, sText As String
                sCity = "" : sDate = "" : sCode = "" : sText = ""
                'Dim I As Integer = 0
                Do While GetFieldValue(sStatus, sWeatherCity, sCity, "City", True)
                    If Not GetFieldValue(sStatus, sWeatherCity, sDate, "Date", True) Then Exit Sub
                    If Not GetFieldValue(sStatus, sWeatherCity, sCode, "Code", True) Then Exit Sub
                    If Not GetFieldValue(sStatus, sWeatherCity, sText, "Text", True) Then Exit Sub
                    'I += 1
                    'sSql = "INSERT INTO ChargeTradeItem (DtTrade, TmTrade, CarID, Seq, ItemID, ItemName, " _
                    '     & "Qty, ItemAmt, DtUpdate) Values ('" _
                    '     & sDtTrade & "','" & sTmTrade & "','" & sCarID & "'," & I & ",'" & sItemID & "',N'" & sItemName & "'," _
                    '     & sQty & "," & sAmt & ", getDate() )"
                    sSql = "IF exists (select 1 from AttWeatherCity where City='" & sCity & "' and CutDate='" & sDate & "')" _
                    & "	Update 	AttWeatherCity set Code=" & sCode & ",Name='" & sText & "' where City='" & sCity & "' and CutDate='" & sDate & "'" _
                    & " Else" _
                    & "	INSERT INTO AttWeatherCity (City, CutDate, Code, Name, DtUpdate) " _
                    & "Values ('" & sCity & "','" & sDate & "'," & sCode & ",'" & sText & "',getdate()) "
                    'If cWriteLog Then modUtil.WriteLog(sSql, "B502")
                    sCmd.CommandText = sSql
                    sCmd.ExecuteNonQuery()
                    sWeatherCity = Mid(sWeatherCity, cEnd + 7)
                Loop
                sStatus = "1"
            End If

            '--------------------------------------------------------------* Step3. 新增氣候代碼(AttWeatherCode)
            'If Not GetFieldValue(sStatus, sTradeData, sRemain, "Remain", False) Then Exit Sub
            cEnd = 0
            If sWeatherCode <> "" Then
                Dim wCode, wName, wAttCode As String
                wCode = "" : wName = "" : wAttCode = ""

                Do While GetFieldValue(sStatus, sWeatherCode, wCode, "Code", True)
                    If Not GetFieldValue(sStatus, sWeatherCode, wName, "Text", True) Then Exit Sub

                    'sSql = "INSERT INTO ChargeTradeItem (DtTrade, TmTrade, CarID, Seq, ItemID, ItemName, " _
                    '     & "Qty, ItemAmt, DtUpdate) Values ('" _
                    '     & sDtTrade & "','" & sTmTrade & "','" & sCarID & "'," & I & ",'" & sItemID & "',N'" & sItemName & "'," _
                    '     & sQty & "," & sAmt & ", getDate() )"
                    If InStr(wName, "雨") > 0 Then
                        wAttCode = "3"
                    ElseIf InStr(wName, "陰") > 0 Then
                        wAttCode = "2"
                    Else
                        wAttCode = "1"
                    End If
                    sSql = "Insert into AttWeatherCode (Code,Name,AttCode,DtUpdate) " _
                    & "select " & wCode & ",'" & wName & "','" & wAttCode & "',getdate() " _
                    & "where not exists (select 1 from AttWeatherCode where Code='" & wCode & "')"
                    'If cWriteLog Then modUtil.WriteLog(sSql, "B502")
                    sCmd.CommandText = sSql
                    sCmd.ExecuteNonQuery()

                    sWeatherCode = Mid(sWeatherCode, cEnd + 7)
                Loop

                sStatus = "1"

            End If

            sTrans.Commit()

        Catch ex As Exception
            sStatus = "2" '* 系統處理錯誤
            If sTrans IsNot Nothing Then sTrans.Rollback()
            If cWriteLog Then modUtil.WriteLog("Err=>" & ex.Message, "ATT_B501")

        Finally
            If sReader IsNot Nothing Then sReader.Close() : sReader = Nothing
            If cWriteLog Then modUtil.WriteLog("處理結果=>" & sStatus, "ATT_B501")
            Response.Clear()

            Response.Write(sStatus)
            'Select Case sStatus
            '    Case "0"
            '        Response.Write(sStatus)
            '        'sStatus = "F" : Response.Write("<DESC>無資料匯入</DESC>")
            '    Case "1" '* 電文欄位錯誤 => 寫入ErrLog後 紀錄成功
            '        modUtil.WriteLog("CHG_B_502 欄位錯誤=>" & sTradeData, "_Error_CHG_Tel")
            '        Response.Write(sStatus)
            '        sStatus = "S" : Response.Write("<DESC>成功</DESC>")
            '    Case "2" '* (電文)序號重複 => 視為成功
            '        Response.Write(sStatus)
            '        sStatus = "F" : Response.Write("<DESC>寫入失敗</DESC>")
            'End Select

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
