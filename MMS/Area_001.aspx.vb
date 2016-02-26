Imports System.Data

Partial Class MMS_Area_001
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            '檢查是否有登錄
            If (Not User.Identity.IsAuthenticated) Then
                FormsAuthentication.RedirectToLoginPage()
            Else
                Dim sFileName As String = Request.Path.Substring(Request.Path.LastIndexOf("/") + 1, Request.Path.Length - (Request.Path.LastIndexOf("/") + 1))
                Dim sRol As String = HttpUtility.UrlDecode(Request.Cookies("ROL").Value)
                ViewState("ROL") = modUtil.GetRolData(sFileName, sRol)
                If ViewState("ROL").Substring(0, 1) = "N" Then FormsAuthentication.RedirectToLoginPage()
            End If

            Select Case Request("FormMode")
                Case "add"
                    Me.FormView1.ChangeMode(FormViewMode.Insert)
                    Me.TitleLabel.Text = "區域資料 - 新增"
                    CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Focus()
                Case "edit"
                    Me.FormView1.ChangeMode(FormViewMode.Edit)
                    Me.TitleLabel.Text = "區域資料 - 修改"
                    Dim tmpsql As String
                    tmpsql = "select * from  MMSArea where AreaCode='" + Request("STNID").Trim + "'"
                    Me.SqlDataSource1.SelectCommand = tmpsql
                    Me.SqlDataSource1.DataBind()
                    If CType(FormView1.FindControl("EffectiveTextBox"), TextBox).Text = "N" Then
                        CType(FormView1.FindControl("Button5"), Button).Text = "取消停用"
                        CType(FormView1.FindControl("Button3"), Button).Visible = False

                        CType(FormView1.FindControl("AreaNameTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("TelNoTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("AddressTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("TaxUnitTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("ProfitNoTextBox"), TextBox).Enabled = False
                        CType(FormView1.FindControl("UnitCodeTextBox"), TextBox).Enabled = False
                    End If
                    CType(FormView1.FindControl("AreaNameTextBox"), TextBox).Focus()
                    disable_edit()
            End Select
            Try
                If ViewState("ROL").Substring(2, 1) = "N" Then CType(FormView1.FindControl("Button6"), TextBox).Visible = False
                If ViewState("ROL").Substring(3, 1) = "N" Then CType(FormView1.FindControl("Button5"), TextBox).Visible = False
            Catch ex As Exception

            End Try

        End If
    End Sub

    '檢查輸入資料
    Private Function CheckData_detail() As String
        CheckData_detail = ""
        Dim getAreaCode As String = CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Text
        Dim getAreaName As String = CType(FormView1.FindControl("AreaNameTextBox"), TextBox).Text
        Dim getTelNoText As String = CType(FormView1.FindControl("TelNoTextBox"), TextBox).Text
        Dim getFaxNoText As String = CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Text
        Dim getAddressText As String = CType(FormView1.FindControl("AddressTextBox"), TextBox).Text
        Dim getTaxUnitText As String = CType(FormView1.FindControl("TaxUnitTextBox"), TextBox).Text
        Dim getProfitNoText As String = CType(FormView1.FindControl("ProfitNoTextBox"), TextBox).Text
        Dim getUnitCodeText As String = CType(FormView1.FindControl("UnitCodeTextBox"), TextBox).Text

        If getAreaCode = "" Then
            CheckData_detail += "區域代號不可空白! \n"
        End If

        Dim sSql As String
        Dim sRow As Collection = New Collection
        sSql = "select * from MMSArea where AreaCode='" + getAreaCode + "'"
        sRow = modDB.GetRowData(sSql)
        If sRow.Count > 0 Then
            CheckData_detail += "區域代號已經被使用! \n"
        End If
    End Function

    '新增
    Protected Sub Button1_Click(sender As Object, e As System.EventArgs)
        Dim sMsg As String = ""
        Dim getMonSE As String = ""

        sMsg = CheckData_detail()

        If sMsg.Length > 0 Then
            modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & sMsg)
            Exit Sub
        End If
        Dim getAreaCode As String = CType(FormView1.FindControl("AreaCodeTextBox"), TextBox).Text
        Dim getAreaName As String = CType(FormView1.FindControl("AreaNameTextBox"), TextBox).Text
        Dim getTelNoText As String = CType(FormView1.FindControl("TelNoTextBox"), TextBox).Text
        Dim getFaxNoText As String = CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Text
        Dim getAddressText As String = CType(FormView1.FindControl("AddressTextBox"), TextBox).Text
        Dim getTaxUnitText As String = CType(FormView1.FindControl("TaxUnitTextBox"), TextBox).Text
        Dim getProfitNoText As String = CType(FormView1.FindControl("ProfitNoTextBox"), TextBox).Text
        Dim getUnitCodeText As String = CType(FormView1.FindControl("UnitCodeTextBox"), TextBox).Text
        Dim tmpsql As String

        tmpsql = "insert into MMSArea(AreaCode,AreaName,TelNo,FaxNo,Address,TaxUnit,ProfitNo,UnitCode,Effective)"
        tmpsql += " values("
        tmpsql += "'" + getAreaCode + "',"
        tmpsql += "'" + getAreaName + "',"
        tmpsql += "'" + getTelNoText + "',"
        tmpsql += "'" + getFaxNoText + "',"
        tmpsql += "'" + getAddressText + "',"
        tmpsql += "'" + getTaxUnitText + "',"
        tmpsql += "'" + getProfitNoText + "',"
        tmpsql += "'" + getUnitCodeText + "',"
        tmpsql += "'Y')"
        Me.SqlDataSource1.InsertCommand = tmpsql
        Me.SqlDataSource1.Insert()
        Response.Redirect("Area.aspx?Returnflag=1&msg=資料新增成功!!")

    End Sub
    '修改刪除
    Protected Sub Button5_Click(sender As Object, e As System.EventArgs)
        Dim getAreaCode As String = CType(FormView1.FindControl("AreaCodeLabel1"), Label).Text
        Dim getTextBox1 As String = CType(FormView1.FindControl("TextBox1"), TextBox).Text
        Dim tmpsql, sMsg As String

        If CType(FormView1.FindControl("Button5"), Button).Text = "取消停用" Then
            tmpsql = "update MMSArea set "
            tmpsql += " Effective='Y'"
            tmpsql += " where AreaCode='" + getAreaCode + "'"
            Me.SqlDataSource1.UpdateCommand = tmpsql
            Me.SqlDataSource1.Update()
            Response.Redirect("Area.aspx?Returnflag=1&msg=該區域資料已取消停用!!")
        Else
            sMsg = CheckAreaChange()
            If sMsg.Length > 0 Then
                modUtil.showMsg(Me.Page, "訊息", "請修正下列異常： \n" & sMsg)
                Exit Sub
            End If

            If getTextBox1 = "Y" Then
                tmpsql = "update MMSArea set "
                tmpsql += " Effective='N'"
                tmpsql += " where AreaCode='" + getAreaCode + "'"
                Me.SqlDataSource1.UpdateCommand = tmpsql
                Me.SqlDataSource1.Update()
                Response.Redirect("Area.aspx?Returnflag=1&msg=該區域資料已使用過不可刪除,已將其停用!!")
            End If

            If getTextBox1 = "N" Then
                tmpsql = "delete MMSArea "
                tmpsql += " where AreaCode='" + getAreaCode + "'"
                Me.SqlDataSource1.DeleteCommand = tmpsql
                Me.SqlDataSource1.Delete()
                Response.Redirect("Area.aspx?Returnflag=1&msg=該區域資料已刪除!!")
            End If
        End If
    End Sub
    '修改更新
    Protected Sub Button3_Click(sender As Object, e As System.EventArgs)
        Dim getAreaCode As String = CType(FormView1.FindControl("AreaCodeLabel1"), Label).Text
        Dim getAreaName As String = CType(FormView1.FindControl("AreaNameTextBox"), TextBox).Text
        Dim getTelNoText As String = CType(FormView1.FindControl("TelNoTextBox"), TextBox).Text
        Dim getFaxNoText As String = CType(FormView1.FindControl("FaxNoTextBox"), TextBox).Text
        Dim getAddressText As String = CType(FormView1.FindControl("AddressTextBox"), TextBox).Text
        Dim getTaxUnitText As String = CType(FormView1.FindControl("TaxUnitTextBox"), TextBox).Text
        Dim getProfitNoText As String = CType(FormView1.FindControl("ProfitNoTextBox"), TextBox).Text
        Dim getUnitCodeText As String = CType(FormView1.FindControl("UnitCodeTextBox"), TextBox).Text
        Dim tmpsql As String
        tmpsql = "update MMSArea set "
        tmpsql += " AreaName='" + getAreaName + "',"
        tmpsql += " TelNo='" + getTelNoText + "',"
        tmpsql += " FaxNo='" + getFaxNoText + "',"
        tmpsql += " Address='" + getAddressText + "',"
        tmpsql += " TaxUnit='" + getTaxUnitText + "',"
        tmpsql += " ProfitNo='" + getProfitNoText + "',"
        tmpsql += " UnitCode='" + getUnitCodeText + "'"
        tmpsql += " where AreaCode='" + getAreaCode + "'"
        Me.SqlDataSource1.UpdateCommand = tmpsql
        Me.SqlDataSource1.Update()
        Response.Redirect("Area.aspx?Returnflag=1&msg=資料修改成功!!")
    End Sub

    '修改取消
    Protected Sub Button4_Click(sender As Object, e As System.EventArgs)
        Response.Redirect("Area.aspx?Returnflag=1&msg=")
    End Sub

    '新增取消
    Protected Sub Button2_Click(sender As Object, e As System.EventArgs)
        Response.Redirect("Area.aspx?Returnflag=1&msg=")
    End Sub



    '檢查身份變更
    Private Function CheckAreaChange() As String
        CheckAreaChange = ""
        Dim tmpsql As String
        tmpsql = "select * from MMSCustomers where AreaCode='" + CType(FormView1.FindControl("AreaCodeLabel1"), Label).Text + "'"
        tmpsql += " and Effective='Y'"
        If get_DataTable(tmpsql).Rows.Count > 0 Then
            CheckAreaChange = "此區域在客戶資料中仍被使用,修正後才可停用區域!!\n"
        End If
        tmpsql = "select * from MMSEmployee where AreaCode='" + CType(FormView1.FindControl("AreaCodeLabel1"), Label).Text + "'"
        tmpsql += " and Effective='Y'"
        If get_DataTable(tmpsql).Rows.Count > 0 Then
            CheckAreaChange += "此區域在員工資料中仍被使用,修正後才可停用區域!!\n"
            Exit Function
        End If

    End Function

#Region "DB function"
    '設定區域資料已使用
    Public Function SetAreaUsed(ByVal ss As String)
        Dim tmpsql As String
        tmpsql = "update MMSArea set Used='Y' where AreaCode='" + ss + "'"
        EXE_SQL(tmpsql)
    End Function

    '連接資料庫取得datatable
    Public Function get_DataTable(ByVal SQL1 As String) As System.Data.DataTable
        Dim CONNSTR1 As String = modDB.GetDataSource.ConnectionString
        Dim CONN1 As New System.Data.SqlClient.SqlConnection(CONNSTR1)
        Dim DATAADAPTER1 As System.Data.SqlClient.SqlDataAdapter
        Try
            DATAADAPTER1 = New System.Data.SqlClient.SqlDataAdapter(SQL1, CONN1)
            DATAADAPTER1.SelectCommand.CommandTimeout = 1600
            Dim MYDATASET1 As New System.Data.DataSet
            MYDATASET1.Reset()
            DATAADAPTER1.Fill(MYDATASET1)
            get_DataTable = MYDATASET1.Tables(0)
        Catch ex As Exception
            get_DataTable = Nothing
        End Try
    End Function

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
#End Region

    Protected Sub Button6_Click(sender As Object, e As System.EventArgs)
        enable_edit()
    End Sub

    Private Sub disable_edit()
        CType(FormView1.FindControl("AreaNameTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("TelNoTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("FaxNoTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("AddressTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("TaxUnitTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("ProfitNoTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("UnitCodeTextBox"), TextBox).ReadOnly = True
        CType(FormView1.FindControl("Button3"), Button).Visible = False
    End Sub


    Private Sub enable_edit()
        CType(FormView1.FindControl("AreaNameTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("TelNoTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("FaxNoTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("AddressTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("TaxUnitTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("ProfitNoTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("UnitCodeTextBox"), TextBox).ReadOnly = False
        CType(FormView1.FindControl("Button3"), Button).Visible = True
        CType(FormView1.FindControl("Button6"), Button).Visible = False
    End Sub
End Class
