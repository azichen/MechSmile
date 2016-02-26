using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.Configuration;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using Microsoft.Reporting.WebForms;
using System.Data.SqlClient;

partial class ATT_Q_015 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        String sRol;  sRol = "";
       
        if (!(Page.IsPostBack))
        {
            string USER_ROL = (string)ViewState["ROL"];
            string USER_ROLTYPE = (string)ViewState["RolType"];
            string USER_ROLDEPTID = (string)ViewState["RolDeptID"];

            if (!User.Identity.IsAuthenticated)
                FormsAuthentication.RedirectToLoginPage();
            else //'-----------------------------------* 檢查是否有 讀取/新增 權限
            {
                modUtil.GetRolData(Request, ref USER_ROL, ref USER_ROLTYPE, ref USER_ROLDEPTID);
                if (USER_ROL.Substring(0, 1) == "N") FormsAuthentication.RedirectToLoginPage();
            }

            if (User.Identity.IsAuthenticated)
            {
                modUtil.GetRolData(Request, ref USER_ROL, ref USER_ROLTYPE, ref USER_ROLDEPTID);
                this.txtStnName.Attributes["readonly"] = "readonly";
                this.txtStnName.Attributes["Class"] = "readonly";
                this.txtEmpName.Attributes["readonly"] = "readonly";
                this.txtEmpName.Attributes["Class"] = "readonly";
                //
                if (USER_ROL.Substring(0, 1) == "N")
                    FormsAuthentication.RedirectToLoginPage();
                else
                    if (modUtil.IsStnRol(this.Request) || modUnset.IsPAUnitRol(this.Request))  //'* 加油站權限
                    {   //
                        this.txtStnID.Text = HttpUtility.UrlDecode(Request.Cookies["STNID"].Value);
                        this.txtStnName.Text = HttpUtility.UrlDecode(Request.Cookies["STNNAME"].Value);
                        this.btnStnID.Visible = false;
                        
                        this.txtStnID.Attributes["readonly"] = "readonly";
                        this.txtStnID.Attributes["Class"] = "readonly";
                    }
                    else
                    {
                        this.txtStnID.Text = "";
                        this.txtStnName.Text = "";
                    };

                //----------------------------------------* 設定GridView樣式
                modDB.SetGridViewStyle(this.grdList, 20);
                this.grdList.RowStyle.Height = 20;

                modDB.SetFields("STNNAME"   , "站別", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("EMPID", "員工職號", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("EMPNAME" , "員工姓名", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("SHSTDATE", "出勤日期", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("WEEKDAYN", "星期", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("SHSTTIME", "起始時間", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("SHEDTIME", "結束時間", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("NETWH", "淨工時", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("DAYOVER", "單日超時", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("TWKOVER", "雙周超時", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("OVERHOUR", "加班小計", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("FOVER", "前2H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("BOVER", "後2H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("NFOVER", "夜前2H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("NBOVER", "夜後2H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("HOLIMARK", "假日註記", this.grdList, HorizontalAlign.Center, "", false, 0, false);                

                //'----------------------------------------* 設定欄位屬性
                txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()");

                modUtil.SetDateObj(txtDtFrom, false, txtDtTo, false, null, null);
                modUtil.SetDateObj(txtDtTo, false, null, false, null, null);
                modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtFrom, true, false, this.txtDtTo, null, null, false);
                modUtil.SetDateImgObj(this.imgDtTo, this.txtDtTo, true, false, null, null, null, false);

                //'----------------------------------------* 重新顯示查詢結果
                this.lblMsg.Text = Request["msg"];
                if (Session["QryField"] == null)
                {
                    DataTable VSEmptyTable = (DataTable)this.ViewState["EmptyTable"];
                    modDB.ShowEmptyDataGridHeader(this.grdList, 0,15, ref VSEmptyTable);
                }

                //'* 業務處/責任區/加油站之畫面設定控制
                //modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, (string)this.ViewState["ROL"], (string)this.ViewState["RolDeptID"]);
                modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, USER_ROL, USER_ROLDEPTID);
                if (USER_ROLTYPE == "STN")                                                                                
                    modUtil.SetObjReadOnly(form1, "txtStnID", false);

            }
        }

        string sqlstr = (string)ViewState["Sql"];

        //this.dscList.SelectCommand = (string)ViewState["Sql"];
        this.dscList.SelectCommand = sqlstr;
        this.ScriptManager1.RegisterPostBackControl(this.btnExcel);

    }


    //******************************************************************************************************
    //* 重新顯示查詢結果 
    //******************************************************************************************************
    protected void ReflashQryData(object sender)
    {
        Hashtable sCol = new Hashtable();
        sCol = (Hashtable)Session["QryField"];

        //string[] sKey = new string[2];
        //string[] sVal = new string[2];
        int scolcnt =0;

        if (sCol != null)
        {            scolcnt = (int)sCol.Count;

            string[] sKey = new string[scolcnt];
            string[] sVal = new string[scolcnt];

            sCol.Keys.CopyTo(sKey, 0);
            sCol.Values.CopyTo(sVal, 0);

            //azitemp 20130306
            for (int I = 0; I <= sCol.Count - 1; I++)
            {
                if (sKey[I].Substring(1, 3) == "ddl")
                    ((DropDownList)this.form1.FindControl(sKey[I])).SelectedValue = sVal[I];
                else
                    ((TextBox)this.form1.FindControl(sKey[I])).Text = sVal[I];
            }
            ShowGrid();
            grdList.PageIndex = (int)Session["iCurPage"];
            this.ViewState["iCurPage"] = Session["iCurPage"];
        }
        
        Session["QryField"] = null;
        Session["iCurPage"] = null;
    }


    //******************************************************************************************************
    //* 回主畫面時，重新顯示查詢結果
    //******************************************************************************************************
    protected void ddlArea_DataBound(object sender,EventArgs e) 
    {
        if (!(Page.IsPostBack))
        {
           if (Session["QryField"] == null)
                ReflashQryData(sender);
        }
    }

    //* 選取業務處/責任區時，清除加油站
    //******************************************************************************************************
    protected void ddlBus_SelectedIndexChanged(object sender,EventArgs e)
    {
        if (this.ViewState["RolType"] != "STN") 
        {
            this.txtStnID.Text = "";
            this.txtStnName.Text = "";
        }
    }

    //******************************************************************************************************
    //* 選取業務處時，先清除舊責任區，再附加
    //******************************************************************************************************
    protected void dscArea_Selecting(object sender,System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs e) 
    {
        string USER_ROL = (string)ViewState["ROL"];

        if ((string)ViewState["ROL"] == null) USER_ROL = "      ";

        if ((! IsPostBack) && (USER_ROL.Substring(5, 1).CompareTo("1") > 0))  //* 0:總 1:業 2:區 3,4:站
              e.Cancel = true;  //* 已依權限之設定畫面，不可重新讀取
        else
        {
            this.ddlArea.Items.Clear();
            this.ddlArea.Items.Add("");
        }
    }

    //******************************************************************************************************
    //* 顯示加油站名稱(依所選取之業務處/責任區過濾)
    //******************************************************************************************************

    protected void txtStnID_TextChanged(object sender,EventArgs e)
    {
        if (this.ddlArea.SelectedValue.Trim() == "")
        //if Strings.Trim(this.ddlArea.SelectedValue) == "")
            modCharge.GetStnName(txtStnID, this.txtStnName, this.ddlBus.SelectedValue);
        else
            modCharge.GetStnName(txtStnID, this.txtStnName, this.ddlArea.SelectedValue);
        //modCharge.GetStnName(txtStnID, this.txtStnName, (string.IsNullOrEmpty(Strings.Trim(@this.ddlArea.SelectedValue)) ? @this.ddlBus.SelectedValue : @this.ddlArea.SelectedValue));
    }

    //******************************************************************************************************
    //* 顯示頁碼
    //******************************************************************************************************
    protected void grdList_DataBound(object sender,EventArgs e)
    {
        modDB.SetGridPageNum(this.grdList, PagerButtons.NumericFirstLast, HorizontalAlign.Left);
    }

    //******************************************************************************************************
    //* 保存瀏覽頁碼
    //******************************************************************************************************
    protected void grdList_PageIndexChanged(object sender,EventArgs e)
    {
        this.ViewState["iCurPage"] = grdList.PageIndex;
    }

    //******************************************************************************************************
    //* 游標移動時，顯示光筆效果
    //******************************************************************************************************
    protected void grdList_RowDataBound(object sender, System.Web.UI.WebControls.GridViewRowEventArgs e)
    {
        modDB.SetGridLightPen(e);
    }

    //******************************************************************************************************
    //* 選擇加油站(依所選取之業務處/責任區過濾)
    //******************************************************************************************************
    protected void btnStnID_Click(object sender, EventArgs e)
    {
        //this.QryStn.Show((string)ViewState["RolType"], (string)ViewState["RolDeptID"], false);
        this.QryStn.ShowBySel(this.ddlBus.SelectedValue, this.ddlArea.SelectedValue);
    }

    //******************************************************************************************************
    //* 清除
    //******************************************************************************************************
    protected void btnClear_Click(object sender, EventArgs e)
    {
        Session["QryField"]= null;
        Session["iCurPage"] = null;
        Response.Redirect("ATT_Q_015.aspx");
    }

    //******************************************************************************************************
    //* Excel 匯出
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ015.xls", this.grdList);
    }

    protected void btnQry_Click(object sender,EventArgs e)
    {
        string sMsg = "";
        Boolean sValidOK = true;
        try
        {
            this.lblMsg.Text = "";
            sValidOK = true;
            this.txtStnID.Text = this.txtStnID.Text.Trim();
            //----------------------------------------* 檢核欄位正確性
            //sValidOK = modUtil.Check2DateObj(this.Control, "txtDtFrom", "txtDtTo", sMsg, "查詢日期");
            sValidOK = modUtil.Check2DateObj(this, "txtDtFrom", "txtDtTo",ref sMsg, "查詢日期");
            //sValidOK = modUtil.Check2TextObj("txtDtFrom", "txtDtTo", sMsg, "查詢日期");

            if ((this.txtStnID.Text == "") || (this.txtStnName.Text.Trim() == "")) 
            {
                sValidOK = false;
                sMsg = sMsg + "\n查無此加油站！";
            }
            if (!( sValidOK)) 
            {
                modUtil.showMsg(this.Page, "訊息", sMsg) ;
                return;
            }
            ShowGrid(); //* 顯示查詢結果
            Session["iCurPage"] = null;
        }
        catch (InvalidCastException ex)
        {
            modUtil.showMsg(this.Page, "錯誤訊息(查詢)", ex.Message);
        }
    }

    //******************************************************************************************************
    //* 顯示明細
    //******************************************************************************************************
    protected void  ShowGrid()
    {
        string sSQL ="";
        string DATEST = txtDtFrom.Text.Substring(0, 4) + txtDtFrom.Text.Substring(5, 2) + txtDtFrom.Text.Substring(8, 2);
        string DATEEN = txtDtTo.Text.Substring(0, 4) + txtDtTo.Text.Substring(5, 2) + txtDtTo.Text.Substring(8, 2);
        string SQLFILTER= " WHERE A.STNID='" + txtStnID.Text + "'"
                        + " AND SHEDDATE>='" + DATEST + "' AND SHEDDATE<='" + DATEEN + "'";
        if (txtEmpID.Text.Trim() != "")
            SQLFILTER = SQLFILTER + "   AND EMPID = '" + txtEmpID.Text + "'";
        //
        Hashtable sFields = new Hashtable();
        //
        sSQL = "SELECT ISNULL(STNNAME,'') AS STNNAME,EMPID,EMPNAME,SHSTDATE,WEEKDAYN,SHSTTIME,SHEDTIME"
             + ",NETWH,DAYOVER,TWKOVER,OVERHOUR,FOVER,BOVER,NFOVER,NBOVER,HOLIMARK FROM ("
             //
             + "SELECT EMPID AS GEMPID,'1' AS GID,STNID,A.EMPID,EMPL_NAME AS EMPNAME"
             + ",SUBSTRING(SHSTDATE,1,4)+'-'+SUBSTRING(SHSTDATE,5,2)+'-'+SUBSTRING(SHSTDATE,7,2) AS SHSTDATE"
             + ",CASE DATEPART(WEEKDAY,convert(Datetime, SHSTDATE, 112)) "
             + " WHEN 1 THEN '日' WHEN 2 THEN '一' WHEN 3 THEN '二' WHEN 4 THEN '三'"
             + " WHEN 5 THEN '四' WHEN 6 THEN '五' WHEN 7 THEN '六' END AS WEEKDAYN"
             + ",SUBSTRING(SHSTTIME,1,2)+':'+SUBSTRING(SHSTTIME,3,2) AS SHSTTIME"
             + ",SUBSTRING(SHEDTIME,1,2)+':'+SUBSTRING(SHEDTIME,3,2) AS SHEDTIME"
             + ",NETWH,DAYOVER,TWKOVER,DAYOVER+TWKOVER AS OVERHOUR"
             + ",CASE WHEN (DAYOVER+TWKOVER)>=2 THEN 2 ELSE (DAYOVER+TWKOVER) END AS FOVER"
             + ",CASE WHEN (DAYOVER+TWKOVER)>2 THEN (DAYOVER+TWKOVER)-2 ELSE 0 END AS BOVER"
             + ",CASE WHEN OVTM_P_OVHOURS>=2 THEN 2 ELSE ISNULL(OVTM_P_OVHOURS,0) END AS NFOVER"
             + ",CASE WHEN OVTM_P_OVHOURS>2 THEN OVTM_P_OVHOURS-2 ELSE 0 END AS NBOVER"
             + ",CASE WHEN Convert(datetime, SHEDDATE, 112) IN (SELECT HOLI_DATE FROM MP_HR.DBO.HOLIDAY) THEN '*' ELSE '' END AS HOLIMARK"
             + "  FROM SCHEDM A "
             + " INNER JOIN MP_HR.DBO.EMPLOYEE C ON A.EMPID=C.EMPL_ID"
             + "  LEFT JOIN MP_HR.DBO.OVERTIME D ON A.OVTMSEQ>0 AND A.EMPID=D.OVTM_EMPLID AND A.OVTMSEQ=D.OVTM_SEQ "
             + SQLFILTER;
        //
        sSQL = sSQL + " UNION ALL "
             + "SELECT EMPID AS GEMPID,'2' AS GID,STNID,A.EMPID,EMPL_NAME AS EMPNAME,'小計' AS SHSTDATE"
             + ",'' AS WEEKDAYN,'' AS SHSTTIME,'' AS SHEDTIME,SUM(NETWH) AS NETWH"
             + ",SUM(DAYOVER) AS DAYOVER,SUM(TWKOVER) AS TWKOVER,SUM(DAYOVER+TWKOVER) AS OVERHOUR"
             + ",SUM(CASE WHEN (DAYOVER+ISNULL(TWKOVER,0))>=2 THEN 2 ELSE (DAYOVER+ISNULL(TWKOVER,0)) END) AS FOVER"
             + ",SUM(CASE WHEN (DAYOVER+ISNULL(TWKOVER,0))>2 THEN (DAYOVER+ISNULL(TWKOVER,0))-2 ELSE 0 END) AS BOVER"
             + ",SUM(CASE WHEN OVTM_P_OVHOURS>=2 THEN 2 ELSE ISNULL(OVTM_P_OVHOURS,0) END) AS NFOVER"
             + ",SUM(CASE WHEN OVTM_P_OVHOURS>2 THEN OVTM_P_OVHOURS-2 ELSE 0 END) AS NBOVER,'' AS HOLIMARK"
             + "  FROM SCHEDM A "
             + " INNER JOIN MP_HR.DBO.EMPLOYEE C ON A.EMPID=C.EMPL_ID"
             + "  LEFT JOIN MP_HR.DBO.OVERTIME D ON A.OVTMSEQ>0 AND A.EMPID=D.OVTM_EMPLID AND A.OVTMSEQ=D.OVTM_SEQ "
             + SQLFILTER + " GROUP BY A.STNID,A.EMPID,EMPL_NAME";
        //
        sSQL = sSQL + " UNION ALL "
             + "SELECT 'zzzzzzzz' AS GEMPID,'3' AS GID,'' AS STNID,'' AS EMPID,'' AS EMPNAME"
             + ",'' AS SHSTDATE,'' AS WEEKDAYN,'' AS SHSTTIME,'總計' AS SHEDTIME"
             + ",SUM(NETWH) AS NETWH,SUM(DAYOVER) AS DAYOVER,SUM(TWKOVER) AS TWKOVER,SUM(DAYOVER+TWKOVER) AS OVERHOUR"
             + ",SUM(CASE WHEN (DAYOVER+ISNULL(TWKOVER,0))>=2 THEN 2 ELSE (DAYOVER+ISNULL(TWKOVER,0)) END) AS FOVER"
             + ",SUM(CASE WHEN (DAYOVER+ISNULL(TWKOVER,0))>2 THEN (DAYOVER+ISNULL(TWKOVER,0))-2 ELSE 0 END) AS BOVER"
             + ",SUM(CASE WHEN OVTM_P_OVHOURS>=2 THEN 2 ELSE ISNULL(OVTM_P_OVHOURS,0) END) AS NFOVER"
             + ",SUM(CASE WHEN OVTM_P_OVHOURS>2 THEN OVTM_P_OVHOURS-2 ELSE 0 END) AS NBOVER,'' AS HOLIMARK"
             + " FROM SCHEDM A "
             + " INNER JOIN MP_HR.DBO.EMPLOYEE C ON A.EMPID=C.EMPL_ID"
             + "  LEFT JOIN MP_HR.DBO.OVERTIME D ON A.OVTMSEQ>0 AND A.EMPID=D.OVTM_EMPLID AND A.OVTMSEQ=D.OVTM_SEQ "
             + SQLFILTER +") A LEFT JOIN MECHSTNM B ON A.STNID=B.STNID "
             + " ORDER BY GEMPID,GID,SHSTDATE";
        //--------------------------------------------* 加入業務處/責任區/加油站之過濾
        sFields.Add("txtStnID", txtStnID.Text);
        sFields.Add("txtStnName", txtStnName.Text);
        sFields.Add("txtDtFrom", txtDtFrom.Text);
        sFields.Add("txtDtTo", txtDtTo.Text);
        sFields.Add("txtEmpID", txtEmpID.Text);

        this.ViewState["QryField"] = sFields;
        this.dscList.SelectCommand = sSQL;
        this.grdList.DataSourceID = this.dscList.ID;
        this.grdList.DataBind();

        if (this.grdList.Rows.Count <= 0)
        {
            System.Data.DataTable VSEmptyTable = (DataTable)ViewState["EmptyTable"];
            modDB.ShowEmptyDataGridHeader(this.grdList, 0, 15, ref VSEmptyTable);
            modUtil.showMsg(this.Page, "訊息", "查無資料！");
        }
        else
        {
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;      //* 用於 Excel/換頁
            //modUtil.showMsg(this.Page, "訊息", "ok-1！"); //azitest
        }
    }

    
    //protected void savelog(string VLOGSTR)
    //{
    //    //System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\20120805\VOIR102log.txt", System.IO.FileMode.OpenOrCreate);
    //    System.IO.FileStream fs = new System.IO.FileStream(@"C:\mech\201310log.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
    //    System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.Unicode);
    //    sw.WriteLine(DateTime.Now.ToString ("hh:mm:ss")+" "+VLOGSTR); //azitemp
    //    sw.Flush();
    //    sw.Close();
    //    fs.Close();
    //}

    //******************************************************************************************************
    //* 選擇員工
    //******************************************************************************************************
    protected void btnEmpID_Click(object sender, EventArgs e)
    {
        if (this.txtStnID.Text != "")
        {
            this.QryEmp.Show(this.txtStnID.Text);
        }
    }

    //******************************************************************************************************
    //* 顯示員工姓名
    //******************************************************************************************************

    protected void txtEmpID_TextChanged(object sender, EventArgs e)
    {
        //modCharge.GetStnName(txtStnID, this.txtStnName, this.ddlArea.SelectedValue);
        modCharge.GetEmpName(this.txtEmpID, this.txtEmpName, true);
    }

    protected void btnEmp_Click(object sender, EventArgs e)
    {
        this.QryEmp.Show(this.txtStnID.Text);
    }
}
