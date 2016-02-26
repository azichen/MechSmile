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
using System.Drawing;

//20130905_唯讀版；不做結算以免造成CONNECTION超載，日後以SQL PROCEDURE處理之。

partial class ATT_Q_005 : System.Web.UI.Page
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
            //----------------------------------------* 設定GridView樣式
            modDB.SetGridViewStyle(this.grdList, 20);
            this.grdList.RowStyle.Height = 20;
            modDB.SetFields("EMPID"   , "員工職號", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("EMPNAME" , "員工姓名", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("SHSTDATE", "排班日期", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("SHSTTIME", "排班時間", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("SHEDTIME", "結束時間", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FGSTTIME", "實際上班", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FGEDTIME", "實際下班", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FACWH", "實工時", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("NIGHT", "夜工時", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("NETWH", "淨工時", this.grdList, HorizontalAlign.Center, "", false, 0, false);

            modDB.SetFields("OVHOUR", "加班時", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FOVER", "前2H時", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("BOVER", "後2H時", this.grdList, HorizontalAlign.Center, "", false, 0, false);

            modDB.SetFields("VMHOUR", "請假時", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("MEMO" , "異常", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            //'--------------------------------------------------* 加入連結欄位

            CommandField sField = new CommandField();
            sField.ShowSelectButton = true;
            sField.HeaderText = "功能";
            sField.SelectText = "請假增修";
            sField.HeaderStyle.ForeColor = Color.White;
            sField.ItemStyle.HorizontalAlign = HorizontalAlign.Center;
            this.grdList.Columns.Add(sField);
            this.grdList.RowStyle.Height = 20;

            //'----------------------------------------* 設定欄位屬性
            txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()");
            modUtil.SetDateObj(txtDtFrom, false, txtDtTo, false, null, null);
            modUtil.SetDateObj(txtDtTo, false, null, false, null, null);
            modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtFrom, true, false, this.txtDtTo, null, null, false);
            modUtil.SetDateImgObj(this.imgDtTo, this.txtDtTo, true, false, null, null, null, false);

            //'----------------------------------------* 重新顯示查詢結果
            this.lblMsg.Text = Request["msg"];

            //'* 業務處/責任區/加油站之畫面設定控制
            //modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, (string)this.ViewState["ROL"], (string)this.ViewState["RolDeptID"]);
            this.txtEmpName.Attributes["readonly"] = "readonly";
            this.txtEmpName.Attributes["Class"] = "readonly";
            modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, USER_ROL, USER_ROLDEPTID);
            //if (USER_ROLTYPE == "STN")
            //    modUtil.SetObjReadOnly(form1,"txtStnID", false);
            if (modUtil.IsStnRol(this.Request) || modUnset.IsPAUnitRol(this.Request))  //'* 加油站權限
            {   //20110519
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

            if (Session["QryField"] == null)
            {
                DataTable VSEmptyTable = (DataTable)this.ViewState["EmptyTable"];
                modDB.ShowEmptyDataGridHeader(this.grdList, 0, 15, ref VSEmptyTable);
            }
            else
            {
                //modDB.InsertSignRecord("AziTest", "ATTQ005_PageLoad_ReflashQryData", User.Identity.Name); //azitest
                ReflashQryData(sender);
            }
        }
        //modDB.InsertSignRecord("AziTest", "ATTQ005_PageLoad_OK-5", User.Identity.Name); //azitest
        string sqlstr = (string)ViewState["Sql"];
        //this.dscList.SelectCommand = (string)ViewState["Sql"];
        this.dscList.SelectCommand = sqlstr;
        this.ScriptManager1.RegisterPostBackControl(this.btnExcel);
        //modDB.InsertSignRecord("AziTest", "ATTQ005_PageLoad_OK-6", User.Identity.Name); //azitest
    }

    //******************************************************************************************************
    //* 重新顯示查詢結果 
    //******************************************************************************************************
    protected void ReflashQryData(object sender)
    {
        //modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_ok-1", User.Identity.Name); //azitest
        Hashtable sCol = new Hashtable();
        sCol = (Hashtable)Session["QryField"];
        int scolcnt =0;
        if (sCol != null)
        {            
            scolcnt = (int)sCol.Count;
            string[] sKey = new string[scolcnt];
            string[] sVal = new string[scolcnt];

            sCol.Keys.CopyTo(sKey, 0);
            sCol.Values.CopyTo(sVal, 0);

            for (int I = 0; I <= sCol.Count - 1; I++)
            {
                //modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_i= " + I.ToString() + " " + sKey[I] + " = " + sVal[I], User.Identity.Name); //azitest
                if (sKey[I].Substring(1, 3) == "ddl")
                    ((DropDownList)this.form1.FindControl(sKey[I])).SelectedValue = sVal[I];
                else
                    ((TextBox)this.form1.FindControl(sKey[I])).Text = sVal[I];
            }
            //modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_ok-2", User.Identity.Name); //azitest
            ShowGrid();
            //modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_ok-2-2", User.Identity.Name); //azitest
            if (Session["iCurPage"] != null)
            {
                //modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_iCurPage is not null", User.Identity.Name); //azitest
                grdList.PageIndex = (int)Session["iCurPage"];
                this.ViewState["iCurPage"] = Session["iCurPage"];
            }
            //else
            //{
            //    modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_iCurPage is null", User.Identity.Name); //azitest
            //}
            //modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_ok-3", User.Identity.Name); //azitest
        }
        //modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_ok-4", User.Identity.Name); //azitest
        Session["QryField"] = null;
        Session["iCurPage"] = null;
        //modDB.InsertSignRecord("AziTest", "ATTQ005_ReflashQryData_ok-5", User.Identity.Name); //azitest
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
    //* 顯示員工姓名
    //******************************************************************************************************

    protected void txtEmpID_TextChanged(object sender, EventArgs e)
    {
         //modCharge.GetStnName(txtStnID, this.txtStnName, this.ddlArea.SelectedValue);
         modCharge.GetEmpName(txtEmpID, this.txtEmpName,true);
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
    //* 清除
    //******************************************************************************************************
    protected void btnClear_Click(object sender, EventArgs e)
    {
        Session["QryField"]= null;
        Session["iCurPage"] = null;
        Response.Redirect("ATT_Q_005.aspx");
    }

    //******************************************************************************************************
    //* Excel 匯出
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ005.xls", this.grdList);
    }

    protected void btnQry_Click(object sender,EventArgs e)
    {
        string sMsg = "";
        Boolean sValidOK = true;
        //要限制查單站+31天的資料
        try
        {
            sValidOK = true;
            this.txtStnID.Text = this.txtStnID.Text.Trim();
            sValidOK = modUtil.Check2DateObj(this, "txtDtFrom", "txtDtTo", ref sMsg, "查詢日期");
            //----------------------------------------* 檢核欄位正確性
            if ((this.txtStnID.Text == "") || (this.txtStnName.Text.Trim() == "")) 
            {
                sValidOK = false;
                sMsg = sMsg + "\n查無此加油站！";
            }
            if (sValidOK)
            {
                if ((txtDtFrom.Text.Trim() == "") || (txtDtTo.Text.Trim() == ""))
                {
                    sValidOK = false;
                    sMsg = sMsg + "\n 需輸入日期範圍!";
                }
            }

            if (sValidOK)
            {
                DateTime STime = DateTime.Parse(txtDtFrom.Text);
                DateTime ETime = DateTime.Parse(txtDtTo.Text);
                TimeSpan tsDay = ETime.Subtract(STime);
                int iDays = tsDay.Days + 1;

                if (iDays > 31) 
                {
                    sValidOK = false;
                    sMsg = sMsg + "\n 限制查詢31天內的資料，請重新輸入日期範圍!";
                }
            }

            if (!( sValidOK)) 
            {
                modUtil.showMsg(this.Page, "訊息", sMsg) ;
                return;
            }
            ShowGrid(); //* 顯示查詢結果
            Session["iCurPage"] = null;
            this.lblMsg.Text = "";
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
        //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_ok-1", User.Identity.Name); //azitest
        string DTSTSTR, DTEDSTR = "";
        //string LCKDT = modUnset.GET_LOCKYM("1", txtStnID.Text) + "21";
        string LCKDT = modUnset.GET_LOCKDT(txtStnID.Text);
        Boolean SCHECKFLAG = true;
        Hashtable sFields = new Hashtable();
        string sSQL, sSQL2 = "";
        string Q005DTST;

        //modDB.InsertSignRecord("AziTest", "OK-1" , User.Identity.Name); //azitest

        DTSTSTR = txtDtFrom.Text.Substring(0, 4) + txtDtFrom.Text.Substring(5, 2) + txtDtFrom.Text.Substring(8, 2);
        DTEDSTR = txtDtTo.Text.Substring(0, 4) + txtDtTo.Text.Substring(5, 2) + txtDtTo.Text.Substring(8, 2);
        //20150729 增加結轉日期起迄，只能結轉未鎖定的日期
        Q005DTST = DTSTSTR;
        if (string.Compare(Q005DTST, LCKDT) < 0)
        {
            Q005DTST = LCKDT;
        }
        //
        if (string.Compare(DTEDSTR, LCKDT) < 0)
        {
            SCHECKFLAG = false;
        }
        //modDB.InsertSignRecord("AziTest", "OK-2", User.Identity.Name); //azitest
        //if (string.Compare(DTEDSTR,LCKDT)<=0)
        //    SCHECKFLAG = false;
        //else if ((string.Compare(DTSTSTR, LCKDT) <= 0) && (string.Compare(DTEDSTR, LCKDT) > 0))
        //    DTSTSTR = LCKDT;
        //
        //modDB.InsertSignRecord("AziTest", "LCKDT=" + LCKDT + " DTSTSTR=" + DTSTSTR + " DTEDSTR=" + DTEDSTR, User.Identity.Name); //azitest
        //過了過帳日期，不可再對資料作異動。
        //modDB.InsertSignRecord("AziTest", "OK-3", User.Identity.Name); //azitest
        //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_ok-2", User.Identity.Name); //azitest
        
        if (SCHECKFLAG)
        {
            sSQL = "EXEC ATTQ005R '" + txtStnID.Text + "','" + Q005DTST + "','" + DTEDSTR + "','" + txtEmpID.Text + "','" + txtEmpID.Text + "'";
            modUnset.PA_ExcuteCommand(sSQL);
        }
        //
        string VFGSTTIME, VFGEDTIME="";
        double VFACWH = 0;
        double VNIGHT = 0;
        string VSHCHKSTR="";
        string SHDATEFILTER,FILTERSTR,SQLFILTER = "";
        //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_OK-3", User.Identity.Name); //azitest
        //

        if (string.Compare(Q005DTST, LCKDT) < 0)
        {
            Q005DTST = LCKDT;
        }

        SHDATEFILTER = "";
        if (string.Compare(DTSTSTR, "20150721") <= 0) 
        {
            DTSTSTR = String.Format("{0:yyyyMMdd}", DateTime.Parse(txtDtFrom.Text).AddDays(-1));
            SHDATEFILTER = " AND SHEDDATE+SHEDTIME>'" + DTSTSTR + "2300'";
        }
        else
        {
            SHDATEFILTER = " AND SHSTDATE>='" + DTSTSTR + "'";
        }

        if (string.Compare(DTEDSTR, "20150720") <= 0)  
        {
            SHDATEFILTER = SHDATEFILTER + " AND SHEDDATE+SHEDTIME<='" + DTEDSTR + "2300'";
        }
        else
        {
            SHDATEFILTER = SHDATEFILTER + " AND SHSTDATE<='" + DTEDSTR + "'";
        }

        //
        FILTERSTR = "    WHERE STNID= '" + txtStnID.Text + "'";
        if (txtEmpID.Text.Trim() != "")
        {
            FILTERSTR = FILTERSTR + "  AND EMPID='" + txtEmpID.Text + "'";
        }
        FILTERSTR = FILTERSTR + SHDATEFILTER;
                  //+ "      AND SHSTDATE >='" + DTSTSTR + "'"
                  //+ "      AND SHSTDATE <='" + DTEDSTR + "'";
        //
        SQLFILTER = "  FROM SCHEDM A WITH (NOLOCK) "
             + "  LEFT JOIN [MP_HR].DBO.VACATM V on A.EMPID=V.VATM_EMPLID "
             + "   AND (((SHSTDATE+SHSTTIME) BETWEEN "
             + "        (Convert(varchar(8),vatm_date_st,112)+vatm_time_st) and (Convert(varchar(8),vatm_date_en,112)+vatm_time_en)) "
             + "    OR  ((SHEDDATE+SHEDTIME) BETWEEN "
             + "        (Convert(varchar(8),vatm_date_st,112)+vatm_time_st) and (Convert(varchar(8),vatm_date_en,112)+vatm_time_en))) "
             + FILTERSTR;
        //     + " WHERE STNID= '" + txtStnID.Text + "'";
        //if (txtEmpID.Text.Trim() != "")
        //    {
        //        SQLFILTER = SQLFILTER + "  AND EMPID='" + txtEmpID.Text + "'";
        //    }
        
        //
        sSQL = "SELECT EMPID,EMPNAME,SHSTDATE,SHSTTIME,SHEDTIME,FGSTTIME,FGEDTIME,FACWH,NIGHT"
             + "      ,ISNULL(NETWH,0) AS NETWH,ISNULL(OVHOUR,0) AS OVHOUR,FOVER,BOVER,ISNULL(VMHOUR,0) AS VMHOUR,MEMO"
             + " FROM (";

        if (rdoType.SelectedValue == "1")
        {
            sSQL = sSQL
                 + " SELECT EMPID AS GEMPID,EMPID,'1' AS KINDID,CONVERT(VARCHAR(10),EMPL_NAME) AS EMPNAME,SHSTDATE,SHSTTIME,SHEDTIME"
                 + "       ,ISNULL(E.VASTR,'')+WKSTTIME AS FGSTTIME,ISNULL(F.VASTR,'')+WKEDTIME AS FGEDTIME,FACWH,NITEWH AS NIGHT"
                 + "       ,NETWH,(ISNULL(DAYOVER,0)+ISNULL(TWKOVER,0)) AS OVHOUR,SUM(VATM_HOURS) AS VMHOUR"
                //+ "  ,CASE WHEN (DAYOVER+TWKOVER)>=2 THEN 2 ELSE (DAYOVER+TWKOVER) END AS FOVER"
                //+ "  ,CASE WHEN (DAYOVER+TWKOVER)>2 THEN (DAYOVER+TWKOVER)-2 ELSE 0 END AS BOVER"
                 + "  ,CASE WHEN (DAYOVER+TWKOVER)>=2 THEN 2 ELSE (DAYOVER+TWKOVER) END AS FOVER"
                 + "  ,CASE WHEN (DAYOVER+TWKOVER)>2 THEN (DAYOVER+TWKOVER)-2 ELSE 0 END AS BOVER"
                 + "       ,CASE WHEN WORKHOUR=FACWH THEN '' ELSE '*' END AS MEMO"
                 + "   FROM SCHEDM A WITH (NOLOCK) "
                 + "  INNER JOIN [MP_HR].DBO.EMPLOYEE B ON A.EMPID=B.EMPL_ID "
                 + "   LEFT JOIN VACADEF E on A.WKSTMARK=E.VABYTE "
                 + "   LEFT JOIN VACADEF F on A.WKEDMARK=F.VABYTE "
                 + "   LEFT JOIN [MP_HR].DBO.VACATM V on A.EMPID=V.VATM_EMPLID "
                 + "         AND (((SHSTDATE+SHSTTIME) BETWEEN "
                 + " (Convert(varchar(8),vatm_date_st,112)+vatm_time_st) and (Convert(varchar(8),vatm_date_en,112)+vatm_time_en)) "
                 + "          OR  ((SHEDDATE+SHEDTIME) BETWEEN "
                 + " (Convert(varchar(8),vatm_date_st,112)+vatm_time_st) and (Convert(varchar(8),vatm_date_en,112)+vatm_time_en)))"
                 + FILTERSTR;
            //     + "  WHERE STNID= '" + txtStnID.Text + "'";
            //if (txtEmpID.Text.Trim() != "")
            //{
            //    sSQL = sSQL + "  AND EMPID='" + txtEmpID.Text + "'";
            //}
            //sSQL = sSQL + "   AND SHSTDATE >='" + DTSTSTR + "'"
            //            + "   AND SHSTDATE <='" + DTEDSTR + "'";
            sSQL = sSQL + " GROUP BY A.EMPID,EMPL_NAME,SHSTDATE,SHSTTIME,SHEDTIME,E.VASTR,F.VASTR"
                        + "         ,WKSTTIME,WKEDTIME,WORKHOUR,FACWH,NITEWH,NETWH,DAYOVER,TWKOVER";
            //20131021 ADD 增加小計
            sSQL = sSQL + " UNION ALL ";
        }
        sSQL = sSQL  
             + " SELECT EMPID AS GEMPID,EMPID,'2' AS KINDID,CONVERT(VARCHAR(10),EMPL_NAME) AS EMPNAME,'小計' AS SHSTDATE,'' AS SHSTTIME,'' AS SHEDTIME"
             + "       ,'' AS FGSTTIME,'' AS FGEDTIME,SUM(FACWH) AS FACWH,SUM(NIGHT) AS NIGHT"
             + "       ,SUM(NETWH) AS NETWH,SUM(DAYOVER+TWKOVER) AS OVHOUR,SUM(VMHOUR) AS VMHOUR"
             + "       ,SUM(FOVER) AS FOVER,SUM(BOVER) AS BOVER,'' AS MEMO FROM ("

             + "   SELECT EMPID,SUM(FACWH) AS FACWH,SUM(NITEWH) AS NIGHT,SUM(NETWH) AS NETWH,SUM(ISNULL(DAYOVER,0)+ISNULL(TWKOVER,0)) AS OVHOUR"
             + "         ,SUM(DAYOVER) AS DAYOVER,SUM(TWKOVER) AS TWKOVER"
             + "         ,SUM(CASE WHEN (DAYOVER+TWKOVER)>=2 THEN 2 ELSE (DAYOVER+TWKOVER) END) AS FOVER"
             + "         ,SUM(CASE WHEN (DAYOVER+TWKOVER)>2 THEN (DAYOVER+TWKOVER)-2 ELSE 0 END) AS BOVER,0 AS VMHOUR "
             + "     FROM SCHEDM A WITH (NOLOCK) " + FILTERSTR + " GROUP BY EMPID"

             + " UNION ALL "

             + " SELECT EMPID,0 AS FACWH,0 AS NIGHT,0 AS NETWH,0 AS OVHOUR"
             + "       ,0 AS DAYOVER,0 AS TWKOVER,0 AS FOVER,0 AS BOVER,ISNULL(SUM(VATM_HOURS),0) AS VMHOUR"
             + SQLFILTER + " GROUP BY EMPID) F "
             + "  INNER JOIN [MP_HR].DBO.EMPLOYEE B ON F.EMPID=B.EMPL_ID "
             + "  GROUP BY EMPID,EMPL_NAME ";
             //+ "  WHERE STNID= '" + txtStnID.Text + "'";
        //if (txtEmpID.Text.Trim() != "")
        //{
        //    sSQL = sSQL + "  AND EMPID='" + txtEmpID.Text + "'";
        //}
        //sSQL = sSQL + "    AND SHEDDATE+SHEDTIME  >'" + DTSTSTR + "2300'"
        //     + "    AND SHEDDATE+SHEDTIME <='" + DTEDSTR + "2300'"
        //   + "  GROUP BY EMPID,EMPL_NAME ";

        //20131024 ADD 增加總計
        sSQL = sSQL + " UNION ALL "
             + " SELECT 'zzzzzzzz' AS GEMPID,'' AS EMPID,'3' AS KINDID,'' AS EMPNAME,'' AS SHSTDATE,'' AS SHSTTIME,'' AS SHEDTIME"
             + "       ,'' AS FGSTTIME,'總計' AS FGEDTIME,SUM(FACWH) AS FACWH,SUM(NIGHT) AS NIGHT"
             + "       ,SUM(NETWH) AS NETWH,SUM(DAYOVER+TWKOVER) AS OVHOUR,SUM(VMHOUR) AS VMHOUR"
             + "       ,SUM(FOVER) AS FOVER,SUM(BOVER) AS BOVER,'' AS MEMO FROM ("
             + "   SELECT SUM(FACWH) AS FACWH,SUM(NITEWH) AS NIGHT,SUM(NETWH) AS NETWH,SUM(ISNULL(DAYOVER,0)+ISNULL(TWKOVER,0)) AS OVHOUR"
             + "         ,SUM(DAYOVER) AS DAYOVER,SUM(TWKOVER) AS TWKOVER"
             + "         ,SUM(CASE WHEN (DAYOVER+TWKOVER)>=2 THEN 2 ELSE (DAYOVER+TWKOVER) END) AS FOVER"
             + "         ,SUM(CASE WHEN (DAYOVER+TWKOVER)>2 THEN (DAYOVER+TWKOVER)-2 ELSE 0 END) AS BOVER,0 AS VMHOUR "
             + "     FROM SCHEDM A WITH (NOLOCK) " + FILTERSTR
             //+ "    WHERE STNID= '" + txtStnID.Text + "'"
             //+ "      AND SHEDDATE+SHEDTIME  >'" + DTSTSTR + "2300'"
             //+ "      AND SHEDDATE+SHEDTIME <='" + DTEDSTR + "2300'"
             + " UNION ALL "
             + " SELECT 0 AS FACWH,0 AS NIGHT,0 AS NETWH,0 AS OVHOUR"
             + "       ,0 AS DAYOVER,0 AS TWKOVER,0 AS FOVER,0 AS BOVER,ISNULL(SUM(VATM_HOURS),0) AS VMHOUR"
             + SQLFILTER + ") A) G"
             + "  ORDER BY GEMPID,KINDID,SHSTDATE,SHSTTIME";
        //--------------------------------------------* 加入業務處/責任區/加油站之過濾
        //savelog(sSQL);

        sFields.Add("txtStnID", txtStnID.Text);
        sFields.Add("txtStnName", txtStnName.Text);
        sFields.Add("txtDtFrom", txtDtFrom.Text);
        sFields.Add("txtDtTo", txtDtTo.Text);
        sFields.Add("txtEmpID", txtEmpID.Text);
        sFields.Add("txtEmpName", txtEmpName.Text);
        this.ViewState["QryField"] = sFields;
        this.dscList.SelectCommand = sSQL;
        this.grdList.DataSourceID = this.dscList.ID;
        this.grdList.DataBind();
        //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_OK_5", User.Identity.Name); //azitest

        //modUtil.showMsg(this.Page, sSQL); //azitest
        if (this.grdList.Rows.Count <= 0)
        {
            //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_OK_6-1", User.Identity.Name); //azitest
            System.Data.DataTable VSEmptyTable = (DataTable)ViewState["EmptyTable"];
            modDB.ShowEmptyDataGridHeader(this.grdList, 0, 15, ref VSEmptyTable);
            modUtil.showMsg(this.Page, "訊息", "查無資料！");
            //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_OK_6-2", User.Identity.Name); //azitest
        }
        else
        {
            //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_OK_6-3", User.Identity.Name); //azitest
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;      //* 用於 Excel/換頁
            //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_OK_6-4", User.Identity.Name); //azitest
        }
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

    //'******************************************************************************************************
    //'* 執行 瀏覽操作 
    //'******************************************************************************************************
    protected void grdList_SelectedIndexChanged(object sender,EventArgs e)
    {
        if (string.Compare(grdList.SelectedRow.Cells[3].Text.Trim(),"")>0)
        {
            string Dtfrom, DtTo = "";
            //modDB.InsertSignRecord("AziTest", "ATT_Q005_SelectedIndexChanged", User.Identity.Name); //azitest
            Hashtable sFields = new Hashtable();
            Session["QryField"] = this.ViewState["QryField"];
            sFields.Add("txtEMPID", grdList.SelectedRow.Cells[0].Text);
            sFields.Add("txtEMPNAME", grdList.SelectedRow.Cells[1].Text);
            Dtfrom = DateTime.ParseExact(grdList.SelectedRow.Cells[2].Text, "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces).ToString("yyyy/MM/dd");
            if (string.Compare(grdList.SelectedRow.Cells[3].Text, grdList.SelectedRow.Cells[4].Text) > 0)
            {
                DtTo = DateTime.ParseExact(grdList.SelectedRow.Cells[2].Text, "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces).AddDays(1).ToString("yyyy/MM/dd");
            }
            else
            {
                DtTo = Dtfrom;
            }
            //Dtfrom = "2015/03/01";
            //DtTo = "2015/03/01";
            sFields.Add("txtDtFrom", Dtfrom);
            sFields.Add("txtDtTo", DtTo);
            Session["QryField_052B"] = sFields;
            Session["iCurPage"] = grdList.PageIndex;
            //this.ViewState["iCurPage"] = grdList.PageIndex;
            //
            //Response.Redirect("ATT_Q_052B.aspx?EMPID=" + grdList.SelectedRow.Cells[0].Text + "&EMPNAME=" + grdList.SelectedRow.Cells[1].Text + "&SHSTDATE=" + grdList.SelectedRow.Cells[2].Text + "&SHSTTIME=" + grdList.SelectedRow.Cells[3].Text + "&SHEDTIME=" + grdList.SelectedRow.Cells[4].Text);
            Response.Redirect("ATT_Q_052B.aspx");
        }
        else
        {
            modUtil.showMsg(this.Page, "訊息", "資料錯誤！");
        }
    }

    //protected void savelog(string VLOGSTR)
    //{
    //    //System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\20120805\VOIR102log.txt", System.IO.FileMode.OpenOrCreate);
    //    System.IO.FileStream fs = new System.IO.FileStream(@"C:\mech\201508log.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
    //    System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.Unicode);
    //    sw.WriteLine(DateTime.Now.ToString("hh:mm:ss") + " " + VLOGSTR); //azitemp
    //    sw.Flush();
    //    sw.Close();
    //    fs.Close();
    //}

}
