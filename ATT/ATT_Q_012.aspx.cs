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

partial class ATT_Q_012 : System.Web.UI.Page
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
            modDB.SetFields("STNNAME", "站名稱"  , this.grdList, HorizontalAlign.Center, "", false, 0, false);

            modDB.SetFields("EMPNAME", "員工姓名", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("AWDWH1" , "第一周可報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("AWDWH2", "第二周可報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("AWDWH3", "第29日可報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("AWDWH4", "第30日可報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("AWDWH5", "第31日可報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);

            modDB.SetFields("OVERHOUR1", "第一周已報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVERHOUR2", "第二周已報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVERHOUR3", "第29日已報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVERHOUR4", "第30日已報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVERHOUR5", "第31日已報時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            //'----------------------------------------* 設定欄位屬性
            txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()");

            modUtil.SetDateObj(txtDtStart, false, null, true, null, null);

            //'----------------------------------------* 重新顯示查詢結果
            this.lblMsg.Text = Request["msg"];

            if (Session["QryField"] == null)
            {
                DataTable VSEmptyTable = (DataTable)this.ViewState["EmptyTable"];
                modDB.ShowEmptyDataGridHeader(this.grdList, 0,11, ref VSEmptyTable);
            }
            //'* 業務處/責任區/加油站之畫面設定控制
            //modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, (string)this.ViewState["ROL"], (string)this.ViewState["RolDeptID"]);
            modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, USER_ROL, USER_ROLDEPTID);
            if (USER_ROLTYPE == "STN")
                modUtil.SetObjReadOnly(form1,"txtStnID", false); 
            //}
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
        Response.Redirect("ATT_Q_012.aspx");
    }

    //******************************************************************************************************
    //* Excel 匯出
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ012.xls", this.grdList);
    }

    protected void btnQry_Click(object sender,EventArgs e)
    {
        string sMsg = "";
        Boolean sValidOK = true;
        //this.lblMsg2.Text = "查詢!1234567890"; //azitest
        //this.TextBox1.Text = "Test";
        //this.TextBox1.Text = this.TextBox1.Text+"Test";
        //lblMsg.Text = "test!!!";
        //modUtil.showMsg(this.Page, "訊息", "Test");
        //return;
        //要限制查單站+31天的資料
        try
        {
            sValidOK = true;
            this.txtStnID.Text = this.txtStnID.Text.Trim();
            //sValidOK = modUtil.Check2DateObj(this, "txtDtFrom", "txtDtTo", ref sMsg, "查詢日期");
            sValidOK = modUtil.CheckNotEmpty(this.txtDtStart, ref sMsg, "查詢月份", true);
            //----------------------------------------* 檢核欄位正確性
            if ((this.txtStnID.Text == "") || (this.txtStnName.Text.Trim() == "")) 
            {
                sValidOK = false;
                sMsg = sMsg + "\n查無此加油站！";
            }
            if (sValidOK)
            {
                if (txtDtStart.Text.Trim() == "") 
                {
                    sValidOK = false;
                    sMsg = sMsg + "\n 需輸入月份日期!";
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
    protected void ShowGrid()
    {
        DateTime ETime = DateTime.Parse(txtDtStart.Text + "/20");
        DateTime STime = ETime.AddMonths(-1).AddDays(1);
        TimeSpan tsDay = ETime.Subtract(STime);
        //int STANDAY = tsDay.Days + 1;
        //int STANDHR = (STANDAY-7) * 8;
        //
        string DTST, DT01, DT02, DT03, DT04, DT05, DTED = "";
        //string DWST, DW01, DW02, DW03, DW04, DW05, DWED = "";
        //DTST = txtDtStart.Text.Substring(0, 4) + txtDtStart.Text.Substring(5, 2) + "20";
        DTST = STime.ToString("yyyyMM") + "20";
        //DT01 = DateTime.Parse(txtDtStart.Text + "/20").AddDays(14).ToString("yyyyMMdd"); 
        //DT02 = DateTime.Parse(txtDtStart.Text + "/20").AddDays(28).ToString("yyyyMMdd"); 
        DT01 = ETime.AddMonths(-1).AddDays(14).ToString("yyyyMMdd");
        DT02 = ETime.AddMonths(-1).AddDays(28).ToString("yyyyMMdd"); 

        //DWST = DateTime.Parse(txtDtStart.Text + "/20").AddDays(-1).ToString("yyyymmdd") + "2300";
        //DW01 = DateTime.Parse(DT01.Substring(0, 4) + "/" + DT01.Substring(4, 2) + "/" + DT01.Substring(6, 2)).AddDays(-1).ToString("yyyymmdd") + "2300";
        //DW02 = DateTime.Parse(DT02.Substring(0, 4) + "/" + DT02.Substring(4, 2) + "/" + DT02.Substring(6, 2)).AddDays(-1).ToString("yyyymmdd") + "2300";
        //DWST = DTST + "2300";
        //DW01 = DT01 + "2300";
        //DW02 = DT02 + "2300";
        DT03 = "";
        DT04 = "";
        DT05 = "";
        string DtMsg = "備註 第一周:" + STime.ToString("MM/dd") + "~" + ETime.AddMonths(-1).AddDays(14).ToString("MM/dd")
                     + " 第二周:" + ETime.AddMonths(-1).AddDays(15).ToString("MM/dd") + "~" + ETime.AddMonths(-1).AddDays(28).ToString("MM/dd");
        
        if (string.Compare(DT02.Substring(6, 2),"20")<0)
        {
            //modDB.InsertSignRecord("AziTest", "ok-1 DTST=" + DTST + " DT01=" + DT01 + " DT02=" + DT02 + " DT03=" + DT03 + " DT04=" + DT04 + " DT05=" + DT05, User.Identity.Name); //AZITEST
            DT03 = DateTime.Parse(DT02.Substring(0, 4) + "/" + DT02.Substring(4, 2) + "/" + DT02.Substring(6, 2)).AddDays(1).ToString("yyyyMMdd");
            //modDB.InsertSignRecord("AziTest", "ok-2 DTST=" + DTST + " DT01=" + DT01 + " DT02=" + DT02 + " DT03=" + DT03 + " DT04=" + DT04 + " DT05=" + DT05, User.Identity.Name); //AZITEST
            //DW03 = DateTime.Parse(DT03.Substring(0, 4) + "/" + DT03.Substring(4, 2) + "/" + DT03.Substring(6, 2)).AddDays(-1).ToString("yyyymmdd") + "2300";
            //DW03 = DT03 + "2300";
            DtMsg = DtMsg + " 第29日:" + ETime.AddMonths(-1).AddDays(29).ToString("MM/dd");
            if (string.Compare(DT03.Substring(6, 2),"20")<0)
            {
                DT04 = DateTime.Parse(DT03.Substring(0, 4) + "/" + DT03.Substring(4, 2) + "/" + DT03.Substring(6, 2)).AddDays(1).ToString("yyyyMMdd");
                //DW04 = DateTime.Parse(DT04.Substring(0, 4) + "/" + DT04.Substring(4, 2) + "/" + DT04.Substring(6, 2)).AddDays(-1).ToString("yyyyMMdd") + "2300";
                //modDB.InsertSignRecord("AziTest", "ok-3 DTST=" + DTST + " DT01=" + DT01 + " DT02=" + DT02 + " DT03=" + DT03 + " DT04=" + DT04 + " DT05=" + DT05, User.Identity.Name); //AZITEST
                DtMsg = DtMsg + " 第30日:" + ETime.AddMonths(-1).AddDays(30).ToString("MM/dd");
                if (string.Compare(DT04.Substring(6, 2),"20")<0)
                {
                    DT05 = DateTime.Parse(DT04.Substring(0, 4) + "/" + DT04.Substring(4, 2) + "/" + DT04.Substring(6, 2)).AddDays(1).ToString("yyyyMMdd");
                    //DW05 = DateTime.Parse(DT05.Substring(0, 4) + "/" + DT05.Substring(4, 2) + "/" + DT05.Substring(6, 2)).AddDays(-1).ToString("yyyyMMdd") + "2300";
                    DtMsg = DtMsg + " 第31日:" + ETime.AddMonths(-1).AddDays(31).ToString("MM/dd");
                }
            }
        }
        this.lblMsg2.Text = DtMsg;
        //
        //modDB.InsertSignRecord("AziTest", "DTST=" + DTST + " DT01=" + DT01 + " DT02=" + DT02 + " DT03=" + DT03 + " DT04=" + DT04 + " DT05=" + DT05, User.Identity.Name); //AZITEST
        string VACCYM = "";
        //
        Hashtable sFields = new Hashtable();
        string sSQL, sSQL2 = "";
        //
        //VACCYM = txtDtStart.Text.Substring(0, 4) + txtDtStart.Text.Substring(5, 2);
        //
        //modDB.InsertSignRecord("AziTest", "VACCYM=" + VACCYM+" STNID=" + txtStnID.Text , User.Identity.Name); //AZITEST
        //modUtil.showMsg(this.Page, "訊息", "資料完成！");
        //modDB.InsertSignRecord("AziTest", "DTST=" + DTST + " DT01=" + DT01 + " DT02=" + DT02 + " DT03=" + DT03 + " DT04=" + DT04 + " DT05=" + DT05, User.Identity.Name); //AZITEST

        sSQL = "SELECT STNNAME,EMPID,EMPL_NAME AS EMPNAME,SUM(AWDWH1) AS AWDWH1,SUM(AWDWH2) AS AWDWH2,SUM(AWDWH3) AS AWDWH3,SUM(AWDWH4) AS AWDWH4,SUM(AWDWH5) AS AWDWH5"
             + "      ,SUM(OVERHOUR1) AS OVERHOUR1,SUM(OVERHOUR2) AS OVERHOUR2,SUM(OVERHOUR3) AS OVERHOUR3,SUM(OVERHOUR4) AS OVERHOUR4,SUM(OVERHOUR5) AS OVERHOUR5"
             + " FROM ("
             + "  SELECT EMPID,CASE WHEN SUM(FACWH)>84 THEN SUM(FACWH)-84 ELSE 0 END AS AWDWH1"
             + "        ,0 AS AWDWH2,0 AS AWDWH3,0 AS AWDWH4,0 AS AWDWH5,0 AS OVERHOUR1,0 AS OVERHOUR2,0 AS OVERHOUR3,0 AS OVERHOUR4,0 AS OVERHOUR5"
             + "    FROM SCHEDM "
             + "   WHERE STNID='"+ txtStnID.Text + "' AND SUBSTRING(EMPID,1,1)='B'"
             + "     AND SHEDDATE+SHEDTIME>'" + DTST +"2300' AND SHEDDATE+SHEDTIME<='"+ DT01+ "2300'"
             + "   GROUP BY EMPID"
             + " Union ALL "
             + " SELECT EMPID,0 AS AWDWH1,CASE WHEN SUM(FACWH)>84 THEN SUM(FACWH)-84 ELSE 0 END AS AWDWH2"
             + "        ,0 AS AWDWH3,0 AS AWDWH4,0 AS AWDWH5,0 AS OVERHOUR1,0 AS OVERHOUR2,0 AS OVERHOUR3,0 AS OVERHOUR4,0 AS OVERHOUR5"
             + "    FROM SCHEDM "
             + "   WHERE STNID='"+ txtStnID.Text + "' AND SUBSTRING(EMPID,1,1)='B'"
             + "     AND SHEDDATE+SHEDTIME>'" + DT01 +"2300' AND SHEDDATE+SHEDTIME<='"+ DT02+ "2300'"
             + "   GROUP BY EMPID"
             + " Union ALL "
             + " SELECT OVTM_EMPLID AS EMPID,0 AS AWDWH1,0 AS AWDWH2,0 AS AWDWH3,0 AS AWDWH4,0 AS AWDWH5"
             + "       ,SUM(OVTM_P_HOURS) AS OVERHOUR1,0 AS OVERHOUR2,0 AS OVERHOUR3,0 AS OVERHOUR4,0 AS OVERHOUR5"
             + "   FROM [MP_HR].DBO.OVERTIME"
             + "  WHERE OVTM_EMPLID IN (SELECT EMPL_ID FROM MP_HR.DBO.EMPLOYEE WHERE EMPL_DEPTID='" + txtStnID.Text + "' AND SUBSTRING(EMPL_ID,1,1)='B')"
             + "    AND OVTM_DATE>'" + DTST +"' AND OVTM_DATE<='" + DT01 + "'"
             + "  GROUP BY OVTM_EMPLID"
             + " Union All "
             + " SELECT OVTM_EMPLID AS EMPID,0 AS AWDWH1,0 AS AWDWH2,0 AS AWDWH3,0 AS AWDWH4,0 AS AWDWH5"
             + "       ,0 AS OVERHOUR1,SUM(OVTM_P_HOURS) AS OVERHOUR2,0 AS OVERHOUR3,0 AS OVERHOUR4,0 AS OVERHOUR5"
             + "   FROM [MP_HR].DBO.OVERTIME"
             + "  WHERE OVTM_EMPLID IN (SELECT EMPL_ID FROM MP_HR.DBO.EMPLOYEE WHERE EMPL_DEPTID='" + txtStnID.Text + "' AND SUBSTRING(EMPL_ID,1,1)='B')"
             + "    AND OVTM_DATE>'" + DT01 +"' AND OVTM_DATE<='" + DT02 + "'"
             + "  GROUP BY OVTM_EMPLID ";
        //
        if (DT03 !="")  
        {
          sSQL = sSQL + " Union ALL "
               + " SELECT EMPID,0 AS AWDWH1,0 AS AWDWH2,CASE WHEN SUM(FACWH)>10 THEN SUM(FACWH)-10 ELSE 0 END AS AWDWH3"
               + "        ,0 AS AWDWH4,0 AS AWDWH5,0 AS OVERHOUR1,0 AS OVERHOUR2,0 AS OVERHOUR3,0 AS OVERHOUR4,0 AS OVERHOUR5"
               + "    FROM SCHEDM "
               + "   WHERE STNID='"+ txtStnID.Text + "' AND SUBSTRING(EMPID,1,1)='B'"
               + "     AND SHEDDATE+SHEDTIME>'" + DT02 +"2300' AND SHEDDATE+SHEDTIME<='"+ DT03 + "2300'"
               + "   GROUP BY EMPID"
               + " Union All "
               + " SELECT OVTM_EMPLID AS EMPID,0 AS AWDWH1,0 AS AWDWH2,0 AS AWDWH3,0 AS AWDWH4,0 AS AWDWH5"
               + "       ,0 AS OVERHOUR1,0 AS OVERHOUR2,SUM(OVTM_P_HOURS) AS OVERHOUR3,0 AS OVERHOUR4,0 AS OVERHOUR5"
               + "   FROM [MP_HR].DBO.OVERTIME"
               + "  WHERE OVTM_EMPLID IN (SELECT EMPL_ID FROM MP_HR.DBO.EMPLOYEE WHERE EMPL_DEPTID='" + txtStnID.Text +"' AND SUBSTRING(EMPL_ID,1,1)='B')"
               + "    AND OVTM_DATE>'" + DT02 +"' AND OVTM_DATE<='" + DT03 + "'"
               + "  GROUP BY OVTM_EMPLID ";
        }    
        //
        if (DT04 !="")  
        {
          sSQL = sSQL + " Union ALL "
               + " SELECT EMPID,0 AS AWDWH1,0 AS AWDWH2,0 AS AWDWH3,CASE WHEN SUM(FACWH)>10 THEN SUM(FACWH)-10 ELSE 0 END AS AWDWH4"
               + "        ,0 AS AWDWH5,0 AS OVERHOUR1,0 AS OVERHOUR2,0 AS OVERHOUR3,0 AS OVERHOUR4,0 AS OVERHOUR5"
               + "    FROM SCHEDM "
               + "   WHERE STNID='"+ txtStnID.Text + "' AND SUBSTRING(EMPID,1,1)='B'"
               + "     AND SHEDDATE+SHEDTIME>'" + DT03 +"2300' AND SHEDDATE+SHEDTIME<='"+ DT04 + "2300'"
               + "   GROUP BY EMPID"
               + " Union All "
               + " SELECT OVTM_EMPLID AS EMPID,0 AS AWDWH1,0 AS AWDWH2,0 AS AWDWH3,0 AS AWDWH4,0 AS AWDWH5"
               + "       ,0 AS OVERHOUR1,0 AS OVERHOUR2,0 AS OVERHOUR3,SUM(OVTM_P_HOURS) AS OVERHOUR4,0 AS OVERHOUR5"
               + "   FROM [MP_HR].DBO.OVERTIME"
               + "  WHERE OVTM_EMPLID IN (SELECT EMPL_ID FROM MP_HR.DBO.EMPLOYEE WHERE EMPL_DEPTID='" + txtStnID.Text +"' AND SUBSTRING(EMPL_ID,1,1)='B')"
               + "    AND OVTM_DATE>'" + DT03 +"' AND OVTM_DATE<='" + DT04 + "'"
               + "  GROUP BY OVTM_EMPLID ";
        }    
        //
        if (DT05 !="")  
        {
          sSQL = sSQL + " Union ALL "
               + " SELECT EMPID,0 AS AWDWH1,0 AS AWDWH2,0 AS AWDWH3,0 AS AWDWH4,CASE WHEN SUM(FACWH)>10 THEN SUM(FACWH)-10 ELSE 0 END AS AWDWH5"
               + "        ,0 AS OVERHOUR1,0 AS OVERHOUR2,0 AS OVERHOUR3,0 AS OVERHOUR4,0 AS OVERHOUR5"
               + "    FROM SCHEDM "
               + "   WHERE STNID='"+ txtStnID.Text + "' AND SUBSTRING(EMPID,1,1)='B'"
               + "     AND SHEDDATE+SHEDTIME>'" + DT04 +"2300' AND SHEDDATE+SHEDTIME<='"+ DT05 + "2300'"
               + "   GROUP BY EMPID"
               + " Union All "
               + " SELECT OVTM_EMPLID AS EMPID,0 AS AWDWH1,0 AS AWDWH2,0 AS AWDWH3,0 AS AWDWH4,0 AS AWDWH5"
               + "       ,0 AS OVERHOUR1,0 AS OVERHOUR2,0 AS OVERHOUR3,0 AS OVERHOUR4,SUM(OVTM_P_HOURS) AS OVERHOUR5"
               + "   FROM [MP_HR].DBO.OVERTIME"
               + "  WHERE OVTM_EMPLID IN (SELECT EMPL_ID FROM MP_HR.DBO.EMPLOYEE WHERE EMPL_DEPTID='" + txtStnID.Text +"' AND SUBSTRING(EMPL_ID,1,1)='B')"
               + "    AND OVTM_DATE>'" + DT04 +"' AND OVTM_DATE<='" + DT05 + "'"
               + "  GROUP BY OVTM_EMPLID ";
        } 
        //
        sSQL = sSQL + ") A INNER JOIN MP_HR.DBO.EMPLOYEE E ON A.EMPID=E.EMPL_ID "
             + " INNER JOIN MECHSTNM S ON S.STNID='" + txtStnID.Text + "'"
             + " GROUP BY STNNAME,EMPID,EMPL_NAME ORDER BY EMPID";
        //--------------------------------------------* 加入業務處/責任區/加油站之過濾
        //savelog(sSQL); //azitest
        
        sFields.Add("txtStnID", txtStnID.Text);
        sFields.Add("txtStnName", txtStnName.Text);
        sFields.Add("txtDtStart", txtDtStart.Text);

        this.ViewState["QryField"] = sFields;
        this.dscList.SelectCommand = sSQL;
        this.grdList.DataSourceID = this.dscList.ID;
        this.grdList.DataBind();
        //
        if (this.grdList.Rows.Count <= 0)
        {
            System.Data.DataTable VSEmptyTable = (DataTable)ViewState["EmptyTable"];
            modDB.ShowEmptyDataGridHeader(this.grdList, 0, 10, ref VSEmptyTable);
            modUtil.showMsg(this.Page, "訊息", "查無出勤異常檢核統計資料！");
        }
        else
        {
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;      //* 用於 Excel/換頁
        }
    }

    /* protected void savelog(string VLOGSTR)
    {
        //System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\20120805\VOIR102log.txt", System.IO.FileMode.OpenOrCreate);
        System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\20140221.log", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
        System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.Unicode);
        sw.WriteLine(DateTime.Now.ToString ("hh:mm:ss")+" "+VLOGSTR); //azitemp
        sw.Flush();
        sw.Close();
        fs.Close();
    } */

    protected void txtDtStart_TextChanged(object sender, EventArgs e)
    {
        lblMsg2.Text = lblMsg2.Text + "Test";
    }
}
