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

partial class ATT_Q_011 : System.Web.UI.Page
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
            modDB.SetFields("EMPID"  , "職號", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("EMPNAME", "姓名", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVMFS"  , "月超時(標準+46H)", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVTW"   , "單日超時數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("COLLNW" , "建教生_大夜出勤數", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("COLLTT" , "建教生_2230後出勤", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("COLLTT2", "建教生_2230後下班", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("DWEEKW" , "未兩周休兩天", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVWA"   , "前兩週>84H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVWB"   , "後兩週>84H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVFW"   , "四週>168H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            //'----------------------------------------* 設定欄位屬性
            txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()");

            modUtil.SetDateObj(txtDtStart, false, null, true, null, null);
            //modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtStart, true, false, null, null, null, false);
            //modUtil.SetDateObj(txtDtFrom, false, txtDtTo, false, null, null);
            //modUtil.SetDateObj(txtDtTo, false, null, false, null, null);
            //modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtFrom, true, false, this.txtDtTo, null, null, false);
            //modUtil.SetDateImgObj(this.imgDtTo, this.txtDtTo, true, false, null, null, null, false);

            //modUtil.SetDateObj(Me.txtDtStart, False, Nothing, True);

            //'----------------------------------------* 重新顯示查詢結果
            this.lblMsg.Text = Request["msg"];

            if (Session["QryField"] == null)
            {
                DataTable VSEmptyTable = (DataTable)this.ViewState["EmptyTable"];
                modDB.ShowEmptyDataGridHeader(this.grdList, 0, 11, ref VSEmptyTable);
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
        Response.Redirect("ATT_Q_011.aspx");
    }

    //******************************************************************************************************
    //* Excel 匯出
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ011.xls", this.grdList);
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
        //
        string VACCYM = "";
        //
        Hashtable sFields = new Hashtable();
        string sSQL, sSQL2 = "";
        //
        VACCYM = txtDtStart.Text.Substring(0, 4) + txtDtStart.Text.Substring(5, 2);
        //
        //modDB.InsertSignRecord("AziTest", "VACCYM=" + VACCYM+" STNID=" + txtStnID.Text , User.Identity.Name); //AZITEST
        //modUtil.showMsg(this.Page, "訊息", "資料完成！");

        sSQL = "SELECT CONVERT(VARCHAR(12),DEPT_NAME) AS STNNAME,A.EMPID,CONVERT(VARCHAR(12),EMPL_NAME) AS EMPNAME"
             + " ,OVTW,ISCOLL,NWQTY,TTQTY,TTQTY2,DWEEKW,OVMFS"
             + " ,CASE WHEN FWHR>168 THEN 1 ELSE 0 END AS OVFW"
             + " ,CASE WHEN WAFACWH>84 THEN 1 ELSE 0 END AS OVWA"
             + " ,CASE WHEN WBFACWH>84 THEN 1 ELSE 0 END AS OVWB"
             + " ,CASE WHEN ISCOLL=1 THEN TTQTY ELSE 0 END AS COLLTT"
             + " ,CASE WHEN ISCOLL=1 THEN TTQTY2 ELSE 0 END AS COLLTT2"
             + " ,CASE WHEN ISCOLL=1 THEN NWQTY ELSE 0 END AS COLLNW"
             + "  FROM MABNCHK A"
             + " INNER JOIN MP_HR.DBO.EMPLOYEE B ON A.EMPID=B.EMPL_ID"
             + " INNER JOIN MP_HR.DBO.DEPARTMENT C ON A.STNID=C.DEPT_ID"
             + " WHERE STNID= '" + txtStnID.Text + "' AND ACCYM='" + VACCYM + "' ORDER BY STNID,EMPID";
       
        //" + txtStnID.Text + "','" + DTSTSTR + "','" + DTEDSTR + "'";
        //--------------------------------------------* 加入業務處/責任區/加油站之過濾
        sFields.Add("txtStnID", txtStnID.Text);
        sFields.Add("txtStnName", txtStnName.Text);
        sFields.Add("txtDtStart", txtDtStart.Text);

        this.ViewState["QryField"] = sFields;
        this.dscList.SelectCommand = sSQL;
        this.grdList.DataSourceID = this.dscList.ID;
        this.grdList.DataBind();
        //modDB.InsertSignRecord("AziTest", "ok-6", User.Identity.Name);

        if (this.grdList.Rows.Count <= 0)
        {
            System.Data.DataTable VSEmptyTable = (DataTable)ViewState["EmptyTable"];
            modDB.ShowEmptyDataGridHeader(this.grdList, 0, 11, ref VSEmptyTable);
            modUtil.showMsg(this.Page, "訊息", "查無出勤異常檢核統計資料！");
        }
        else
        {
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;      //* 用於 Excel/換頁
        }
    }

}
