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

partial class ATT_Q_019 : System.Web.UI.Page
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
            modDB.SetFields("STNID", "訪站代號", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("STNNAME", "訪站站名", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDA", "平日訪次/訪站日期", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDB", "晚班訪次/起始時間", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDC", "夜間訪次/迄止時間", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDD", "假日訪次/時間(分)", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDE", "時間不符/歸類", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            
            //'--------------------------------------------------* 加入連結欄位
            this.grdList.RowStyle.Height = 20;

            //'----------------------------------------* 設定欄位屬性
            
            modUtil.SetDateObj(txtDtFrom, false, txtDtTo, false, null, null);
            modUtil.SetDateObj(txtDtTo, false, null, false, null, null);
            modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtFrom, true, false, this.txtDtTo, null, null, false);
            modUtil.SetDateImgObj(this.imgDtTo, this.txtDtTo, true, false, null, null, null, false);

            //'----------------------------------------* 重新顯示查詢結果
            this.lblMsg.Text = Request["msg"];

            //'* 業務處/責任區/加油站之畫面設定控制
            this.txtEmpName.Attributes["readonly"] = "readonly";
            this.txtEmpName.Attributes["Class"] = "readonly";

            if (Session["QryField"] == null)
            {
                DataTable VSEmptyTable = (DataTable)this.ViewState["EmptyTable"];
                modDB.ShowEmptyDataGridHeader(this.grdList, 0, 8, ref VSEmptyTable);
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
        }
        //
        Session["QryField"] = null;
        Session["iCurPage"] = null;
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
    //* 清除
    //******************************************************************************************************
    protected void btnClear_Click(object sender, EventArgs e)
    {
        Session["QryField"]= null;
        Session["iCurPage"] = null;
        Response.Redirect("ATT_Q_019.aspx");
    }

    //******************************************************************************************************
    //* Excel 匯出
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ019.xls", this.grdList);
    }

    protected void btnQry_Click(object sender,EventArgs e)
    {
        string sMsg = "";
        Boolean sValidOK = true;
        //要限制查單站+31天的資料
        try
        {
            sValidOK = true;
            sValidOK = modUtil.Check2DateObj(this, "txtDtFrom", "txtDtTo", ref sMsg, "查詢日期");
            //----------------------------------------* 檢核欄位正確性

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
        string DTSTSTR, DTEDSTR = "";
        string sSQL = "";

        //modDB.InsertSignRecord("AziTest", "OK-1" , User.Identity.Name); //azitest

        DTSTSTR = txtDtFrom.Text.Substring(0, 4) + txtDtFrom.Text.Substring(5, 2) + txtDtFrom.Text.Substring(8, 2);
        DTEDSTR = txtDtTo.Text.Substring(0, 4) + txtDtTo.Text.Substring(5, 2) + txtDtTo.Text.Substring(8, 2);
      
        sSQL = "EXEC ATTQ019R '" + DTSTSTR + "','" + DTEDSTR+"'";
        if (txtEmpID.Text != "" )
        {
            sSQL =sSQL + ",'"+txtEmpID.Text+"'";
        }
        else
        {
            sSQL = sSQL + ",''";
        }

        if (rdoType.SelectedValue == "1")
        {
            sSQL = sSQL + ",'1'"; //統計表
        }
        else
        {
            sSQL = sSQL + ",'2'"; //明細表
        }

        //modDB.InsertSignRecord("AziTest", sSQL, User.Identity.Name); //azitest

        //modUnset.PA_ExcuteCommand(sSQL);

        //--------------------------------------------* 加入業務處/責任區/加油站之過濾
        //savelog(sSQL);
        this.dscList.SelectCommand = sSQL;
        this.grdList.DataSourceID = this.dscList.ID;
        this.grdList.DataBind();
        //modDB.InsertSignRecord("AziTest", "ATTQ005_ShowGrid_OK_5", User.Identity.Name); //azitest

        //modUtil.showMsg(this.Page, sSQL); //azitest
        if (this.grdList.Rows.Count <= 0)
        {
            System.Data.DataTable VSEmptyTable = (DataTable)ViewState["EmptyTable"];
            modDB.ShowEmptyDataGridHeader(this.grdList, 0, 8, ref VSEmptyTable);
            modUtil.showMsg(this.Page, "訊息", "查無資料！");
        }
        else
        {
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;      //* 用於 Excel/換頁
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
