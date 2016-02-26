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

partial class ATT_Q_003 : System.Web.UI.Page
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
            else //'-----------------------------------* �ˬd�O�_�� Ū��/�s�W �v��
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

                if (USER_ROL.Substring(0, 1) == "N")
                    FormsAuthentication.RedirectToLoginPage();
                else
                    //if (modUtil.IsStnRol(this.Request))  //'* �[�o���v��
                    if (modUtil.IsStnRol(this.Request) || modUnset.IsPAUnitRol(this.Request))  //'* �[�o���v��
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

                //----------------------------------------* �]�wGridView�˦�
                modDB.SetGridViewStyle(this.grdList, 20);
                this.grdList.RowStyle.Height = 20;

                modDB.SetFields("EMPID"   , "���u¾��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("EMPNAME" , "���u�m�W", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("OVTMDTST", "�_�l���", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("OVTMTMST", "�_�l�ɶ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("OVTMDTEN", "�������", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("OVTMTMEN", "�����ɶ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("VHOURS", "�ɫݥ��", this.grdList, HorizontalAlign.Right, "", false, 0, false);
                modDB.SetFields("PHOURS", "�[�Z�O��", this.grdList, HorizontalAlign.Right, "", false, 0, false);

                //'----------------------------------------* �]�w����ݩ�
                txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()");
                modUtil.SetDateObj(txtDtFrom, false, txtDtTo, false, null, null);
                modUtil.SetDateObj(txtDtTo, false, null, false, null, null);
                modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtFrom, true, false, this.txtDtTo, null, null, false);
                modUtil.SetDateImgObj(this.imgDtTo, this.txtDtTo, true, false, null, null, null, false);

                //'----------------------------------------* ���s��ܬd�ߵ��G
                this.lblMsg.Text = Request["msg"];
                if (Session["QryField"] == null)
                {
                    DataTable VSEmptyTable = (DataTable)this.ViewState["EmptyTable"];
                    modDB.ShowEmptyDataGridHeader(this.grdList, 0, 7, ref VSEmptyTable);
                }

                //'* �~�ȳB/�d����/�[�o�����e���]�w����
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
    //* ���s��ܬd�ߵ��G 
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
    //* �^�D�e���ɡA���s��ܬd�ߵ��G
    //******************************************************************************************************
    protected void ddlArea_DataBound(object sender,EventArgs e) 
    {
        if (!(Page.IsPostBack))
        {
           if (Session["QryField"] == null)
                ReflashQryData(sender);
        }
    }

    //* ����~�ȳB/�d���ϮɡA�M���[�o��
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
    //* ����~�ȳB�ɡA���M���³d���ϡA�A���[
    //******************************************************************************************************
    protected void dscArea_Selecting(object sender,System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs e) 
    {
        string USER_ROL = (string)ViewState["ROL"];

        if ((string)ViewState["ROL"] == null) USER_ROL = "      ";

        if ((! IsPostBack) && (USER_ROL.Substring(5, 1).CompareTo("1") > 0))  //* 0:�` 1:�~ 2:�� 3,4:��
              e.Cancel = true;  //* �w���v�����]�w�e���A���i���sŪ��
        else
        {
            this.ddlArea.Items.Clear();
            this.ddlArea.Items.Add("");
        }
    }

    //******************************************************************************************************
    //* ��ܥ[�o���W��(�̩ҿ�����~�ȳB/�d���ϹL�o)
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
    //* ��ܭ��X
    //******************************************************************************************************
    protected void grdList_DataBound(object sender,EventArgs e)
    {
        modDB.SetGridPageNum(this.grdList, PagerButtons.NumericFirstLast, HorizontalAlign.Left);
    }

    //******************************************************************************************************
    //* �O�s�s�����X
    //******************************************************************************************************
    protected void grdList_PageIndexChanged(object sender,EventArgs e)
    {
        this.ViewState["iCurPage"] = grdList.PageIndex;
    }

    //******************************************************************************************************
    //* ��в��ʮɡA��ܥ����ĪG
    //******************************************************************************************************
    protected void grdList_RowDataBound(object sender, System.Web.UI.WebControls.GridViewRowEventArgs e)
    {
        modDB.SetGridLightPen(e);
    }

    //******************************************************************************************************
    //* ��ܥ[�o��(�̩ҿ�����~�ȳB/�d���ϹL�o)
    //******************************************************************************************************
    protected void btnStnID_Click(object sender, EventArgs e)
    {
        //this.QryStn.Show((string)ViewState["RolType"], (string)ViewState["RolDeptID"], false);
        this.QryStn.ShowBySel(this.ddlBus.SelectedValue, this.ddlArea.SelectedValue);
    }

    //******************************************************************************************************
    //* �M��
    //******************************************************************************************************
    protected void btnClear_Click(object sender, EventArgs e)
    {
        Session["QryField"]= null;
        Session["iCurPage"] = null;
        Response.Redirect("ATT_Q_003.aspx");
    }

    //******************************************************************************************************
    //* Excel �ץX
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ003.xls", this.grdList);
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
            //----------------------------------------* �ˮ���쥿�T��
            sValidOK = modUtil.Check2DateObj(this, "txtDtFrom", "txtDtTo", ref sMsg, "�d�ߤ��");
            if ((this.txtStnID.Text == "") || (this.txtStnName.Text.Trim() == "")) 
            {
                sValidOK = false;
                sMsg = sMsg + "\n�d�L���[�o���I";
            }
            if (!( sValidOK)) 
            {
                modUtil.showMsg(this.Page, "�T��", sMsg) ;
                return;
            }
            ShowGrid(); //* ��ܬd�ߵ��G
            Session["iCurPage"] = null;
        }
        catch (InvalidCastException ex)
        {
            modUtil.showMsg(this.Page, "���~�T��(�d��)", ex.Message);
        }
    }

    //******************************************************************************************************
    //* ��ܩ���
    //******************************************************************************************************
    protected void  ShowGrid()
    {
        string sSQL ="";
        Hashtable sFields = new Hashtable();
        //
        //sSQL = "SELECT VATM_EMPLID,EMPL_NAME AS EMPNAME,VATM_VANMID,VACA_NAME AS VACANAME,VATM_DATE_ST AS VATMDTST,VATM_DATE_EN AS VATMDTEN"  //,VATM_SEQ,VATM_VANMID
        sSQL = "SELECT OVTM_EMPLID AS EMPID,EMPL_NAME AS EMPNAME,CONVERT(char(10),OVTM_DATE,111) AS OVTMDTST,OVTM_TIME_ST AS OVTMTMST"
             + "      ,CONVERT(char(10),OVTM_DATE_EN,111) AS OVTMDTEN,OVTM_TIME_EN AS OVTMTMEN,OVTM_V_HOURS AS VHOURS,OVTM_P_HOURS AS PHOURS"
             + "  FROM MP_HR.DBO.OVERTIME A"
             + " INNER JOIN MP_HR.DBO.EMPLOYEE B ON A.OVTM_EMPLID=B.EMPL_ID"
             + " WHERE OVTM_DATE>='" + txtDtFrom.Text + "'"
             + "   AND OVTM_DATE<='" + txtDtTo.Text + "'";
        if (txtEmpID.Text.Trim() != "")
            sSQL = sSQL + "   AND OVTM_EMPLID = '" + txtEmpID.Text + "'";
        //
        sSQL = sSQL + "   AND OVTM_EMPLID IN "
             + "       (SELECT EMPL_ID FROM MP_HR.DBO.EMPLOYEE WHERE EMPL_DEPTID='" + this.txtStnID.Text + "'"
             + "           AND ((EMPL_LEV_DATE IS NULL) OR (EMPL_LEV_DATE>='"+txtDtFrom.Text+"')))"
             + " ORDER BY OVTM_EMPLID,OVTM_DATE ";

        //--------------------------------------------* �[�J�~�ȳB/�d����/�[�o�����L�o
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
            modDB.ShowEmptyDataGridHeader(this.grdList, 0, 7, ref VSEmptyTable);
            modUtil.showMsg(this.Page, "�T��", "�d�L��ơI");
        }
        else
        {
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;     //* �Ω� Excel/����
            //modUtil.showMsg(this.Page, "�T��", "ok-1�I"); //azitest
        }
    }


    //******************************************************************************************************
    //* ��ܭ��u
    //******************************************************************************************************
    protected void btnEmpID_Click(object sender, EventArgs e)
    {
        if (this.txtStnID.Text != "")
        {
            this.QryEmp.Show(this.txtStnID.Text);
        }
    }

    //******************************************************************************************************
    //* ��ܭ��u�m�W
    //******************************************************************************************************

    protected void txtEmpID_TextChanged(object sender, EventArgs e)
    {
        //modCharge.GetStnName(txtStnID, this.txtStnName, this.ddlArea.SelectedValue);
        modCharge.GetEmpName(this.txtEmpID, this.txtEmpName, true);
    }

    /*
    protected void savelog(string VLOGSTR)
    {
        //System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\20120805\VOIR102log.txt", System.IO.FileMode.OpenOrCreate);
        System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\20120805\VOIR102log.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
        System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.Unicode);
        sw.WriteLine(DateTime.Now.ToString ("hh:mm:ss")+" "+VLOGSTR); //azitemp
        sw.Flush();
        sw.Close();
        fs.Close();
    }*/

}
