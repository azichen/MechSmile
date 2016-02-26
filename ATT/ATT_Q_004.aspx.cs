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

partial class ATT_Q_004 : System.Web.UI.Page
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

            //----------------------------------------* �]�wGridView�˦�
            modDB.SetGridViewStyle(this.grdList, 20);
            this.grdList.RowStyle.Height = 20;

            modDB.SetFields("EMPID"  , "���u¾��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("EMPNAME", "���u�m�W", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            //modDB.SetFields("INDATE" , "��¾���", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("HOD11"  , "�S��ɼ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("HOD12"  , "�w��S��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("HOD13"  , "�Ѿl�S��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("HOD21"  , "�ɥN��ɼ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("HOD22"  , "�w��ݥ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("HOD23"  , "�Ѿl�ݥ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);

            //'----------------------------------------* �]�w����ݩ�
            txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()");
            //'----------------------------------------* ���s��ܬd�ߵ��G
            this.lblMsg.Text = Request["msg"];
            if (Session["QryField"] == null)
            {
                DataTable VSEmptyTable = (DataTable)this.ViewState["EmptyTable"];
                modDB.ShowEmptyDataGridHeader(this.grdList, 0, 7, ref VSEmptyTable);
            }

            //'* �~�ȳB/�d����/�[�o�����e���]�w����
            //modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, (string)this.ViewState["ROL"], (string)this.ViewState["RolDeptID"]);
            this.txtEmpName.Attributes["readonly"] = "readonly";
            this.txtEmpName.Attributes["Class"] = "readonly";
            modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, USER_ROL, USER_ROLDEPTID);
            //if (modUtil.IsStnRol(this.Request) || modUnset.IsPAUnitRol(this.Request))  //'* �[�o���v��
            //if (USER_ROLTYPE == "STN")                                                                                
            //    modUtil.SetObjReadOnly(form1, "txtStnID", false);
            //}
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
            txtDtStart.Text = modUnset.GET_LOCKYM("2", "").Substring(0, 4);
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
        Response.Redirect("ATT_Q_004.aspx");
    }

    //******************************************************************************************************
    //* Excel �ץX
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ004.xls", this.grdList);
    }

    protected void btnQry_Click(object sender,EventArgs e)
    {
        string sMsg = "";
        Boolean sValidOK = true;
        //
        try
        {
            this.lblMsg.Text = "";
            sValidOK = true;
            this.txtStnID.Text = this.txtStnID.Text.Trim();
            //----------------------------------------* �ˮ���쥿�T��
            if ((this.txtStnID.Text == "") || (this.txtStnName.Text.Trim() == "")) 
            {
                sValidOK = false;
                sMsg = sMsg + "\n�d�L���[�o���I";
            }
            //if (sValidOK)
            //{
            //    modUtil.CheckNotEmpty(this, "txtDtStart", sMsg, "�~��", false);
            //}
            //
            //modDB.InsertSignRecord("AziTest", "input accym=" + txtDtStart.Text, User.Identity.Name); //AZITEST
            //modDB.InsertSignRecord("AziTest", "input accym-1=" + Convert.ToString(ACCYEAR - 1) + " accym+1=" + Convert.ToString(ACCYEAR + 1), User.Identity.Name); //AZITEST
            if (!(txtDtStart.Text.Length == 4))
            {
                sValidOK = false;
                sMsg = sMsg + "\n�~�׻ݿ�J�褸�~4�X!!";
            }
            else
            {
                //SqlCommand cmd;
                //SqlDataReader dr;
                //SqlConnection cn = new SqlConnection();
                //cn.ConnectionString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["MECHPAConnectionString"].ToString();
                //cn.Open();
                //string sSQL = "";
                //sSQL = "SELECT CLCT_YEAR,CLCT_MONTH FROM MP_HR.DBO.CLOSE_CONTROL "
                //     + " WHERE CLCT_COMPID='01' AND CLCT_KIND='0'";
                //cmd = new SqlCommand(sSQL, cn);
                //dr = cmd.ExecuteReader();
                //dr.Read();
                int ACCYEAR = Convert.ToInt16(modUnset.GET_LOCKYM("2", "").Substring(0,4));
                //
                //modDB.InsertSignRecord("AziTest", "AAinput accym-1=" + Convert.ToString(ACCYEAR - 1) + " accym+1=" + Convert.ToString(ACCYEAR + 1), User.Identity.Name); //AZITEST
                //
                if (!((Convert.ToInt16(txtDtStart.Text)>=(ACCYEAR-1)) && (Convert.ToInt16(txtDtStart.Text)<=(ACCYEAR+1))))
                {
                    sValidOK = false;
                    sMsg = sMsg + "\n �~�׭����J:" + Convert.ToString(ACCYEAR - 1) + " ~ " + Convert.ToString(ACCYEAR + 1);
                }
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
        string sSQL = "";
        Hashtable sFields = new Hashtable();
        //

        string VACCYM = modUnset.GET_LOCKYM("2", "");
        //string VACCYM = dr.GetInt16(0).ToString() + String.Format("{0:00}", dr.GetInt16(1));
        string VACCYM_Pre = "";
        if (VACCYM.Substring(4, 2) == "01")
            VACCYM_Pre = Convert.ToString(Convert.ToInt16(VACCYM.Substring(0, 4)) - 1) + "12";
        else
            VACCYM_Pre = VACCYM.Substring(0, 4) + string.Format("{0:00}", Convert.ToInt16(VACCYM.Substring(4, 2)) - 1);
        //
        string VYEAR = txtDtStart.Text.Substring(0, 4) ;
        //
        
        string FILTERSTR = " IN (SELECT DISTINCT EMCG_EMPLID FROM [MP_HR].DBO.EMPLOYEE_CHANGE "
                      + "      WHERE EMCG_DEPTID='" + this.txtStnID.Text + "' AND SUBSTRING(EMCG_EMPLID,1,1)<>'B' "
                      + "        AND EMCG_DATE<='" + VYEAR + "1220' AND EMCG_END_DATE>='" + VACCYM_Pre + "21'" //+ "        AND EMCG_DATE<='" + VACCYM.Substring(0, 4) +"1220' AND EMCG_END_DATE>='"+VACCYM_Pre+"21'"
                      + "        AND EMCG_TYPE NOT IN ('5','6')) ";

        sSQL = "SELECT A.EMPID,EMPl_NAME AS EMPNAME,SUM(HOD11) AS HOD11,SUM(HOD12) AS HOD12,(SUM(HOD11)-SUM(HOD12)) AS HOD13"
             + "            ,SUM(HOD21) AS HOD21,SUM(HOD22) AS HOD22,(SUM(HOD21)-SUM(HOD22)) AS HOD23"
             + "  FROM ( "
             + "  SELECT VABL_EMPLID AS EMPID,VABL_ADD_HOURS AS HOD11,0 AS HOD12,0 AS HOD21,0 AS HOD22 "
             + "    FROM [MP_HR].DBO.VACA_BALANCE "
             + "   WHERE VABL_YEAR='" + VYEAR + "' AND VABL_EMPLID " + FILTERSTR;
        if (txtEmpID.Text.Trim() != "")
            sSQL = sSQL + " AND VABL_EMPLID='" + txtEmpID.Text.Trim() + "'";
        sSQL = sSQL + " Union All "
             + " SELECT VATM_EMPLID AS EMPID,0 AS HOD11,SUM(VATM_HOURS) AS HOD12,0 AS HOD21,0 AS HOD22 "
             + "   FROM [MP_HR].DBO.VACATM "
             + "  WHERE VATM_VANMID='000'"
             + "   AND VATM_DATE_ST>='" + Convert.ToString(Convert.ToInt16(VYEAR) - 1) + "1221'"
             + "   AND VATM_DATE_ST<='" + VYEAR + "1220'"
             +"   AND VATM_EMPLID" + FILTERSTR;
        if (txtEmpID.Text.Trim() != "")
            sSQL = sSQL + " AND VATM_EMPLID='" + txtEmpID.Text.Trim() + "'";
        sSQL = sSQL + "  GROUP BY VATM_EMPLID "
             + " Union All "
             + " SELECT OVTM_EMPLID AS EMPID,0 AS HOD11,0 AS HOD12,SUM(OVTM_V_HOURS) AS HOD21,0 AS HOD22 "
             + "   FROM [MP_HR].DBO.OVERTIME "
             + "  WHERE OVTM_DATE>='" + Convert.ToString(Convert.ToInt16(VYEAR) - 1) + "1221'"
             + "    AND OVTM_DATE<='" + VYEAR + "1220'"
             + "    AND OVTM_EMPLID" + FILTERSTR;
        if (txtEmpID.Text.Trim() != "")
            sSQL = sSQL + " AND OVTM_EMPLID='" + txtEmpID.Text.Trim() + "'";
        sSQL = sSQL + "  GROUP BY OVTM_EMPLID "
            + " Union All "
            + " SELECT VATM_EMPLID AS EMPID,0 AS HOD11,0 AS HOD12,0 AS HOD21,SUM(VATM_HOURS) AS HOD22 "
            + "   FROM [MP_HR].DBO.VACATM "
            + "  WHERE VATM_VANMID='013'"
            + "   AND VATM_DATE_ST>='" + Convert.ToString(Convert.ToInt16(VYEAR) - 1) + "1221'"
            + "   AND VATM_DATE_ST<='" + VYEAR + "1220'"
            +"    AND VATM_EMPLID" + FILTERSTR;
        if (txtEmpID.Text.Trim() != "")
            sSQL = sSQL + " AND VATM_EMPLID='" + txtEmpID.Text.Trim() + "'";
        sSQL = sSQL + "  GROUP BY VATM_EMPLID ) A "
            + " INNER JOIN [MP_HR].DBO.EMPLOYEE B ON A.EMPID=B.EMPL_ID"
            + " GROUP BY A.EMPID,EMPL_NAME "
            + " ORDER BY A.EMPID,EMPL_NAME ";

        //savelog(sSQL); //azitest
        //--------------------------------------------* �[�J�~�ȳB/�d����/�[�o�����L�o
        sFields.Add("txtStnID", txtStnID.Text);
        sFields.Add("txtStnName", txtStnName.Text);
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

    
    //protected void savelog(string VLOGSTR)
    //{
    //    //System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\20120805\VOIR102log.txt", System.IO.FileMode.OpenOrCreate);
    //    System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\sqlog.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
    //    System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.Unicode);
    //    sw.WriteLine(DateTime.Now.ToString ("hh:mm:ss")+" "+VLOGSTR); //azitemp
    //    sw.Flush();
    //    sw.Close();
    //    fs.Close();
    //}

}
