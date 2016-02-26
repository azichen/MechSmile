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

partial class ATT_Q_018 : System.Web.UI.Page
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
                //
                if (USER_ROL.Substring(0, 1) == "N")
                    FormsAuthentication.RedirectToLoginPage();
                else
                    if (modUtil.IsStnRol(this.Request) || modUnset.IsPAUnitRol(this.Request))  //'* �[�o���v��
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

                //----------------------------------------* �]�wGridView�˦�
                modDB.SetGridViewStyle(this.grdList, 20);
                this.grdList.RowStyle.Height = 20;

                modDB.SetFields("STNNAME", "���O", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("EMPID", "���u�s��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("EMPNAME", "���u�m�W", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("SHSTDATE", "�X�Ԯɼ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("SHSTTIME", "�_�l�ɶ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("SHEDTIME", "�W��ɶ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("FACWH", "��ڤu��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("DAYOVER", "���W��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
                modDB.SetFields("TWKOVER", "���P�W��", this.grdList, HorizontalAlign.Center, "", false, 0, false);

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
                    modDB.ShowEmptyDataGridHeader(this.grdList, 0,8, ref VSEmptyTable);
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
        Response.Redirect("ATT_Q_018.aspx");
    }

    //******************************************************************************************************
    //* Excel �ץX
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ017.xls", this.grdList);
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
        //
        Hashtable sFields = new Hashtable();
        if ((txtDtFrom.Text!="") && (txtDtTo.Text!="")) 
        {
            
            string VDATEST, VDATEED = "";

            //sSQL = modUnset.GETFWDS(txtDtFrom.Text);
            VDATEST = txtDtFrom.Text.Substring(0, 4) + txtDtFrom.Text.Substring(5, 2) + txtDtFrom.Text.Substring(8, 2);
            VDATEED = txtDtTo.Text.Substring(0, 4) + txtDtTo.Text.Substring(5, 2) + txtDtTo.Text.Substring(8, 2) ;
            //
            //LBLMSG.TEXT = "�d�߶g������϶�: " + VDATEST + " ~ " + VDATEED;
            //LABEL1.TEXT = "�d�߶g������϶�: " + VDATEST + " ~ " + VDATEED;
            //modDB.InsertLogRecord("AZITEST", "ATTQ018_ddlBus =" + ddlBus.SelectedValue + " ddlArea=" + ddlArea.SelectedValue, User.Identity.Name); //AZITEST 

            //savelog("ATTQ018_ddlBus =" + ddlBus.SelectedValue + " ddlArea=" + ddlArea.SelectedValue);

            sSQL = "SELECT S.STNID,STNNAME,S.EMPID,EMPNAME,SHSTDATE,SHSTTIME,SHEDTIME,FACWH,DAYOVER,TWKOVER FROM "
                 + "( SELECT STNID,EMPID "
                 + "        ,SUBSTRING(SHSTDATE,1,4)+'/'+SUBSTRING(SHSTDATE,5,2)+'/'+SUBSTRING(SHSTDATE,7,2) AS SHSTDATE"
                 + "        ,SUBSTRING(SHSTTIME,1,2)+':'+SUBSTRING(SHSTTIME,3,2) AS SHSTTIME"
                 + "        ,SUBSTRING(SHEDTIME,1,2)+':'+SUBSTRING(SHEDTIME,3,2) AS SHEDTIME,FACWH,DAYOVER,0 AS TWKOVER"
                 + "    FROM SCHEDM WHERE SHSTDATE>='" + VDATEST + "' AND SHSTDATE<='" + VDATEED + "' AND DAYOVER>0 ";
            
            if (txtStnID.Text.Trim() != "")
            {
                sSQL = sSQL + " AND STNID='" + txtStnID.Text.Trim() + "'";
                //savelog("txtStnID.Text =" + txtStnID.Text);
            }
            else if (ddlArea.SelectedValue != "")
            {
                sSQL = sSQL + " AND STNID IN (SELECT STNID FROM [10.1.1.100].SMILE_HQ.DBO.MECHSTNM_VIEW WHERE ARE_ID='" + ddlArea.SelectedValue + "')";
                //savelog(" ddlArea=" + ddlArea.SelectedValue);
            }
            else if (ddlBus.SelectedValue != "")
            {
                sSQL = sSQL + " AND STNID IN (SELECT STNID FROM [10.1.1.100].SMILE_HQ.DBO.MECHSTNM_VIEW WHERE BUS_ID='" + ddlBus.SelectedValue + "')";
                //savelog("ATTQ018_ddlBus =" + ddlBus.SelectedValue);
            };
            
            if (txtEmpID.Text.Trim() != "")
            {
                sSQL = sSQL + " AND EMPID='" + txtEmpID.Text.Trim() + "'";
            }

            SqlCommand cmd;
            SqlDataReader dr;
            SqlConnection cn = new SqlConnection();
            string sSQL2 = "";
            cn.ConnectionString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["MECHPAConnectionString"].ToString();
            cn.Open();
            sSQL2 = "Select FWDS_ID,FWDS_DATEST,FWDS_DATEEN,FWDS_STNWH from FWDUTYSH "
                  + " WHERE FWDS_DATEST BETWEEN '" + txtDtFrom.Text + "' AND '" + txtDtTo.Text + "'";
            cmd = new SqlCommand(sSQL2, cn);
            dr = cmd.ExecuteReader();
            string PDATEST, PDATEED = "";
            double PSTNWH=0;
            if (dr.HasRows)
            {
                while (dr.Read())
                {
                    //sSQL = modUnset.GETFWDS(txtDtFrom.Text);
                    PDATEST = dr.GetString(1).Substring(0, 4) + dr.GetString(1).Substring(5, 2) + dr.GetString(1).Substring(8, 2);
                    PDATEED = dr.GetString(2).Substring(0, 4) + dr.GetString(2).Substring(5, 2) + dr.GetString(2).Substring(8, 2);
                    PSTNWH = dr.GetDouble(3);

                    sSQL = sSQL + " UNION ALL "
                         + " SELECT STNID,EMPID,'" + dr.GetString(1) + "' AS SHSTDATE,'00:00' AS SHSTTIME,'00:00' AS SHEDTIME "
                         + "       ,0 AS FACWH,0 AS DAYOVER,(SUM(FACWH)-" + PSTNWH.ToString() + ") AS TWKOVER FROM SCHEDM "
                         + " WHERE SHSTDATE BETWEEN '" + PDATEST + "' AND '" + PDATEED + "' ";
                    if (txtStnID.Text.Trim() != "")
                    {
                        sSQL = sSQL + " AND STNID='" + txtStnID.Text.Trim() + "'";
                        //savelog("txtStnID.Text =" + txtStnID.Text);
                    }
                    else if (ddlArea.SelectedValue != "")
                    {
                        sSQL = sSQL + " AND STNID IN (SELECT STNID FROM [10.1.1.100].SMILE_HQ.DBO.MECHSTNM_VIEW WHERE ARE_ID='" + ddlArea.SelectedValue + "')";
                        //savelog(" ddlArea=" + ddlArea.SelectedValue);
                    }
                    else if (ddlBus.SelectedValue != "")
                    {
                        sSQL = sSQL + " AND STNID IN (SELECT STNID FROM [10.1.1.100].SMILE_HQ.DBO.MECHSTNM_VIEW WHERE BUS_ID='" + ddlBus.SelectedValue + "')";
                        //savelog("ATTQ018_ddlBus =" + ddlBus.SelectedValue);
                    };

                    if (txtEmpID.Text.Trim() != "")
                    {
                        sSQL = sSQL + " AND EMPID='" + txtEmpID.Text.Trim() + "'";
                    }
                    sSQL = sSQL + "  GROUP BY STNID,EMPID";
                }
            }
            dr.Close();

            sSQL = sSQL + ") S "
                 + " INNER JOIN MECHSTNM M ON S.STNID=M.STNID "
                 + " INNER JOIN MECHEMPM E ON S.EMPID=E.EMPID "
                 + " ORDER BY S.STNID,S.EMPID,SHSTDATE,SHSTTIME ";
            //AZITEMP 20151110
            //sSQL = sSQL + " '','','" + TXTSTNID.TEXT + "'";
            //savelog(sSQL);
            //MODDB.INSERTSIGNRECORD("AZITEST", "ATTQ018_SSQL =" + SSQL, USER.IDENTITY.NAME); //AZITEST
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
                modDB.ShowEmptyDataGridHeader(this.grdList, 0, 15, ref VSEmptyTable);
                modUtil.showMsg(this.Page, "�T��", "�d�L��ơI");
            }
            else
            {
                this.btnExcel.Enabled = true;
                ViewState["Sql"] = sSQL;      //* �Ω� Excel/����
                //modUtil.showMsg(this.Page, "�T��", "ok-1�I"); //azitest
            }
        }
        else
        {
            //MODUTIL.SHOWMSG(THIS.PAGE, "�T��", "�ݿ�J�_�W����I");
            modUtil.showMsg(this.Page, "���~�T��(�d��)", "�ݿ�J�_�W����I");
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

    protected void btnEmp_Click(object sender, EventArgs e)
    {
        this.QryEmp.Show(this.txtStnID.Text);
    }


    //protected void savelog(string VLOGSTR)
    //{
    //    //System.IO.FileStream fs = new System.IO.FileStream(@"C:\temp\20120805\VOIR102log.txt", System.IO.FileMode.OpenOrCreate);
    //    System.IO.FileStream fs = new System.IO.FileStream(@"C:\mech\201310log.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
    //    System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.Unicode);
    //    sw.WriteLine(DateTime.Now.ToString("hh:mm:ss") + " " + VLOGSTR); //azitemp
    //    sw.Flush();
    //    sw.Close();
    //    fs.Close();
    //}

}
