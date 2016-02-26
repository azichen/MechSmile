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

partial class ATT_Q_006 : System.Web.UI.Page
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
            //SHOWEMPID,FIELD02,FIELD03,FIELD04,FIELD05,FIELD06
            //modDB.SetFields("STNNAME" , "���W��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("SHOWEMPID", "���u¾��/�m�W", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            //modDB.SetFields("EMPNAME" , "���u�m�W", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELD02", "�ƯZ���", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELD03", "�Ʃw�W�Z", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELD04", "�Ʃw�U�Z", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELD05", "��ڤW�Z", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELD06", "��ڤU�Z", this.grdList, HorizontalAlign.Center, "", false, 0, false);
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
                modDB.ShowEmptyDataGridHeader(this.grdList, 0, 5, ref VSEmptyTable);
            }
            //'* �~�ȳB/�d����/�[�o�����e���]�w����
            //modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, (string)this.ViewState["ROL"], (string)this.ViewState["RolDeptID"]);
            this.txtEmpName.Attributes["readonly"] = "readonly";
            this.txtEmpName.Attributes["Class"] = "readonly";
            modUtil.SetRolScreen(this.txtStnID, this.txtStnName, this.btnStnID, this.ddlBus, this.ddlArea, USER_ROL, USER_ROLDEPTID);
            //if (USER_ROLTYPE == "STN")
            //    modUtil.SetObjReadOnly(form1,"txtStnID", false);
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
    //* ��ܭ��u�m�W
    //******************************************************************************************************

    protected void txtEmpID_TextChanged(object sender, EventArgs e)
    {
         //modCharge.GetStnName(txtStnID, this.txtStnName, this.ddlArea.SelectedValue);
         modCharge.GetEmpName(txtEmpID, this.txtEmpName,true);
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
    //* �M��
    //******************************************************************************************************
    protected void btnClear_Click(object sender, EventArgs e)
    {
        Session["QryField"]= null;
        Session["iCurPage"] = null;
        Response.Redirect("ATT_Q_006.aspx");
    }

    //******************************************************************************************************
    //* Excel �ץX
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ006.xls", this.grdList);
    }

    protected void btnQry_Click(object sender,EventArgs e)
    {
        string sMsg = "";
        Boolean sValidOK = true;
        //�n����d�毸+31�Ѫ����
        try
        {
            sValidOK = true;
            this.txtStnID.Text = this.txtStnID.Text.Trim();
            sValidOK = modUtil.Check2DateObj(this, "txtDtFrom", "txtDtTo", ref sMsg, "�d�ߤ��");
            //----------------------------------------* �ˮ���쥿�T��
            if ((this.txtStnID.Text == "") || (this.txtStnName.Text.Trim() == "")) 
            {
                sValidOK = false;
                sMsg = sMsg + "\n�d�L���[�o���I";
            }
            if (sValidOK)
            {
                if ((txtDtFrom.Text.Trim() == "") || (txtDtTo.Text.Trim() == ""))
                {
                    sValidOK = false;
                    sMsg = sMsg + "\n �ݿ�J����d��!";
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
                    sMsg = sMsg + "\n ����d��31�Ѥ�����ơA�Э��s��J����d��!";
                }
            }

            if (!( sValidOK)) 
            {
                modUtil.showMsg(this.Page, "�T��", sMsg) ;
                return;
            }
            ShowGrid(); //* ��ܬd�ߵ��G
            Session["iCurPage"] = null;
            this.lblMsg.Text = "";
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
        string DTSTSTR, DTEDSTR = "";
        //string LCKDT = modUnset.GET_LOCKYM("1", txtStnID.Text) + "20";
        string LCKDT = modUnset.GET_LOCKDT(txtStnID.Text);
        Boolean SCHECKFLAG = true;
        Hashtable sFields = new Hashtable();
        string sSQL, sSQL2 = "";
        string Q005DTST;
        //�����X�Ը�ƪ��B�z(�PATT_Q_005R�P)
        DTSTSTR = txtDtFrom.Text.Substring(0, 4) + txtDtFrom.Text.Substring(5, 2) + txtDtFrom.Text.Substring(8, 2);
        DTEDSTR = txtDtTo.Text.Substring(0, 4) + txtDtTo.Text.Substring(5, 2) + txtDtTo.Text.Substring(8, 2);
        //if (string.Compare(DTSTSTR, LCKDT) <= 0)
        //{
        //    SCHECKFLAG = false;
        //}
        //if (string.Compare(DTEDSTR, LCKDT) <= 0)
        //    SCHECKFLAG = false;
        //else if ((string.Compare(DTSTSTR, LCKDT) <= 0) && (string.Compare(DTEDSTR, LCKDT) > 0))
        //    DTSTSTR = LCKDT;
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
        //
        //modDB.InsertSignRecord("AziTest", "LCKDT=" + LCKDT + " DTSTSTR=" + DTSTSTR + " DTEDSTR=" + DTEDSTR, User.Identity.Name); //azitest
        //�L�F�L�b����A���i�A���Ƨ@���ʡC
        if (SCHECKFLAG)
        {
            sSQL = "EXEC ATTQ005R '" + txtStnID.Text + "','" + Q005DTST + "','" + DTEDSTR + "','" + txtEmpID.Text + "','" + txtEmpID.Text + "'";
            modUnset.PA_ExcuteCommand(sSQL);
        }
        //
        //dr.Close();
        //cn.Close();
        //cmd.Dispose(); //�������귽
        //cn.Dispose();
        //cn = null; //��������
        string PREDAYSTR, AFTDAYSTR = ""; //�_�l�e�@�ѡB�̫᪺�j�ѡA�Ѹ�Ʈw�d�ߥ�
        PREDAYSTR = String.Format("{0:yyyyMMdd}", DateTime.Parse(txtDtFrom.Text).AddDays(-1));
        AFTDAYSTR = String.Format("{0:yyyyMMdd}", DateTime.Parse(txtDtTo.Text).AddDays(+1));

        //modUtil.showMsg(this.Page, "�T��", "��Ƨ����I");

        sSQL = "SELECT SHOWEMPID,FIELD02,FIELD03,FIELD04,FIELD05,FIELD06 FROM ("
             + "  SELECT EMPID,EMPID+' '+EMPL_NAME AS SHOWEMPID,SHSTDATE+SHSTTIME AS DATADT,'1' AS KINDID "
             + "        ,SUBSTRING(SHSTDATE,1,4)+'/'+SUBSTRING(SHSTDATE,5,2)+'/'+SUBSTRING(SHSTDATE,7,2) AS FIELD02"
             + "        ,SUBSTRING(SHSTTIME,1,2)+':'+SUBSTRING(SHSTTIME,3,2) AS FIELD03"
             + "        ,SUBSTRING(SHEDTIME,1,2)+':'+SUBSTRING(SHEDTIME,3,2) AS FIELD04"
             + "        ,SUBSTRING(WKSTTIME,1,2)+':'+SUBSTRING(WKSTTIME,3,2) AS FIELD05"
             + "        ,SUBSTRING(WKEDTIME,1,2)+':'+SUBSTRING(WKEDTIME,3,2) AS FIELD06 FROM SCHEDM S"
             + " INNER JOIN MP_HR.DBO.EMPLOYEE E ON S.EMPID=E.EMPL_ID "
             + " WHERE STNID='" + txtStnID.Text + "'"
             + "   AND SHSTDATE >= '" + DTSTSTR + "'"
             + "   AND SHSTDATE <= '" + DTEDSTR + "'";
             //+ "   AND (SHEDDATE+SHEDTIME) >  '" + PREDAYSTR + "2300'"
             //+ "   AND (SHEDDATE+SHEDTIME) <= '" + DTEDSTR + "2300'";
        if (txtEmpID.Text.Trim() != "")
            sSQL = sSQL + "   AND S.EMPID = '" + txtEmpID.Text + "'";
        //
        //sSQL = sSQL + "   AND FACFLAG<>'Y' AND ((WKEDTIME<SHEDTIME) OR (WKSTTIME>SHSTTIME) OR (WKSTTIME IS NULL) OR (WKEDTIME IS NULL)) "
        sSQL = sSQL + "   AND FACFLAG<>'Y' AND ((FACWH<>WORKHOUR) OR (WORKHOUR<=0)) "
             + " UNION ALL "
             + " SELECT EMPID AS EMPID,'' AS SHOWEMPID,SHSTDATE+SHSTTIME AS DATADT,'2' AS KINDID,'' AS FIELD02 "
             + "       ,SUBSTRING([1],1,2)+':'+SUBSTRING([1],3,2) AS FIELD03,SUBSTRING([2],1,2)+':'+SUBSTRING([2],3,2) AS FIELD04"
             + "       ,SUBSTRING([3],1,2)+':'+SUBSTRING([3],3,2) AS FIELD05,SUBSTRING([4],1,2)+':'+SUBSTRING([4],3,2) AS FIELD06"
             + "   FROM ("
             + "       SELECT F.EMPID,S.SHSTDATE,S.SHSTTIME,FINTIME,ROW_NUMBER() OVER ( Partition by F.EMPID,S.SHSTDATE,S.SHSTTIME Order by FINDATE,FINTIME ) AS EMPROW"
             + "         FROM SCHEDM S, FINGER F"
             + "        WHERE S.STNID = '" + txtStnID.Text + "'" // --AND S.EMPID>=@QEMPID1 AND S.EMPID<=@QEMPID2 AND S.FACFLAG<>'Y'
             + "   AND SHSTDATE >= '" + DTSTSTR + "'"
             + "   AND SHSTDATE <= '" + DTEDSTR + "'"
             //+ "          AND (S.SHEDDATE+S.SHEDTIME) >  '" + PREDAYSTR + "2300'"
             //+ "          AND (S.SHEDDATE+S.SHEDTIME) <= '" + DTEDSTR + "2300'"
             //+ "          AND S.FACFLAG<>'Y'" //modify in 20150914
             //+ "          AND S.FACFLAG<>'Y' AND ((S.WKEDTIME<S.SHEDTIME) OR (S.WKSTTIME>S.SHSTTIME) OR (S.WKSTTIME IS NULL) OR (S.WKEDTIME IS NULL))"  //MARK IN 20150925
             + "          AND S.FACFLAG<>'Y' AND ((FACWH<>WORKHOUR) OR (WORKHOUR<=0))"              
             + "          AND S.STNID = F.STNID"
             + "          AND S.EMPID = F.EMPID"
             + "          AND F.FINDATE BETWEEN '" + PREDAYSTR + "' AND '" + AFTDAYSTR + "'";
        if (txtEmpID.Text.Trim() != "")
            sSQL = sSQL + "   AND S.EMPID = '" + txtEmpID.Text + "'";
        //
        sSQL = sSQL + "          AND substring(f.findate,1,4) + '-' + substring(f.findate,5,2) + '-' + substring(f.findate,7,2) + ' ' "
             + "            + substring(f.fintime,1,2) + ':' + substring(f.fintime,3,2) + ':' + substring(f.fintime,5,2)"
             + "          between dateadd(hour, -3, "
             + "                  substring(s.shstdate,1,4) + '-' + substring(s.shstdate,5,2) + '-' + substring(s.shstdate,7,2) + ' ' + substring(s.shsttime,1,2) + ':' + substring(s.shsttime,3,2) + ':00')"
             + "              and dateadd(hour, +5, "
             + "                  substring(s.sheddate,1,4) + '-' + substring(s.sheddate,5,2) + '-' + substring(s.sheddate,7,2) + ' ' + substring(s.shedtime,1,2) + ':' + substring(s.shedtime,3,2) + ':00')"
             + "       GROUP BY F.EMPID,S.SHSTDATE,S.SHSTTIME,FINDATE,FINTIME"
             + "        ) AS GROUPTABLE"
             + "   PIVOT"
             + "   (MAX(FINTIME) FOR EMPROW IN ([1],[2],[3],[4])) AS PIVOTTABLE"
             + " UNION ALL "
            //--VATM OKEY
             + " SELECT VATM_EMPLID AS EMPID,'' AS SHOWEMPID,SHSTDATE+SHSTTIME AS DATADT,'3' AS KINDID,'' AS FIELD02,VATM_VANMID AS FIELD03,VACA_NAME AS FIELD04"
             + "       ,SUBSTRING(CONVERT(char(8),VATM_DATE_ST,101),1,5)+' '+SUBSTRING(VATM_TIME_ST,1,2)+':'+SUBSTRING(VATM_TIME_ST,3,2) AS FIELD05"
             + "       ,SUBSTRING(CONVERT(char(8),VATM_DATE_EN,101),1,5)+' '+SUBSTRING(VATM_TIME_EN,1,2)+':'+SUBSTRING(VATM_TIME_EN,3,2) AS FIELD06"
             + "   FROM MP_HR.DBO.VACATM V,SCHEDM S,MP_HR.DBO.VACAMF F"
             + "  WHERE STNID='" + txtStnID.Text + "' AND EMPID=VATM_EMPLID AND VATM_VANMID=VACA_ID"
             + "   AND SHSTDATE >= '" + DTSTSTR + "'"
             + "   AND SHSTDATE <= '" + DTEDSTR + "'";
             //+ "    AND (SHSTDATE+SHSTTIME) > '" + PREDAYSTR + "2300'"
             //+ "    AND (SHSTDATE+SHSTTIME) <='" + DTEDSTR + "2300'";
        if (txtEmpID.Text.Trim() != "")
            sSQL = sSQL + "   AND VATM_EMPLID = '" + txtEmpID.Text + "'";
        //
        sSQL = sSQL + "    AND FACFLAG<>'Y' AND ((WKEDTIME<SHEDTIME) OR (WKSTTIME>SHSTTIME) OR (WKSTTIME IS NULL) OR (WKEDTIME IS NULL))"
             + "    AND ((CONVERT(char(8), VATM_DATE_ST,112)+VATM_TIME_ST BETWEEN (SHSTDATE+SHSTTIME) AND (SHEDDATE+SHEDTIME)) "
             + "      OR (CONVERT(char(8), VATM_DATE_EN,112)+VATM_TIME_EN BETWEEN (SHSTDATE+SHSTTIME) AND (SHEDDATE+SHEDTIME)))"
             + "   ) A ORDER BY EMPID,DATADT,KINDID";

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
        //modDB.InsertSignRecord("AziTest", "ok-6", User.Identity.Name);

        if (this.grdList.Rows.Count <= 0)
        {
            System.Data.DataTable VSEmptyTable = (DataTable)ViewState["EmptyTable"];
            modDB.ShowEmptyDataGridHeader(this.grdList, 0, 5, ref VSEmptyTable);
            modUtil.showMsg(this.Page, "�T��", "�d�L���`��ơI");
        }
        else
        {
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;      //* �Ω� Excel/����
        }
    }
    

}
