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
            else //'-----------------------------------* �ˬd�O�_�� Ū��/�s�W �v��
            {
                modUtil.GetRolData(Request, ref USER_ROL, ref USER_ROLTYPE, ref USER_ROLDEPTID);
                if (USER_ROL.Substring(0, 1) == "N") FormsAuthentication.RedirectToLoginPage();
            }
            //----------------------------------------* �]�wGridView�˦�
            modDB.SetGridViewStyle(this.grdList, 20);
            this.grdList.RowStyle.Height = 20;
            modDB.SetFields("STNNAME", "���W��"  , this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("EMPID"  , "¾��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("EMPNAME", "�m�W", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVMFS"  , "��W��(�з�+46H)", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVTW"   , "���W�ɼ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("COLLNW" , "�رХ�_�j�]�X�Լ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("COLLTT" , "�رХ�_2230��X��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("COLLTT2", "�رХ�_2230��U�Z", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("DWEEKW" , "����P����", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVWA"   , "�e��g>84H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVWB"   , "���g>84H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("OVFW"   , "�|�g>168H", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            //'----------------------------------------* �]�w����ݩ�
            txtStnID.Attributes.Add("onkeypress", "return CheckKeyNumber()");

            modUtil.SetDateObj(txtDtStart, false, null, true, null, null);
            //modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtStart, true, false, null, null, null, false);
            //modUtil.SetDateObj(txtDtFrom, false, txtDtTo, false, null, null);
            //modUtil.SetDateObj(txtDtTo, false, null, false, null, null);
            //modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtFrom, true, false, this.txtDtTo, null, null, false);
            //modUtil.SetDateImgObj(this.imgDtTo, this.txtDtTo, true, false, null, null, null, false);

            //modUtil.SetDateObj(Me.txtDtStart, False, Nothing, True);

            //'----------------------------------------* ���s��ܬd�ߵ��G
            this.lblMsg.Text = Request["msg"];

            if (Session["QryField"] == null)
            {
                DataTable VSEmptyTable = (DataTable)this.ViewState["EmptyTable"];
                modDB.ShowEmptyDataGridHeader(this.grdList, 0, 11, ref VSEmptyTable);
            }
            //'* �~�ȳB/�d����/�[�o�����e���]�w����
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
        Response.Redirect("ATT_Q_011.aspx");
    }

    //******************************************************************************************************
    //* Excel �ץX
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ011.xls", this.grdList);
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
            //sValidOK = modUtil.Check2DateObj(this, "txtDtFrom", "txtDtTo", ref sMsg, "�d�ߤ��");
            sValidOK = modUtil.CheckNotEmpty(this.txtDtStart, ref sMsg, "�d�ߤ��", true);
            //----------------------------------------* �ˮ���쥿�T��
            if ((this.txtStnID.Text == "") || (this.txtStnName.Text.Trim() == "")) 
            {
                sValidOK = false;
                sMsg = sMsg + "\n�d�L���[�o���I";
            }
            if (sValidOK)
            {
                if (txtDtStart.Text.Trim() == "") 
                {
                    sValidOK = false;
                    sMsg = sMsg + "\n �ݿ�J������!";
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
        //modUtil.showMsg(this.Page, "�T��", "��Ƨ����I");

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
        //--------------------------------------------* �[�J�~�ȳB/�d����/�[�o�����L�o
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
            modUtil.showMsg(this.Page, "�T��", "�d�L�X�Բ��`�ˮֲέp��ơI");
        }
        else
        {
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;      //* �Ω� Excel/����
        }
    }

}
