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

//20130905_��Ū���F��������H�K�y��CONNECTION�W���A���HSQL PROCEDURE�B�z���C

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
            else //'-----------------------------------* �ˬd�O�_�� Ū��/�s�W �v��
            {
                modUtil.GetRolData(Request, ref USER_ROL, ref USER_ROLTYPE, ref USER_ROLDEPTID);
                if (USER_ROL.Substring(0, 1) == "N") FormsAuthentication.RedirectToLoginPage();
            }
            //----------------------------------------* �]�wGridView�˦�
            modDB.SetGridViewStyle(this.grdList, 20);
            this.grdList.RowStyle.Height = 20;
            modDB.SetFields("EMPID"   , "���u¾��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("EMPNAME" , "���u�m�W", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("STNID", "�X���N��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("STNNAME", "�X�����W", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDA", "����X��/�X�����", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDB", "�߯Z�X��/�_�l�ɶ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDC", "�]���X��/����ɶ�", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDD", "����X��/�ɶ�(��)", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            modDB.SetFields("FIELDE", "�ɶ�����/�k��", this.grdList, HorizontalAlign.Center, "", false, 0, false);
            
            //'--------------------------------------------------* �[�J�s�����
            this.grdList.RowStyle.Height = 20;

            //'----------------------------------------* �]�w����ݩ�
            
            modUtil.SetDateObj(txtDtFrom, false, txtDtTo, false, null, null);
            modUtil.SetDateObj(txtDtTo, false, null, false, null, null);
            modUtil.SetDateImgObj(this.imgDtFrom, this.txtDtFrom, true, false, this.txtDtTo, null, null, false);
            modUtil.SetDateImgObj(this.imgDtTo, this.txtDtTo, true, false, null, null, null, false);

            //'----------------------------------------* ���s��ܬd�ߵ��G
            this.lblMsg.Text = Request["msg"];

            //'* �~�ȳB/�d����/�[�o�����e���]�w����
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
    //* ���s��ܬd�ߵ��G 
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
    //* ��ܭ��u�m�W
    //******************************************************************************************************

    protected void txtEmpID_TextChanged(object sender, EventArgs e)
    {
         //modCharge.GetStnName(txtStnID, this.txtStnName, this.ddlArea.SelectedValue);
         modCharge.GetEmpName(txtEmpID, this.txtEmpName,true);
    }

    //******************************************************************************************************
    //* ��в��ʮɡA��ܥ����ĪG
    //******************************************************************************************************
    protected void grdList_RowDataBound(object sender, System.Web.UI.WebControls.GridViewRowEventArgs e)
    {
        modDB.SetGridLightPen(e);
    }

    //******************************************************************************************************
    //* �M��
    //******************************************************************************************************
    protected void btnClear_Click(object sender, EventArgs e)
    {
        Session["QryField"]= null;
        Session["iCurPage"] = null;
        Response.Redirect("ATT_Q_019.aspx");
    }

    //******************************************************************************************************
    //* Excel �ץX
    //******************************************************************************************************
    protected void btnExcel_Click(object sender,EventArgs e)
    {
        modUtil.GridView2Excel("ATTQ019.xls", this.grdList);
    }

    protected void btnQry_Click(object sender,EventArgs e)
    {
        string sMsg = "";
        Boolean sValidOK = true;
        //�n����d�毸+31�Ѫ����
        try
        {
            sValidOK = true;
            sValidOK = modUtil.Check2DateObj(this, "txtDtFrom", "txtDtTo", ref sMsg, "�d�ߤ��");
            //----------------------------------------* �ˮ���쥿�T��

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
            sSQL = sSQL + ",'1'"; //�έp��
        }
        else
        {
            sSQL = sSQL + ",'2'"; //���Ӫ�
        }

        //modDB.InsertSignRecord("AziTest", sSQL, User.Identity.Name); //azitest

        //modUnset.PA_ExcuteCommand(sSQL);

        //--------------------------------------------* �[�J�~�ȳB/�d����/�[�o�����L�o
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
            modUtil.showMsg(this.Page, "�T��", "�d�L��ơI");
        }
        else
        {
            this.btnExcel.Enabled = true;
            ViewState["Sql"] = sSQL;      //* �Ω� Excel/����
        }
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
