using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using ScriptsExecutionUtility.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ScriptsExecutionUtility
{
    public partial class InvoicePrerequisite : Form
    {
        string ConnectionString = ConfigurationManager.AppSettings["Bill800ConnectionString"];
        DataBaseConnectivity db = new DataBaseConnectivity();
        ExcelExport ep=new ExcelExport();
        DataTable dbresponse = new DataTable();
        public InvoicePrerequisite()
        {
         
           
           InitializeComponent();
            loadBillCycle("preScript");
            loadBillCycle("ReportStrip");
            Action_btn.Visible = false;
            joblevel.Visible = false;
            JobBox.Visible = false;
            lblcount.Visible = false;
            loaderimg.Visible = false;
            this.MaximizeBox = false;
            // btnExport.Visible = false;
        }

      
     

        private async void Execbtn_Click(object sender, EventArgs e)
        {
            string ReqType = radiobtngrp.Controls.OfType<System.Windows.Forms.RadioButton>().FirstOrDefault(r => r.Checked).Text.ToString();
            string BillCycleid = billPeriodcombobx.SelectedValue.ToString();
            loaderimg.Visible = true;
            DisableEvent();
            var res= await GetData(ReqType, BillCycleid, "ScriptPreReq");
            EnableEvent();
            if (res.Rows.Count == 0)
            {
                Alldatagridview.Columns.Clear();
                MessageBox.Show("Not found");
            }
            else
            {
                switch (ReqType)
                {
                    case "Sage Open Invoices":
                        {
                            if (res.Rows.Count > 0)
                            {

                                Alldatagridview.DataSource = res;
                                lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                lblcount.Visible = true;
                            }
                            break;
                        }
                    case "Sage Open Adjustment":
                        {
                            if (res.Rows.Count > 0)
                            {
                                Alldatagridview.DataSource = res;
                                lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                lblcount.Visible = true;
                            }
                            break;
                        }
                    case "Move Dissconnected Dial Number":
                        {
                            if (res.Rows.Count > 0)
                            {
                                Alldatagridview.DataSource = res;
                                lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                lblcount.Visible = true;
                            }
                            break;
                        }
                    case "Negative Invoice Balance":
                        {
                            if (res.Rows.Count > 0)
                            {
                                Alldatagridview.DataSource = res;
                                lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                lblcount.Visible = true;
                            }
                            break;
                        }
                    case "Merge Job ID":
                        {
                            if (res.Rows.Count > 0)
                            {
                                Alldatagridview.DataSource = res;
                                lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                lblcount.Visible = true;
                                List<string> list = new List<string>();
                                for (int i = 0; i < res.Rows.Count - 1; i++)
                                {
                                    list.Add(res.Rows[i]["JobID"].ToString());
                                }
                                if (list.Count > 0)
                                {
                                    JobBox.DataSource = list;
                                    JobBox.DropDownStyle = ComboBoxStyle.DropDownList;
                                }
                            }

                            break;
                        }
                    case "Validate the Difference":
                        {
                            if (res.Rows.Count > 0)
                            {
                                Alldatagridview.DataSource = res;
                                lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                lblcount.Visible = true;

                                if (res.Rows[0].ItemArray[2].ToString() != "")
                                {
                                    Action_btn.Visible = true;
                                    Action_btn.Text = "Merge Difference";
                                }

                            }
                            break;
                        }
                    case "Multiple Occurrences":
                        {
                            Alldatagridview.ReadOnly = true;
                            if (res.Rows.Count > 0)
                            {

                                Alldatagridview.DataSource = res;
                                DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                                checkBoxColumn.HeaderText = "Select"; // Set the header text
                                checkBoxColumn.Name = "SelectedAccount";
                                Alldatagridview.Columns.Add(checkBoxColumn);
                                Alldatagridview.ReadOnly = false;
                                lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                lblcount.Visible = true;
                                Action_btn.Visible = true;
                                Action_btn.Text = "Delete Duplicate Invoice";
                                JobBox.Visible = false;
                            }

                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }
            
        }
        public async Task<DataTable> GetData(string ReqType, string BillCycleid,string procName)
        {
            try
            {
                DataTable res = new DataTable();
                if (BillCycleid == "")
                {
                    MessageBox.Show("Please Select Date");
                }
                else
                {
                    var parm = new
                    {
                        BillCycleID = BillCycleid,
                        ReqType = ReqType
                    };
                    await Task.Run(() =>
                    {
                        res = db.ExecuteProc(procName, db.returnSppram(parm));
                    });

                }


                return res;
            }
            catch (Exception)
            {

               return null;
            }
        
        }


        private void btnExport_Click(object sender, EventArgs e)
        {
           
            if (Alldatagridview.Rows.Count > 0)
            {
                
               var res= ep.ExporttoExcel(Alldatagridview);
                if (res == true)
                {
                    MessageBox.Show("Report Successfully Generated");
                }
                else
                {
                    MessageBox.Show("Fail to Generate Report");
                }
                
            }
        }

        private void btnexit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void SageOpenInvoices_CheckedChanged(object sender, EventArgs e)
        {
            if (SageOpenInvoices.Checked == true)
            {
                lblcount.Visible = false;
                Execbtn.Visible = true;
                Alldatagridview.Columns.Clear();
                loadBillCycle("preScript");
                JobBox.Visible = false;
            }
            
        }
        private void loadBillCycle(string strip)
        {
            var data = db.ConverttoObject(db.ExecuteProc("getBillingcycle", null), typeof(BillPeriodDto)).Cast<BillPeriodDto>().ToList();
            if (strip == "preScript")
            {
               
                billPeriodcombobx.DataSource = data;
                billPeriodcombobx.DropDownStyle = ComboBoxStyle.DropDownList;
                billPeriodcombobx.DisplayMember = "BillDate";
                billPeriodcombobx.ValueMember = "BillCycleID";
            

            }
            else if (strip == "ReportStrip")
            {

                ReportscomboBox.DataSource = data;
                ReportscomboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                ReportscomboBox.DisplayMember = "BillDate";
                ReportscomboBox.ValueMember = "BillCycleID";
            
            }
            
        }

     
        private void sageAdjustment_CheckedChanged(object sender, EventArgs e)
        {
            if (sageAdjustment.Checked == true)
            {
                lblcount.Visible = false;
                Execbtn.Visible = true;
                Alldatagridview.DataSource = null;
                loadBillCycle("preScript");
            }
        }

        private void moveddisconnectdailnumber_CheckedChanged(object sender, EventArgs e)
        {
            inputbxlbl.Text = "Select Bill Cycle";
            if (moveddisconnectdailnumber.Checked==true)
             {
                inputbxlbl.Text = "Select Delete Date";
                lblcount.Visible = false;
                var date = DateTime.Now.AddMonths(-5).Month.ToString()+ "/14/" + DateTime.Now.Year.ToString();
                List<string> Dissconnecteddate = new List<string>();
                Dissconnecteddate.Add(date.ToString());
                billPeriodcombobx.DataSource = Dissconnecteddate;
                billPeriodcombobx.DropDownStyle = ComboBoxStyle.DropDownList;
      
            }
           
        }

        private void negativeinvoicebalanace_CheckedChanged(object sender, EventArgs e)
        {
          

            if (negativeinvoicebalanace.Checked == true)
            {
                lblcount.Visible = false;
                Action_btn.Visible = false;
                Alldatagridview.Columns.Clear();
                Action_btn.Text = "Update Negative Balance";
                loadBillCycle("preScript");
                
            }
        }

    
        private void mergejobid_CheckedChanged(object sender, EventArgs e)
        {
            if (mergejobid.Checked == true)
            {
                Alldatagridview.Columns.Clear();
                lblcount.Visible = false;
                joblevel.Visible = false;
                Action_btn.Visible = false;
                Action_btn.Text = "Merge Job";
                loadBillCycle("preScript");
            }
     
        }

        private void validatediffer_CheckedChanged(object sender, EventArgs e)
        {
            if (validatediffer.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("preScript");
            }
        }

        private void MultipleAccuranance_CheckedChanged(object sender, EventArgs e)
        {
            Alldatagridview.ReadOnly = true;
            if (MultipleAccuranance.Checked == true)
            {
                Alldatagridview.Columns.Clear();
                lblcount.Visible = false;
                loadBillCycle("preScript");
            }
        }

        private async void ReportsExecbtn_Click(object sender, EventArgs e)
        {
          
            dbresponse = null;
        
            string ReqType = Reportfroupbx.Controls.OfType<System.Windows.Forms.RadioButton>().FirstOrDefault(r => r.Checked).Text.ToString();
            string BillCycleid = ReportscomboBox.SelectedValue.ToString();
            
            //string BillCycleid = "217";
            if (BillCycleid == "")
            {
                Alldatagridview.Columns.Clear();
                MessageBox.Show("Please Select Date");
            }
            else
            {
                var parm = new
                {
                    BillCycleID = BillCycleid,
                    ReqType = ReqType
                };
             
                DisableEvent();
                var res = await GetData(ReqType, BillCycleid, "Invoicereports");
                 EnableEvent();
                if (res == null)
                {
                    MessageBox.Show("No Record Found");
                }
               else if (res.Rows.Count > 0)
                {
                    switch (ReqType)
                    {
                        case "Invoice Balance Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "Dunning Past due Report DL1":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "Dunning Past due Report DL2":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "International DL2 Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            } 
                        case "DL Statistic":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "DL Statistics 6 Year Counts Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "DL Statistics DL1 Details Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "DL Statistics DL2 Details Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "Disconnection Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "FCC Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "Dunning Summary Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        case "Wallet Report":
                            {
                                if (res.Rows.Count > 0)
                                {

                                    Alldatagridview.DataSource = res;
                                    lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                                    lblcount.Visible = true;
                                }
                                break;
                            }
                        default:
                            {
                                break;
                            }
                    }
                }
                else
                {
                    MessageBox.Show("No Record Found");
                }

            }
        }

        private void Action_btn_Click(object sender, EventArgs e)
        {
            string ReqType = radiobtngrp.Controls.OfType<System.Windows.Forms.RadioButton>().FirstOrDefault(r => r.Checked).Text.ToString();

            switch (ReqType)
            {
                case "Negative Invoice Balance":
                    {
                        string BillCycleid = billPeriodcombobx.SelectedValue.ToString();
                        if (BillCycleid == "")
                        {
                            MessageBox.Show("Please Select Date");
                        }
                        else
                        {
                            var parm = new
                            {
                                BillCycleID = billPeriodcombobx.SelectedValue.ToString(),
                                ReqType = "Update Negative Invoice Balance"
                            };
                            var res = db.ExecuteProc("ScriptPreReq", db.returnSppram(parm));
                            Alldatagridview.DataSource = res;
                            lblcount.Text = "Total Records " + res.Rows.Count.ToString();
                        }
                        break;
                    }
                case "Merge Job ID":
                    {
                        if (JobBox.Items.Count <= 1)
                        {
                            MessageBox.Show("You have only one job");
                        }
                        else
                        {
                            string selectedjob = JobBox.SelectedValue.ToString();
                            var parm = new
                            {
                                BillCycleID = billPeriodcombobx.SelectedValue.ToString(),
                                ReqType = "Update Merge Job ID",
                                jobid = selectedjob
                            };
                            var res = db.ExecuteProc("ScriptPreReq", db.returnSppram(parm));
                            MessageBox.Show("Successfully Updated");

                        }
                        break;
                    }
                case "Validate the Difference":
                    {
                        if (Alldatagridview.Rows.Count < 1)
                        {
                            MessageBox.Show("NO Record Found");

                        }
                        else
                        {
                            var parm = new
                            {
                                BillCycleID = billPeriodcombobx.SelectedValue.ToString(),
                                ReqType = "Update Difference",

                            };
                            var res = db.ExecuteProc("ScriptPreReq", db.returnSppram(parm));
                            MessageBox.Show("Successfully Updated");
                            Action_btn.Visible = false;
                            Alldatagridview.Columns.Clear();
                        }
                        break;
                    }
                case "Multple Accurance":
                    {
                        List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();

                        foreach (DataGridViewRow row in Alldatagridview.Rows)
                        {
                            // Assuming the checkbox column name is "NewColumn"
                            DataGridViewCheckBoxCell checkBoxCell = row.Cells["SelectedAccount"] as DataGridViewCheckBoxCell;

                            if (checkBoxCell != null && Convert.ToBoolean(checkBoxCell.Value) == true)
                            {
                                selectedRows.Add(row);
                            }
                        }
                        if (selectedRows.Count > 0)
                        {
                            List<string> accountlsit = new List<string>();

                            foreach (var item in selectedRows)
                            {
                                accountlsit.Add(Alldatagridview.Rows[item.Index].Cells[0].Value.ToString());
                            }
                            accountlsit = accountlsit.Distinct().ToList();
                            for (int i = 0; i < accountlsit.Count; i++)
                            {
                                var parm = new
                                {
                                    BillCycleID = billPeriodcombobx.SelectedValue.ToString(),
                                    ReqType = "Delete Accounts",
                                    jobid = accountlsit[i]
                                };
                                var res = db.ExecuteProc("ScriptPreReq", db.returnSppram(parm));
                            }

                        }
                        else
                        {
                            MessageBox.Show("No Account Selected");
                        }
                        Action_btn.Visible = false;
                        Alldatagridview.Columns.Clear();
                        break;
                    }
            }
        }

        private void InvBln_radiobtn_CheckedChanged(object sender, EventArgs e)
        {
            if (InvBln_radiobtn.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("ReportStrip");
                JobBox.Visible = false;
            }
        }

        private void duningsumrybtn_CheckedChanged(object sender, EventArgs e)
        {
            if (duningsumrybtn.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("ReportStrip");
                JobBox.Visible = false;
            }
        }

        private void Dstatistic_CheckedChanged(object sender, EventArgs e)
        {
            ReportscomboBox.Visible = true;
            selectlabel.Visible = true;
            if (Dstatistic.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("ReportStrip");
                ReportscomboBox.Visible = false;
                selectlabel.Visible = false;
                JobBox.Visible = false;
            }
        }

        private void staticscountrepo_CheckedChanged(object sender, EventArgs e)
        {
            ReportscomboBox.Visible = true;
            selectlabel.Visible = true;
            if (staticscountrepo.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("ReportStrip");
                ReportscomboBox.Visible = false;
                selectlabel.Visible = false;
                JobBox.Visible = false;
            }
              
        }

        private void detailsstaticsrepo_CheckedChanged(object sender, EventArgs e)
        {
            ReportscomboBox.Visible = true;
            selectlabel.Visible = true;
            if (detailsstaticsrepo.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                ReportscomboBox.Visible = false;
                selectlabel.Visible = false;
                loadBillCycle("ReportStrip");
                JobBox.Visible = false;
            }
        }

        private void detailsstaticsrepodl2_CheckedChanged(object sender, EventArgs e)
        {
            ReportscomboBox.Visible = true;
            selectlabel.Visible = true;
            if (detailsstaticsrepodl2.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("ReportStrip");
                ReportscomboBox.Visible = false;
                selectlabel.Visible = false;
                JobBox.Visible = false;
            }
        }

        private void dunpasrptdl1_CheckedChanged(object sender, EventArgs e)
        {
            if (dunpasrptdl1.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("ReportStrip");
                JobBox.Visible = false;
            }
        }

        private void dunpasrptdl2_CheckedChanged(object sender, EventArgs e)
        {
            if (dunpasrptdl2.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("ReportStrip");
                JobBox.Visible = false;
            }
        }

        private void IntDl2rpt_CheckedChanged(object sender, EventArgs e)
        {
            if (IntDl2rpt.Checked == true)
            {
                lblcount.Visible = false;
                Alldatagridview.Columns.Clear();
                loadBillCycle("ReportStrip");
                JobBox.Visible = false;
            }
        }
        private void DisableEvent()

        {
            loaderimg.Visible = true;
            SageOpenInvoices.Enabled = false;
            sageAdjustment.Enabled = false;
            moveddisconnectdailnumber.Enabled = false;
            negativeinvoicebalanace.Enabled = false;
            mergejobid.Enabled = false;
            MultipleAccuranance.Enabled = false;
            validatediffer.Enabled = false;
            InvBln_radiobtn.Enabled = false;
            duningsumrybtn.Enabled = false;
            dunpasrptdl1.Enabled = false;
            dunpasrptdl2.Enabled = false;
            IntDl2rpt.Enabled = false;
            Disconradiobtn.Enabled=false;
            Dstatistic.Enabled = false;
            staticscountrepo.Enabled = false;
            detailsstaticsrepo.Enabled = false;
            detailsstaticsrepodl2.Enabled = false;
            wltrep.Enabled = false;
            fccrepo.Enabled = false;
            billPeriodcombobx.Enabled = false;
            ReportsExecbtn.Enabled = false;
            ReportscomboBox.Enabled = false;
            JobBox.Enabled = false;
            lblcount.Enabled = false;
            btnexit.Enabled = false;
            btnExport.Enabled = false;
            Execbtn.Enabled = false;
        }
        private void EnableEvent()

        {
            loaderimg.Visible = false;
            SageOpenInvoices.Enabled = true;
            sageAdjustment.Enabled = true;
            moveddisconnectdailnumber.Enabled = true;
            negativeinvoicebalanace.Enabled = true;
            mergejobid.Enabled = true;
            MultipleAccuranance.Enabled = true;
            validatediffer.Enabled = true;
            InvBln_radiobtn.Enabled = true;
            duningsumrybtn.Enabled = true;
            dunpasrptdl1.Enabled = true;
            dunpasrptdl2.Enabled = true;
            IntDl2rpt.Enabled = true;
            Disconradiobtn.Enabled = true;
            Dstatistic.Enabled = true;
            staticscountrepo.Enabled = true;
            detailsstaticsrepo.Enabled = true;
            detailsstaticsrepodl2.Enabled = true;
            wltrep.Enabled = true;
            fccrepo.Enabled = true;
            billPeriodcombobx.Enabled = true;
            ReportsExecbtn.Enabled=true;
            ReportscomboBox.Enabled=true;
            JobBox.Enabled = true;
            lblcount.Enabled = true;
            btnexit.Enabled = true;
            btnExport.Enabled = true;
            Execbtn.Enabled = true;
        }

        private void Alldatagridview_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void billPeriodcombobx_SelectedIndexChanged(object sender, EventArgs e)
        {
            Alldatagridview.Columns.Clear();
            lblcount.Visible = false;
        }

        private void wltrep_CheckedChanged(object sender, EventArgs e)
        {
            ReportscomboBox.Visible = true;
            selectlabel.Visible = true;
         
              
                if (wltrep.Checked == true)
                {
                    lblcount.Visible = false;
                    Alldatagridview.Columns.Clear();
                    ReportscomboBox.Visible = false;
                    selectlabel.Visible = false;
                    loadBillCycle("ReportStrip");
                    JobBox.Visible = false;
                }
            
        }

        private void Disconradiobtn_CheckedChanged(object sender, EventArgs e)
        {
            ReportscomboBox.Visible = true;
            selectlabel.Visible = true;
            if (Disconradiobtn.Checked == true)
            {

                    lblcount.Visible = false;
                    Alldatagridview.Columns.Clear();
                    ReportscomboBox.Visible = false;
                    selectlabel.Visible = false;
                    loadBillCycle("ReportStrip");
                    JobBox.Visible = false;
                        }
        }
    }
}
