using BrightIdeasSoftware;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReadExcel
{
    public partial class frmSearchImportationOrder : Form
    {
        public frmSearchImportationOrder()
        {
            InitializeComponent();
        }
        int latestposition = 0;
        Classes.ProductSetup oProductSetup = new Classes.ProductSetup();
        Classes.ProductSetup onewProductSetup = null;
        Classes.FileImportFormat OFileImportFormat = new Classes.FileImportFormat();
        Classes.FileImportFormat ONewFileImportFormat = null;
        Classes.ImportFileNames oImportFilenames = new Classes.ImportFileNames();
        Classes.ImportFileNames oNewImportFileNames = null;
        private void frmSearchMigrationOrder_Load(object sender, EventArgs e)
        {
            getImportProductList();
            chkCheckAll.Checked = true;//all are checked by default
            cmbFormat.Items.Clear();
            populateCmbwithImportFileNames();

        }

        private void populateCmbwithImportFileNames()
        {
            ArrayList myList = new ArrayList();
            myList = oImportFilenames.GetImportFileNames();
            foreach (Classes.ImportFileNames oimport in myList)
            {
                // string fileformat = oimport.FormatName;
                cmbFormat.Items.Add(new ItemData.itemData(oimport.ImportFileName, oimport));
            }
        }

        private void getImportProductList()
        {
            ArrayList productList = oProductSetup.GetMigrationProducts();
            objListImportationOrder.SetObjects(productList);
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
             string error = "";
            if (cmbFormat.SelectedIndex<0)
            {
                MessageBox.Show("Import File Name Is  Required", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbFormat.Focus();
                return;
            }
           

            for (int i = 0; i < objListImportationOrder.Items.Count; i++)
            {
                if (ONewFileImportFormat == null)
                                   ONewFileImportFormat = new Classes.FileImportFormat();

                if (objListImportationOrder.Items[i].Checked)
                {
                    Classes.ProductSetup odefvalues = (Classes.ProductSetup )objListImportationOrder.GetModelObject(i);

                    // onewMemberLead.Position = objListDefaultFields.CheckedItems.Count;

                   
                    if (ONewFileImportFormat != null)
                    {
                        ONewFileImportFormat.Position = odefvalues.PositionId;
                        ONewFileImportFormat.IsLoan = odefvalues.IsLoan;
                        ONewFileImportFormat.ProductName  = odefvalues .ProductName ;
                        ONewFileImportFormat.ProductId = odefvalues.ProductId;
                        if(oNewImportFileNames !=null)
                        ONewFileImportFormat.ImportFileNameId = oNewImportFileNames .ImportFileNameId;
                        ONewFileImportFormat.FileFormatId = ONewFileImportFormat.AddFileImportFormat( ref error);
                        if (error != "")
                        {
                            break;
                        }
                        
                    }
                }
                ONewFileImportFormat = null;
            }
            if (error == "")
            {
                MessageBox.Show("Process succeeded", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                //LoadMemberLeads();
                
            }
            else
            {
                MessageBox.Show(error, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

           
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (OLVListItem olv in objListImportationOrder.Items)
            {
                olv.Checked = chkCheckAll.Checked;
            }
        }

        private void objSharetypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (objListImportationOrder.SelectedObject != null)
            {

                onewProductSetup = (Classes.ProductSetup)objListImportationOrder.SelectedObject;

            }
        }

        private void cmbFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            object obj = ((ItemData.itemData)(cmbFormat.SelectedItem))._itemData;
            oNewImportFileNames = (Classes.ImportFileNames)obj;
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {

        }
    }
}
