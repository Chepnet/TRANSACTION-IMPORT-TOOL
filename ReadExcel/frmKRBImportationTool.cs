using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReadExcel
{
    public partial class frmKRBImportationTool : Form
    {
        Classes.ProductSetup oProductSetup = new Classes.ProductSetup();
        Classes.ProductSetup onewProductSetup = null;
        Classes.LoanTypes oLoanType = new Classes.LoanTypes();
        Classes.LoanTypes oNewLoanType = null;
        Classes.ShareTypes oShareType = new Classes.ShareTypes();
        Classes.ShareTypes oNewShareType = null;
        private int ProductId = 0;
        public frmKRBImportationTool()
        {
            InitializeComponent();
        }

        private void newToolStripButton_Click(object sender, EventArgs e)
        {
            ClearTexts();

        }

        private void ClearTexts()
        {
            txtDescription.Text = "";
            txtProduct.Text = "";
            chkIsLoan.Checked = false;
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            if(txtProduct.Text =="")
            {
                MessageBox.Show("Product Is Required",this.Text,  MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtProduct.Focus();
                return;
            }
            string err = "";
            if (onewProductSetup == null)
                onewProductSetup = new Classes.ProductSetup();
            onewProductSetup.ProductName = txtProduct.Text;
            onewProductSetup.ProductId = ProductId;
            onewProductSetup.Description = txtDescription.Text;
            onewProductSetup.IsLoan = chkIsLoan.Checked;
            onewProductSetup.ProductImportId  = onewProductSetup.AddProduct( ref err);
            if (err == "")
            {
                MessageBox.Show("Process succeded", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearTexts();
            }
            else
            {
                MessageBox.Show(err, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void btnSubCounty_Click(object sender, EventArgs e)
        {

        }

        private void btnProduct_Click(object sender, EventArgs e)
        {
            if(chkIsLoan.Checked)
            {
                frmSearchLoanTypes frm = new frmSearchLoanTypes();
                frm.ShowDialog();
                oNewLoanType = oLoanType.GetLoanType(frm.selInt);
                if(oNewLoanType !=null)
                {
                    txtProduct.Text = oNewLoanType.LoanTypename;
                    txtDescription.Text = oNewLoanType.LoanTypecode;
                    ProductId = oNewLoanType.LoanTypeid;
                }
            }
            else
            {
                frmSearchShareTypes frm = new frmSearchShareTypes();
                frm.ShowDialog();
                oNewShareType  = oShareType.GetShareType(frm.selInt);
                if (oNewShareType != null)
                {
                    txtProduct.Text = oNewShareType.Sharename;
                    txtDescription.Text = oNewShareType.Sharecode;
                    ProductId = oNewShareType.Shareid;
                }
            }
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {

        }
    }
}
