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
    public partial class frmTransationImportFileName : Form
    {
        public frmTransationImportFileName()
        {
            InitializeComponent();
        }
        Classes.ImportFileNames oImportFielNames = new Classes.ImportFileNames();
        Classes.ImportFileNames oNewImportFileNames = null;

        private void frmTransationImportFileName_Load(object sender, EventArgs e)
        {

        }

        private void newToolStripButton_Click(object sender, EventArgs e)
        {
            ClearTexts();

        }

        private void ClearTexts()
        {
            txtImportFileName.Text = "";
            txtRemarks.Text = "";
            oNewImportFileNames = null;
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            if (txtImportFileName .Text == "")
            {
                MessageBox.Show("Import File Name Is Required", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtImportFileName.Focus();
                return;
            }
            string err = "";
            if (oNewImportFileNames  == null)
                oNewImportFileNames = new Classes.ImportFileNames ();
            oNewImportFileNames.ImportFileName = txtImportFileName .Text;
            oNewImportFileNames.Remarks = txtRemarks.Text;
            oNewImportFileNames.ImportFileNameId  = oNewImportFileNames.AddTransationImportFileName(ref err);
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
    }
}
