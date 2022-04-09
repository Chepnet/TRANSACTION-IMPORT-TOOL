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
    public partial class frmImportLoanTypes : Form
    {
        public frmImportLoanTypes()
        {
            InitializeComponent();
        }
        string filename = "";
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDailog1 = new OpenFileDialog();
            openFileDailog1.ShowDialog();
            txtFilenName .Text = openFileDailog1.FileName;
            filename = openFileDailog1.FileName;
        }
    }
}
