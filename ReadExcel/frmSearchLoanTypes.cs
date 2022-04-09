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
    public partial class frmSearchLoanTypes : Form
    {
        public frmSearchLoanTypes()
        {
            InitializeComponent();
        }
        Classes.LoanTypes oLoanType = new Classes.LoanTypes();
        Classes.LoanTypes oNewLoanType = null;
        public int selInt = 0;
        private void objSharetypes_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (objLoantypes.SelectedObject != null)
            {
                oNewLoanType = (Classes.LoanTypes)objLoantypes.SelectedObject;
                if (oNewLoanType != null)
                {
                    this.selInt = oNewLoanType.LoanTypeid;
                }
            }
        }

        private void frmSearchLoanTypes_Load(object sender, EventArgs e)
        {
            ArrayList myList = oLoanType .GetLoanTypes ();
            objLoantypes.SetObjects(myList);
        }

        private void objLoantypes_DoubleClick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
