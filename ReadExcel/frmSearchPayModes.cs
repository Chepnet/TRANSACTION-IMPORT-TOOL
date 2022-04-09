using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;

namespace ReadExcel
{
    public partial class frmSearchPayModes : Form
    {
        public frmSearchPayModes()
        {
            InitializeComponent();

            
        }
        Classes.PayModes oPayModes = new Classes.PayModes();
        Classes.PayModes oNewPayModes = null;
        public int selInt = 0;
        private void frmSearchPayModes_Load(object sender, EventArgs e)
        {
                ArrayList myList =  oPayModes.GetPayModes();
            objListPayMode.SetObjects(myList);

        }

        private void objListPayMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (objListPayMode != null)
                oNewPayModes = (Classes.PayModes)objListPayMode.SelectedObject;
            if (oNewPayModes != null)
                this.selInt = oNewPayModes.PaymentModeId;
        }

        private void objListPayMode_DoubleClick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
