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
    public partial class frmSearchShareTypes : Form
    {
        public frmSearchShareTypes()
        {
            InitializeComponent();
        }
        Classes.ShareTypes oShareTypes = new Classes.ShareTypes();
        Classes.ShareTypes oNewShareTypes = null;
        public int selInt = 0;
        private void frmSearchShareTypes_Load(object sender, EventArgs e)
        {
            ArrayList myList = oShareTypes.GetShareTypes();
            objSharetypes.SetObjects(myList);
        }

        private void objSharetypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(objSharetypes.SelectedObject!=null)
            {
                oNewShareTypes = (Classes.ShareTypes)objSharetypes.SelectedObject;
                if(oNewShareTypes !=null)
                {
                    this.selInt = oNewShareTypes.Shareid;
                }
            }
        }

        private void objSharetypes_DoubleClick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
