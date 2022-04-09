using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcepApp = Microsoft.Office.Interop.Excel;

namespace ReadExcel
{
    public partial class frmMigratemembersKRB : Form
    {
        public frmMigratemembersKRB()
        {
            InitializeComponent();
        }
        string filename = "";
        Classes.Member oMember = new Classes.Member();
        Classes.Member oNewMember = null;
        Classes.ShareTransactions oShareTransactions = new Classes.ShareTransactions();
        Classes.ShareTransactions oNewShareTransations = null;
        Classes.Loan oLoan = new Classes.Loan();
        Classes.Loan oNewLoan = null;
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDailog1 = new OpenFileDialog();
            openFileDailog1.ShowDialog();
            txtFilenName.Text = openFileDailog1.FileName;
            filename = openFileDailog1.FileName;
        }

        private void btnReadAndImportData_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(filename);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                oNewMember = new Classes.Member();

                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    try
                    {
                        oNewMember.Mpayroll = xlRange.Cells[i, 1].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 6] != null && xlRange.Cells[i, 6].Value2 != null)
                {
                    try
                    {
                        oNewMember.IDNO = xlRange.Cells[i, 6].Value2.ToString();
                    }
                    catch {; }
                }

                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5].ToString() != "" && xlRange.Cells[i, 5].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mcode = xlRange.Cells[i, 5].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2].ToString() != "" && xlRange.Cells[i, 2].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mfirstname = xlRange.Cells[i, 2].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3].ToString() != "" && xlRange.Cells[i, 3].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Msurname = xlRange.Cells[i, 3].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4].ToString() != "" && xlRange.Cells[i, 4].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mothername = xlRange.Cells[i, 4].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 9] != null && xlRange.Cells[i, 9].Value2 != null && xlRange.Cells[i, 9].ToString() != "" && xlRange.Cells[i, 9].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mcell = xlRange.Cells[i, 9].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 7] != null && xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7].ToString() != "" && xlRange.Cells[i, 7].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Memail = xlRange.Cells[i, 7].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8].ToString() != "" && xlRange.Cells[i, 8].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mfax = xlRange.Cells[i, 8].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null && xlRange.Cells[i, 10].ToString() != "" && xlRange.Cells[i, 10].Value2.ToString() != "")
                {
                    try
                    {
                        int employer = 0;
                        int.TryParse(xlRange.Cells[i, 10].Value2.ToString(), out employer);
                        oNewMember.Employerid =employer;
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 11] != null && xlRange.Cells[i, 11].Value2 != null && xlRange.Cells[i, 11].ToString() != "" && xlRange.Cells[i, 11].Value2.ToString() != "")
                {
                    try
                    {
                        int blocked = 0;
                        int.TryParse(xlRange.Cells[i, 11].Value2.ToString(), out blocked);
                      
                        oNewMember.Blocked = blocked;
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 12] != null && xlRange.Cells[i, 12].Value2 != null && xlRange.Cells[i, 12].ToString() != "" && xlRange.Cells[i, 12].Value2.ToString() != "")
                {
                    try
                    {

                        oNewMember.Mdate = xlRange.Cells[i, 12].Value2.ToString();
                    }
                    catch {; }
                }
                Application.DoEvents();
                oNewMember.Memberid = oNewMember.AddEditBarabaraMember(ref error);

            }
        }

        private void btnMigrateShares_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(filename);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 5; i <= 164; i++)
            {
               
                string Fullname = "";
                string memberno = "";
                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    try
                    {
                        memberno = xlRange.Cells[i, 1].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
                {
                    try
                    {
                        Fullname = xlRange.Cells[i, 5].Value2.ToString();
                    }
                    catch {; }
                }
               
                for (int x = 7; x <= 10; x++)
                {
                    oNewShareTransations = new Classes.ShareTransactions();
                    int sharetypeid = 0;
                    if (x == 7)
                        sharetypeid = 1;
                    if (x == 8)
                        sharetypeid = 2;
                    if (x == 9)
                        sharetypeid = 3;
                    if (x == 10)
                        sharetypeid = 4;
                    oNewShareTransations.Mcode = memberno;
                    oNewShareTransations.Description = Fullname;
                    if (xlRange.Cells[i, x] != null && xlRange.Cells[i, x].Value2 != null && xlRange.Cells[i, x].ToString() != "" && xlRange.Cells[i, x].Value2.ToString() != "")
                    {
                        try
                        {
                            oNewShareTransations.Amount = Convert.ToDouble(xlRange.Cells[i, x].Value2.ToString());

                        }
                        catch {; }
                    }
                    if (xlRange.Cells[i, x] != null && xlRange.Cells[i, x].Value2 != null && xlRange.Cells[i, x].ToString() != "" && xlRange.Cells[i, x].Value2.ToString() != "")
                    {
                        try
                        {
                            oNewShareTransations.SharetypeId = sharetypeid;

                        }
                        catch {; }
                    }
                    Application.DoEvents();
                    //if (oNewShareTransations.Amount != 0)
                    //{

                        oNewShareTransations.TransId = oNewShareTransations.AddEditBarabaraSharesandDeposit(ref error);
                    //}

                }





            }
        }

        private void btn_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(filename);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 12; i <= 147; i++)
            {
                oNewLoan = new Classes.Loan();

                string Fullname = "";
                string memberno = "";
                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    try
                    {
                        oNewLoan.MemberNo  = xlRange.Cells[i, 1].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                {
                    try
                    {
                        oNewLoan.MemberName  = xlRange.Cells[i,2].Value2.ToString();
                    }
                    catch {; }
                }

               
                    if (xlRange.Cells[i, 6] != null && xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6].ToString() != "" && xlRange.Cells[i, 6].Value2.ToString() != "")
                    {
                        try
                        {
                            oNewLoan.Loantransdate =xlRange.Cells[i, 6].Value2.ToString();

                        }
                        catch {; }
                    }
                    if (xlRange.Cells[i, 7] != null && xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7].ToString() != "" && xlRange.Cells[i, 7].Value2.ToString() != "")
                    {
                        try
                        {
                        int period = 0;
                        int.TryParse(xlRange.Cells[i, 7].Value2.ToString(), out period);
                            oNewLoan.Period = period; 

                        }
                        catch {; }
                    }
                if (xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null && xlRange.Cells[i, 8].ToString() != "" && xlRange.Cells[i, 8].Value2.ToString() != "")
                {
                    try
                    {
                        double amount = 0;
                        double.TryParse(xlRange.Cells[i, 8].Value2.ToString(), out amount);
                        oNewLoan.Loanamount = amount;

                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 11] != null && xlRange.Cells[i, 11].Value2 != null && xlRange.Cells[i, 11].ToString() != "" && xlRange.Cells[i, 11].Value2.ToString() != "")
                {
                    try
                    {
                        int loantypeid = 0;
                        int.TryParse(xlRange.Cells[i, 11].Value2.ToString(), out loantypeid);
                        oNewLoan.LoanTypeId = loantypeid;

                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 12] != null && xlRange.Cells[i, 12].Value2 != null && xlRange.Cells[i, 12].ToString() != "" && xlRange.Cells[i, 12].Value2.ToString() != "")
                {
                    try
                    {
                        double paidamount = 0;
                        double.TryParse(xlRange.Cells[i, 12].Value2.ToString(), out paidamount);
                        oNewLoan.Paymentamount = paidamount;

                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 9] != null && xlRange.Cells[i, 9].Value2 != null && xlRange.Cells[i, 9].ToString() != "" && xlRange.Cells[i, 9].Value2.ToString() != "")
                {
                    try
                    {
                        double installment = 0;
                        double.TryParse(xlRange.Cells[i, 9].Value2.ToString(), out installment);
                        oNewLoan.OriginalAmount = installment;

                    }
                    catch {; }
                }

                Application.DoEvents();
                    //if (oNewShareTransations.Amount != 0)
                    //{

                    oNewLoan.LoanId = oNewLoan.AddEditKRBLoan(ref error);
                    //}

                





            }
        }
    }
}
