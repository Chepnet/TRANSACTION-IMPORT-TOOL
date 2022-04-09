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
    public partial class frmNascaMigration : Form
    {
        public frmNascaMigration()
        {
            InitializeComponent();
        }
        public string FileName = "";
        Classes.Member oMember = new Classes.Member();
        Classes.Member oNewMember = null;
        Classes.ShareTransactions oShareTransactions = new Classes.ShareTransactions();
        Classes.ShareTransactions oNewShareTransations = null;
        Classes.Loan oLoan = new Classes.Loan();
        Classes.Loan oNewLoan = null;
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            txtFileName.Text = openFileDialog.FileName;
            FileName = openFileDialog.FileName;
        }

        private void btnReadFile_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(FileName);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                oNewMember = new Classes.Member();

                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null)
                {
                    try
                    {
                        oNewMember.IDNO = xlRange.Cells[i, 5].Value2.ToString();
                    }
                    catch {; }
                }

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1].ToString() != "" && xlRange.Cells[i, 1].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mcode = xlRange.Cells[i, 1].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3].ToString() != "" && xlRange.Cells[i, 3].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mfirstname = xlRange.Cells[i, 3].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2].ToString() != "" && xlRange.Cells[i, 2].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Msurname = xlRange.Cells[i, 2].Value2.ToString();
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
                if (xlRange.Cells[i, 6] != null && xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6].ToString() != "" && xlRange.Cells[i, 6].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mcell = xlRange.Cells[i, 6].Value2.ToString();
                    }
                    catch {; }
                }
                Application.DoEvents();
                oNewMember.Memberid = oNewMember.AddEditBarabaraMember(ref error);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(FileName);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 4; i <= 62; i++)
            {

                string Fullname = "";
                string memberno = "";
                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                {
                    try
                    {
                        memberno = xlRange.Cells[i, 2].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3].ToString() != "" && xlRange.Cells[i, 3].Value2.ToString() != "")
                {
                    try
                    {
                        Fullname = xlRange.Cells[i, 3].Value2.ToString();
                    }
                    catch {; }
                }
                for (int x = 4; x <= 15; x++)
                {
                    oNewShareTransations = new Classes.ShareTransactions();
                    oNewShareTransations.Description = Fullname;
                    oNewShareTransations.Mcode = memberno;
                    if (x == 4)
                    {
                        oNewShareTransations.TransDate = "20210131";
                    }
                    else if (x == 5)
                    {
                        oNewShareTransations.TransDate = "20210228";
                    }
                    else if (x == 6)
                    {
                        oNewShareTransations.TransDate = "20210331";
                    }
                    else if (x == 7)
                    {
                        oNewShareTransations.TransDate = "20210430";
                    }
                    else if (x == 8)
                    {
                        oNewShareTransations.TransDate = "20210531";
                    }
                    else if (x == 9)
                    {
                        oNewShareTransations.TransDate = "20210630";
                    }
                    else if (x == 10)
                    {
                        oNewShareTransations.TransDate = "20210731";
                    }
                    else if (x == 11)
                    {
                        oNewShareTransations.TransDate = "20210831";
                    }
                    else if (x == 12)
                    {
                        oNewShareTransations.TransDate = "20210930";
                    }
                    else if (x == 13)
                    {
                        oNewShareTransations.TransDate = "20211031";
                    }
                    else if (x == 14)
                    {
                        oNewShareTransations.TransDate = "20211130";
                    }
                    else if (x == 15)
                    {
                        oNewShareTransations.TransDate = "20211231";
                    }

                    if (xlRange.Cells[i, x] != null && xlRange.Cells[i, x].Value2 != null && xlRange.Cells[i, x].ToString() != "" && xlRange.Cells[i, x].Value2.ToString() != "")
                    {
                        try
                        {
                            oNewShareTransations.Amount = Convert.ToDouble(xlRange.Cells[i, x].Value2.ToString());
                            if (x == 15)
                            {
                                oNewShareTransations.Amount = -(Convert.ToDouble(xlRange.Cells[i, x].Value2.ToString()));
                            }
                        }
                        catch {; }
                    }
                    Application.DoEvents();
                    if (oNewShareTransations.Amount != 0)
                    {

                        oNewShareTransations.TransId = oNewShareTransations.AddEditBarabaraSharesandDeposit(ref error);
                    }
                }






            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(FileName);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[6];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            string TransDate = "";
            for (int i = 3; i <= 71; i++)
            {
                oNewLoan = new Classes.Loan();


                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 22] != null && xlRange.Cells[i, 22].Value2 != null)
                {
                    try
                    {
                        TransDate = xlRange.Cells[i, 22].Value2.ToString();
                    }
                    catch {; }
                }

                oNewLoan.Loantransdate = TransDate;


                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2].ToString() != "" && xlRange.Cells[i, 2].Value2.ToString() != "")
                {
                    try
                    {
                        oNewLoan.MemberNo = xlRange.Cells[i, 2].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3].ToString() != "" && xlRange.Cells[i, 3].Value2.ToString() != "")
                {
                    try
                    {
                        oNewLoan.MemberName = xlRange.Cells[i, 3].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4].ToString() != "" && xlRange.Cells[i, 4].Value2.ToString() != "")
                {
                    try
                    {
                        oNewLoan.Loanamount = Convert.ToDouble(xlRange.Cells[i, 4].Value2.ToString());
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5].ToString() != "" && xlRange.Cells[i, 5].Value2.ToString() != "")
                {
                    try
                    {
                        oNewLoan.Period = Convert.ToDouble(xlRange.Cells[i, 5].Value2.ToString());
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 11] != null && xlRange.Cells[i, 11].Value2 != null && xlRange.Cells[i, 11].ToString() != "" && xlRange.Cells[i, 11].Value2.ToString() != "")
                {
                    try
                    {
                        oNewLoan.Interest = Convert.ToDouble(xlRange.Cells[i, 11].Value2.ToString());
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 7] != null && xlRange.Cells[i, 7].Value2 != null && xlRange.Cells[i, 7].ToString() != "" && xlRange.Cells[i, 7].Value2.ToString() != "")
                {
                    try
                    {
                        oNewLoan.MonthlyInterest = Convert.ToDouble(xlRange.Cells[i, 7].Value2.ToString());
                    }
                    catch {; }
                }
                oNewLoan.LoanTypeId = 6;
                Application.DoEvents();
                if (oNewLoan.Loanamount > 0)
                {
                    oNewLoan.LoanId = oNewLoan.AddEditNascaLoan(ref error);
                }


            }
        }

        private void btnExpense_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(FileName);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            int parentid = 0;
            int account = 0;
            for (int i = 5; i <= 61; i++)
            {

               
                string accountname = "";
               
                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                {
                    try
                    {
                        accountname = xlRange.Cells[i, 2].Value2.ToString();//accountname
                    }
                    catch {; }
                }
                if (accountname == "ADMINISTRATION EXPENSES")
                {
                    parentid = 510;
                    account = parentid;
                }
                if (accountname == "GOVERNANCE EXPENSES")
                {
                    parentid = 520;
                    account = parentid;

                }
                if (accountname == "STAFF EXPENSES")
                {
                    parentid = 530;
                    account = parentid;

                }
                if (accountname == "PROFESSIONAL FEES")
                {
                    parentid = 540;
                    account = parentid;

                }
                if (accountname == "PROFESSIONAL FEES")
                {
                    parentid = 550;
                    account = parentid;

                }
                if (accountname == "INTEREST EXPENSE")
                {
                    parentid = 560;
                    account = parentid;

                }
                if (accountname == "OTHER EXPENSES")
                {
                    parentid = 570;
                    account = parentid;



                }
                else
                {
                    account = account + 1;
                    


                    for (int x = 3; x <= 15; x++)
                    {
                        oNewMember = new Classes.Member();
                        oNewMember.Mcode = accountname;
                        oNewMember.Employerid = account;
                        oNewMember.Stationid = parentid;
                        
                     
                        if (x == 3)
                        {
                            oNewMember.Mpayroll = "20210131";
                        }
                        else if (x == 4)
                        {
                            oNewMember.Mpayroll = "20210228";
                        }
                        else if (x == 5)
                        {
                            oNewMember.Mpayroll = "20210331";
                        }
                        else if (x == 6)
                        {
                            oNewMember.Mpayroll = "20210430";
                        }
                        else if (x == 7)
                        {
                            oNewMember.Mpayroll = "20210531";
                        }
                        else if (x == 8)
                        {
                            oNewMember.Mpayroll = "20210630";
                        }
                        else if (x == 9)
                        {
                            oNewMember.Mpayroll = "20210731";
                        }
                        else if (x == 10)
                        {
                            oNewMember.Mpayroll = "20210831";
                        }
                        else if (x == 11)
                        {
                            oNewMember.Mpayroll = "20210930";
                        }
                        else if (x == 12)
                        {
                            oNewMember.Mpayroll = "20211031";
                        }
                        else if (x == 13)
                        {
                            oNewMember.Mpayroll = "20211130";
                        }
                        else if (x == 14)
                        {
                            oNewMember .Mpayroll  = "20211231";
                        }

                        if (xlRange.Cells[i, x] != null && xlRange.Cells[i, x].Value2 != null && xlRange.Cells[i, x].ToString() != "" && xlRange.Cells[i, x].Value2.ToString() != "")
                        {
                            try
                            {
                                oNewMember.RegFee = Convert.ToDouble(xlRange.Cells[i, x].Value2.ToString());
                                
                            }
                            catch {; }
                        }
                        Application.DoEvents();
                        //if (oNewMember.RegFee != 0)
                        //{

                            oNewMember.Memberid = oNewMember.MigrateExpense(ref error);
                        //}
                    }






                }
            }
        }
    }
}
