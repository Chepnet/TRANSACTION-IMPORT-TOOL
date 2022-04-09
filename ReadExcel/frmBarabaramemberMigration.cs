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
    public partial class frmBarabaramemberMigration : Form
    {
        public frmBarabaramemberMigration()
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

                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                {
                    try
                    {
                        oNewMember.Mpayroll = xlRange.Cells[i, 2].Value2.ToString();
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
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4].ToString() != "" && xlRange.Cells[i, 4].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Msurname = xlRange.Cells[i, 4].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 5] != null && xlRange.Cells[i, 5].Value2 != null && xlRange.Cells[i, 5].ToString() != "" && xlRange.Cells[i, 5].Value2.ToString() != "")
                {
                    try
                    {
                        oNewMember.Mothername = xlRange.Cells[i,5].Value2.ToString();
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
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[2];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 2; i <= 2; i++)
            {
                oNewShareTransations  = new Classes.ShareTransactions();

                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    try
                    {
                        oNewShareTransations.Mcode = xlRange.Cells[i, 1].Value2.ToString();
                    }
                    catch {; }
                }

                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2].ToString() != "" && xlRange.Cells[i, 2].Value2.ToString() != "")
                {
                    try
                    {
                        oNewShareTransations.Mpayroll = xlRange.Cells[i, 2].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3].ToString() != "" && xlRange.Cells[i, 3].Value2.ToString() != "")
                {
                    try
                    {
                        oNewShareTransations.Description = xlRange.Cells[i, 3].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4].ToString() != "" && xlRange.Cells[i, 4].Value2.ToString() != "")
                {
                    try
                    {
                        oNewShareTransations.Amount = Convert.ToDouble(xlRange.Cells[i, 4].Value2.ToString());
                    }
                    catch {; }
                }
               
                Application.DoEvents();
                oNewShareTransations.TransId = oNewShareTransations.AddEditBarabaraSharesandDeposit(ref error);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(FileName);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 2; i <=2; i++)
            {
                oNewLoan  = new Classes.Loan();

                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    try
                    {
                        oNewLoan.MemberNo = xlRange.Cells[i, 1].Value2.ToString();
                    }
                    catch {; }
                }

                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2].ToString() != "" && xlRange.Cells[i, 2].Value2.ToString() != "")
                {
                    try
                    {
                        oNewLoan.IDNo = xlRange.Cells[i, 2].Value2.ToString();
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
                        oNewLoan.RefNo  = xlRange.Cells[i, 5].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 6] != null && xlRange.Cells[i, 6].Value2 != null && xlRange.Cells[i, 6].ToString() != "" && xlRange.Cells[i, 6].Value2.ToString() != "")
                {
                    try
                    {
                        oNewLoan.OriginalAmount = Convert.ToDouble(xlRange.Cells[i, 6].Value2.ToString());
                    }
                    catch {; }
                }
                Application.DoEvents();
                oNewLoan.LoanId = oNewLoan.AddEditBarabaraLoan(ref error);

            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void txtFileName_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblFileName_Click(object sender, EventArgs e)
        {

        }

        private void btnUpdaterate_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(FileName);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 2; i <= 397; i++)
            {
                oNewShareTransations = new Classes.ShareTransactions();

                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    try
                    {
                        oNewShareTransations.Mcode = xlRange.Cells[i, 1].Value2.ToString();
                    }
                    catch {; }
                }

                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2].ToString() != "" && xlRange.Cells[i, 2].Value2.ToString() != "")
                {
                    try
                    {
                        oNewShareTransations.Mpayroll = xlRange.Cells[i, 2].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null && xlRange.Cells[i, 3].ToString() != "" && xlRange.Cells[i, 3].Value2.ToString() != "")
                {
                    try
                    {
                        oNewShareTransations.Description = xlRange.Cells[i, 3].Value2.ToString();
                    }
                    catch {; }
                }
                if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null && xlRange.Cells[i, 4].ToString() != "" && xlRange.Cells[i, 4].Value2.ToString() != "")
                {
                    try
                    {
                        oNewShareTransations.Amount = Convert.ToDouble(xlRange.Cells[i, 4].Value2.ToString());
                    }
                    catch {; }
                }

                Application.DoEvents();
                oNewShareTransations.TransId = oNewShareTransations.AddEditBarabaraSharesandDeposit(ref error);

            }
        }
    }
    }

