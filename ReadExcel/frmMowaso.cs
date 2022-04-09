using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace ReadExcel
{
    public partial class frmMowaso : Form
    {
        public frmMowaso()
        {
            InitializeComponent();
        }
        string filename = "";
        Classes.Loan oloan = new Classes.Loan();
        Classes.Loan onewLoan = null;
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDailog1 = new OpenFileDialog();
            openFileDailog1.ShowDialog();
            txtFileName.Text = openFileDailog1.FileName;
            filename = openFileDailog1.FileName;
        }

        private void btnMigrate_Click(object sender, EventArgs e)
        {
            btnMigrate.Enabled = false;
            MigrateToMowasco();
            btnMigrate.Enabled = true;
        }
        private void MigrateToMowasco()
        {
            Application.DoEvents();
            string error = "";
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            ExcelApp.Worksheet excelWorksheets = excelWorkbook.Sheets[1];
            ExcelApp.Range excelRange = excelWorksheets.UsedRange;
            for (int i = 2; i <= excelRange.Rows.Count; i++)
            {
                onewLoan = new Classes.Loan();
                label2.Text = i.ToString() + " Of " + excelRange.Rows.Count;
                onewLoan = new Classes.Loan();
                label2.Text = i.ToString() + " Of " + excelRange.Rows.Count;
                onewLoan.MemberNo = (excelRange.Cells[i, 1].Value2.ToString());
                onewLoan.OriginalAmount = Convert.ToDouble(excelRange.Cells[i, 6].Value2.ToString());
                onewLoan.Loanamount = Convert.ToDouble(excelRange.Cells[i, 6].Value2.ToString());
                //onewLoan.SumOtherCharges = Convert.ToDouble(excelRange.Cells[i, 6].Value2.ToString());
                //onewLoan.Interest = Convert.ToDouble(excelRange.Cells[i, 3].Value2.ToString());
                //string dat = excelRange.Cells[i, 10].Value2.ToString();
                ////DateTime loandate = DateTime.Now;
                ////DateTime.TryParse(dat, out loandate);
                //double val = Convert.ToDouble(excelRange.Cells[i, 10].Value2.ToString());
                //DateTime date = DateTime.FromOADate(val);

                onewLoan.Loantransdate = excelRange.Cells[i, 4].Value2.ToString();
                onewLoan.MemberName = (excelRange.Cells[i, 2].Value2.ToString());
                onewLoan.IDNo = (excelRange.Cells[i, 3].Value2.ToString());


                if (excelRange.Cells[i, 8] != null && excelRange.Cells[i, 8].Value2 != null)
                {
                    onewLoan.RemainingPeriod = int.Parse(excelRange.Cells[i, 8].Value2.ToString());

                }
                onewLoan.LoanNo = 1;

                onewLoan.LoanId = onewLoan.AddEditLoan(ref error);

            }
        }

        private void MigrateToMowascoSheet2()
        {
            Application.DoEvents();
            string error = "";
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            ExcelApp.Worksheet excelWorksheets = excelWorkbook.Sheets[2];
            ExcelApp.Range excelRange = excelWorksheets.UsedRange;
            for (int i = 2; i <= excelRange.Rows.Count; i++)
            {
                onewLoan = new Classes.Loan();
                label2.Text = i.ToString() + " Of " + excelRange.Rows.Count;
                onewLoan = new Classes.Loan();
                label2.Text = i.ToString() + " Of " + excelRange.Rows.Count;
                onewLoan.MemberNo = (excelRange.Cells[i, 1].Value2.ToString());
                onewLoan.OriginalAmount = Convert.ToDouble(excelRange.Cells[i, 6].Value2.ToString());
                onewLoan.Loanamount = Convert.ToDouble(excelRange.Cells[i, 6].Value2.ToString());
                //onewLoan.SumOtherCharges = Convert.ToDouble(excelRange.Cells[i, 6].Value2.ToString());
                //onewLoan.Interest = Convert.ToDouble(excelRange.Cells[i, 3].Value2.ToString());
                string dat = excelRange.Cells[i, 10].Value2.ToString();
                //DateTime loandate = DateTime.Now;
                //DateTime.TryParse(dat, out loandate);
                double val = Convert.ToDouble(excelRange.Cells[i, 10].Value2.ToString());
                DateTime date = DateTime.FromOADate(val);

                onewLoan.Loantransdate = excelRange.Cells[i, 4].Value2.ToString();
                onewLoan.MemberName = (excelRange.Cells[i, 2].Value2.ToString());
                onewLoan.IDNo = (excelRange.Cells[i, 3].Value2.ToString());


                if (excelRange.Cells[i, 8] != null && excelRange.Cells[i, 8].Value2 != null)
                {
                    onewLoan.RemainingPeriod = int.Parse(excelRange.Cells[i, 8].Value2.ToString());

                }
                onewLoan.LoanNo = 2;

                onewLoan.LoanId = onewLoan.AddEditLoan(ref error);

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

            button1.Enabled = false;
            MigrateSharesToMowasco();
            button1.Enabled = true;
        }
        private void MigrateSharesToMowasco()
        {
            Application.DoEvents();
            string error = "";
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            ExcelApp.Worksheet excelWorksheets = excelWorkbook.Sheets[1];
            ExcelApp.Range excelRange = excelWorksheets.UsedRange;
            for (int i = 2; i <= excelRange.Rows.Count; i++)
            {
                onewLoan = new Classes.Loan();
                label2.Text = i.ToString() + " Of " + excelRange.Rows.Count;
                if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                {
                    onewLoan.MemberNo = (excelRange.Cells[i, 1].Value2.ToString());
                }
                if (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                {
                    onewLoan.IDNo = (excelRange.Cells[i, 3].Value2.ToString());
                }
                if (excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null)
                {
                    onewLoan.MemberName = (excelRange.Cells[i, 2].Value2.ToString());
                }

                if (excelRange.Cells[i, 5] != null && excelRange.Cells[i, 5].Value2 != null)
                {
                    onewLoan.Loanamount = Convert.ToDouble(excelRange.Cells[i, 5].Value2.ToString());
                }
                if (excelRange.Cells[i, 4] != null && excelRange.Cells[i, 4].Value2 != null)
                {
                    //double val = Convert.ToDouble(excelRange.Cells[i, 1].Value2.ToString());
                    //DateTime date = DateTime.FromOADate(val);

                    onewLoan.Loantransdate = excelRange.Cells[i, 4].Value2.ToString(); 
                }

                if (onewLoan.MemberNo != null)
                {
                    onewLoan.TransId = onewLoan.AddEditShareTransactions(ref error);
                }


            }
        }

        private void frmMowaso_Load(object sender, EventArgs e)
        {

        }
        private void MigrateSharesToMowascoSheet2()
        {
            Application.DoEvents();
            string error = "";
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            ExcelApp.Worksheet excelWorksheets = excelWorkbook.Sheets[2];
            ExcelApp.Range excelRange = excelWorksheets.UsedRange;
            for (int i = 2; i <= excelRange.Rows.Count; i++)
            {
                onewLoan = new Classes.Loan();
                label2.Text = i.ToString() + " Of " + excelRange.Rows.Count;
                if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                {
                    onewLoan.MemberNo = (excelRange.Cells[i, 1].Value2.ToString());
                }
                if (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                {
                    onewLoan.IDNo = (excelRange.Cells[i, 3].Value2.ToString());
                }
                if (excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null)
                {
                    onewLoan.MemberName = (excelRange.Cells[i, 2].Value2.ToString());
                }

                if (excelRange.Cells[i, 5] != null && excelRange.Cells[i, 5].Value2 != null)
                {
                    onewLoan.Loanamount = Convert.ToDouble(excelRange.Cells[i, 5].Value2.ToString());
                }
                if (excelRange.Cells[i, 4] != null && excelRange.Cells[i, 4].Value2 != null)
                {
                    //double val = Convert.ToDouble(excelRange.Cells[i, 1].Value2.ToString());
                    //DateTime date = DateTime.FromOADate(val);

                    onewLoan.Loantransdate = excelRange.Cells[i, 4].Value2.ToString();
                }

                if (onewLoan.MemberNo != null)
                {
                    onewLoan.TransId = onewLoan.AddEditShareTransactions(ref error);
                }


            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            MigrateSharesToMowascoSheet2();
            button2.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            MigrateToMowascoSheet2();
            button3.Enabled = true;
        }
    }
}
