using System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop;
using ExcelApp = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Collections;
using ReadExcel.Classes;
using System.Runtime.InteropServices;

namespace ReadExcel
{
    public partial class frmEquityTransactions2018 : Form
    {
        string filename = "";
        Classes.Transaction otrans = new Classes.Transaction();
        Classes.Transaction onewtrans = null;
        public frmEquityTransactions2018()
        {
            InitializeComponent();
        }

        private void btnMigrate_Click(object sender, EventArgs e)
        {
            btnMigrate.Enabled = false;
            MigrateUniqueEquityTrans();
            btnMigrate.Enabled = true;
        }
        private void MigrateUniqueEquityTrans()
        {
            Application.DoEvents();
            string error = "";
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            for (int s = 4; s <= 86; s+=2)
            {
                ExcelApp.Worksheet excelWorksheets = excelWorkbook.Sheets[s];
                ExcelApp.Range excelRange = excelWorksheets.UsedRange;
                for (int i = 2; i <= excelRange.Rows.Count; i++)
                {
                    if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                    {
                        onewtrans = new Classes.Transaction();
                        label2.Text = i.ToString() + " Of " + excelRange.Rows.Count;
                        double val = Convert.ToDouble(excelRange.Cells[i, 1].Value2.ToString());
                        DateTime date = DateTime.FromOADate(val);
                        onewtrans.TransDate = date;
                        double val2 = Convert.ToDouble(excelRange.Cells[i, 2].Value2.ToString());
                        DateTime date2 = DateTime.FromOADate(val2);
                        onewtrans.ValueDate = date2;
                        if (excelRange.Cells[i, 4] != null && excelRange.Cells[i, 4].Value2 != null)
                        {
                            onewtrans.Debit = Convert.ToDouble(excelRange.Cells[i, 4].Value2.ToString());
                        }

                        if (excelRange.Cells[i, 5] != null && excelRange.Cells[i, 5].Value2 != null)
                        {
                            onewtrans.Credit = Convert.ToDouble(excelRange.Cells[i, 5].Value2.ToString());
                        }
                        if (excelRange.Cells[i, 6] != null && excelRange.Cells[i, 6].Value2 != null)
                        {
                            onewtrans.Balance = Convert.ToDouble(excelRange.Cells[i, 6].Value2.ToString());
                        }
                        if (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                        {
                            onewtrans.Particulars = excelRange.Cells[i, 3].Value2.ToString();

                        }
                        onewtrans.TransID = onewtrans.AddEdditUniqueEquityTrans(ref error);
                    }




                }
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDailog1 = new OpenFileDialog();
            openFileDailog1.ShowDialog();
            txtFileName.Text = openFileDailog1.FileName;
            filename = openFileDailog1.FileName;
        }
    }
}
