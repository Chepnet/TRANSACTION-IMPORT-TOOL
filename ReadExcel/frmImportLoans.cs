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
    public partial class frmImportLoans : Form
    {
        public frmImportLoans()
        {
            InitializeComponent();
        }
        string filename = "";
        Classes.Transaction onewtrans = new Transaction();
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDailog1 = new OpenFileDialog();
            openFileDailog1.ShowDialog();
            txtFilenName.Text = openFileDailog1.FileName;
            filename = openFileDailog1.FileName;
        }

        private void btnReadAndImportData_Click(object sender, EventArgs e)
        {
            btnReadAndImportData.Enabled = false;
            MigrateUniqueEquityTrans();
            btnReadAndImportData.Enabled = true;
        }
        private void MigrateUniqueEquityTrans()
        {
            Application.DoEvents();
            string error = "";
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            //for (int s = 1; s <= 1; s += 2)
            //{
                ExcelApp.Worksheet excelWorksheets = excelWorkbook.Sheets[1];
                ExcelApp.Range excelRange = excelWorksheets.UsedRange;
                for (int i = 3; i <= excelRange.Rows.Count; i++)
                {
                    
                        
                        label1 .Text = i.ToString() + " Of " + excelRange.Rows.Count;
                    string date = "";


                    for (int j = 10; j <= excelRange.Columns .Count; j++)
                    {
                        onewtrans = new Classes.Transaction();
                        if(j==10)
                        {
                            date  ="20210731";
                        }
                        if (j == 11)
                        {
                            date = "20210831";
                        }
                        if (j == 12)
                        {
                            date = "20210930";
                        }
                        if (j == 13)
                        {
                            date = "20211031";
                        }
                        if (j == 14)
                        {
                            date = "20211130";
                        }
                        onewtrans.TransDate1  = date;
                      
                            if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                        {
                            onewtrans.MemberNumber  = excelRange.Cells[i, 1].Value2.ToString();
                        }
                        if (excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null)
                        {
                            onewtrans.StaffNumber   = excelRange.Cells[i, 2].Value2.ToString();
                        }
                        if (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                        {
                            onewtrans.Membername  = excelRange.Cells[i, 3].Value2.ToString();
                        }
                        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        {
                            onewtrans.Amount  = Convert.ToDouble(excelRange.Cells[i, j].Value2.ToString());
                        }
                        onewtrans.TransID = onewtrans.AddEdditBarabaradirectdeposit (ref error);

                    }





                //}
            }
        }
        private void MigrateBARABARATrans()
        {
            Application.DoEvents();
            string error = "";
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            
                ExcelApp.Worksheet excelWorksheets = excelWorkbook.Sheets[1];
                ExcelApp.Range excelRange = excelWorksheets.UsedRange;
                for (int i = 4; i <= excelRange.Rows.Count; i++)
                {
                    if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                    {
                        onewtrans = new Classes.Transaction();
                        label1.Text = i.ToString() + " Of " + excelRange.Rows.Count;
                        //double val = Convert.ToDouble(excelRange.Cells[i, 1].Value2.ToString());
                        //DateTime date = DateTime.FromOADate(val);
                        //onewtrans.TransDate = date;
                        //double val2 = Convert.ToDouble(excelRange.Cells[i, 2].Value2.ToString());
                        //DateTime date2 = DateTime.FromOADate(val2);
                        //onewtrans.ValueDate = date2;
                        if (excelRange.Cells[i, 1] != null && excelRange.Cells[i, 1].Value2 != null)
                        {
                            onewtrans.MemberNumber  = excelRange.Cells[i, 1].Value2.ToString();
                        }

                        if (excelRange.Cells[i, 2] != null && excelRange.Cells[i, 2].Value2 != null)
                        {
                            onewtrans.Membername  = excelRange.Cells[i, 2].Value2.ToString();
                        }
                        if (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                        {
                            onewtrans.Amount  = Convert.ToDouble(excelRange.Cells[i, 3].Value2.ToString());
                        }
                        if (excelRange.Cells[i, 4] != null && excelRange.Cells[i, 4].Value2 != null)
                        {
                            onewtrans.LoanTypeName  = excelRange.Cells[i, 4].Value2.ToString();

                        }
                    if (excelRange.Cells[i, 5] != null && excelRange.Cells[i, 5].Value2 != null)
                    {
                        onewtrans.Credit = Convert.ToDouble(excelRange.Cells[i, 5].Value2.ToString());

                    }
                    if (excelRange.Cells[i, 6] != null && excelRange.Cells[i, 6].Value2 != null)
                    {
                        onewtrans.ShareTypeId = int.Parse (excelRange .Cells[i, 6].Value2.ToString());

                    }
                    onewtrans.TransID = onewtrans.AddEdditbarabaraTrans (ref error);
                    }




                
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            MigrateBARABARATrans();
            button1.Enabled = true;
        }
    }
}
