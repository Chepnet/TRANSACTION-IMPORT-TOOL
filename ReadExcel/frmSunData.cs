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
    public partial class frmSunData : Form
    {
        public frmSunData()
        {
            InitializeComponent();
        }
        string filename = "";
        Classes.Member onewmember = null;
        Classes.Kin onewkin = null;
        Classes.MemberRegistration onewmemberregfee = null;
        Classes.Transaction onewtransaction = null;
        Classes.Bens onewben = null;
        Classes.GroupOfficial ogroupofficial = new GroupOfficial();
        Classes.GroupOfficial onewofficial = null;
     
        string error = "";
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDailog1 = new OpenFileDialog();
            openFileDailog1.ShowDialog();
            txtfile.Text = openFileDailog1.FileName;
            filename = openFileDailog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled=false;
            readExcel();
            button2.Enabled = true;
        }
        private void readExcel()
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            ExcelApp.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            ExcelApp.Range excelRange = excelWorksheet.UsedRange;
            string data = string.Empty;
            int rowcount = 83;
            int colcount = 37;// excelRange.Columns.Count;//7; // to be checked later



            string errm = "";

            Application.DoEvents();


            if (excelWorksheet != null)
            {

                for (int i = 71; i <= rowcount; i++)
                {
                    onewmember = new Classes.Member();
                    lblread.Text = i.ToString() + " Of " + rowcount.ToString();
                    try {
                        onewmember.Mfirstname = (excelRange.Cells[i, 2].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewmember.Msurname = (excelRange.Cells[i, 3].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewmember.Mcell = (excelRange.Cells[i, 4].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewmember.Mtel1 = (excelRange.Cells[i, 1].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewmember.Mgender = (excelRange.Cells[i, 8].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewmember.IDNO = (excelRange.Cells[i, 6].Value2.ToString());
                    }
                    catch {; }
                    try {
                        if (excelRange.Cells[i, 21] != null && excelRange.Cells[i, 21].Value2 != null)
                        {
                            double val = Convert.ToDouble(excelRange.Cells[i, 21].Value2.ToString());
                            DateTime date = DateTime.FromOADate(val);

                            onewmember.DatePaid = date;
                        }
                    }
                    catch {; }
                    try {
                        onewmember.Mdate = Convert.ToDateTime(excelRange.Cells[i, 9].Value2.ToString());
                    }
                    catch {; }
                    //if (excelRange.Cells[i, 9] != null && excelRange.Cells[i, 9].Value2 != null)
                    //{
                    //    double val = Convert.ToDouble(excelRange.Cells[i, 9].Value2.ToString());
                    //    DateTime dateofbirth = DateTime.FromOADate(val);

                    //    onewmember.Mdate = excelRange.Cells[i, 9].Value2.ToString();
                    //}
                    try {
                        double regfe = 0;
                        Double.TryParse(excelRange.Cells[i, 20].Value2.ToString(), out regfe);
                        onewmember.RegFee = regfe;
                    }
                    catch {; }



                    onewmember.Memberid = onewmember.AddEditMember(ref error);

                    onewkin = new Classes.Kin();
                    onewkin.Memberid = onewmember.Memberid;
                    try {
                        onewkin.Kinname = (excelRange.Cells[i, 10].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewkin.Town = (excelRange.Cells[i, 15].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewkin.Relationship = (excelRange.Cells[i, 2].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewkin.Telephone = (excelRange.Cells[i, 12].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewkin.Kincode = (excelRange.Cells[i, 16].Value2.ToString());
                    } catch {; }
                    try {
                        onewkin.Kinaddress = (excelRange.Cells[i, 14].Value2.ToString());
                    }
                    catch {; }
                    try {
                        onewkin.CreatedBy = (excelRange.Cells[i, 11].Value2.ToString());
                    }
                    catch {; }
                    onewkin.Kinid = onewkin.AddEditKin(ref error);


                    for (int x = 22; x <= 37; x += 2)
                    {

                        if (excelRange.Cells[i, x] != null && excelRange.Cells[i, x].Value2 != null)
                        {
                            try
                            {
                                onewtransaction = new Classes.Transaction();
                                onewtransaction.MemberID = onewmember.Memberid;
                                double paid = 0;
                                Double.TryParse(excelRange.Cells[i, x].Value2.ToString(), out paid);
                                try {
                                    onewtransaction.Amount = paid;
                                }
                                catch {; }
                                try {
                                    if (excelRange.Cells[i, x + 1] != null && excelRange.Cells[i, x + 1].Value2 != null)
                                    {
                                        double val = Convert.ToDouble(excelRange.Cells[i, x + 1].Value2.ToString());
                                        DateTime date = DateTime.FromOADate(val);

                                        onewtransaction.TransDate = date;
                                    }
                                }
                                catch {; }
                                onewtransaction.TransID = onewtransaction.AddEdditLongHornShareTransactions(ref error);
                            }
                            catch (Exception ex)
                            {
                                string err = ex.Message.ToString();


                            }





                        }



                    }




                }



            }




            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(excelWorksheet);

            //close and release
            // excelWorksheet.Close();
            Marshal.ReleaseComObject(excelWorksheet);

            //quit and release
            // ExcelApp.Quit();
            // Marshal.ReleaseComObject(ExcelApp);

        }




        private void readExcelMigrateOfficials()
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            ExcelApp.Worksheet excelWorksheet = excelWorkbook.Sheets[3];
            ExcelApp.Range excelRange = excelWorksheet.UsedRange;
            string data = string.Empty;
            int rowcount = excelRange.Rows.Count;
            int colcount = excelRange.Columns.Count;//7; // to be checked later



            string errm = "";

            Application.DoEvents();


            if (excelWorksheet != null)
            {

                for (int i = 3; i <= rowcount; i++)
                {
                    onewofficial = new GroupOfficial();
                    onewofficial.Employername = excelRange.Cells[i, i].Value2.ToString();

                    lblread.Text = i.ToString() + " Of " + rowcount.ToString();

                    for (int x = 1; x <= colcount; x++)
                    {

                        data = "";
                        if (excelRange.Cells[i, x] != null && excelRange.Cells[i, x].Value2 != null)
                        {
                            try
                            {
                                data = excelRange.Cells[i, x].Value2.ToString();
                            }
                            catch (Exception ex)
                            {
                                string err = ex.Message.ToString();


                            }





                        }
                        //   onewben.DateOfBirth = (excelRange.Cells[i, 3].Value2.ToString());

                        onewben.Benid = onewben.AddEditBen(ref error);





                    }



                }




                GC.Collect();
                GC.WaitForPendingFinalizers();


                Marshal.ReleaseComObject(excelRange);
                Marshal.ReleaseComObject(excelWorksheet);


                Marshal.ReleaseComObject(excelWorksheet);








            } }
    
        private void readExcelBeneficiary()
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            ExcelApp.Worksheet excelWorksheet = excelWorkbook.Sheets[2];
            ExcelApp.Range excelRange = excelWorksheet.UsedRange;
            string data = string.Empty;
            int rowcount = 237;// excelRange.Rows.Count;
            int colcount = 37;// excelRange.Columns.Count;//7; // to be checked later



            string errm = "";

            Application.DoEvents();


            if (excelWorksheet != null)
            {

                for (int i = 3; i<= rowcount; i++)
                {
                    onewben = new Classes.Bens();
                    //.Text = i.ToString() + " Of " + rowcount.ToString();
                    onewben .Benname = (excelRange.Cells[i, 1].Value2.ToString());
                    onewben.Bencode  = (excelRange.Cells[i, 4].Value2.ToString());
                    if (excelRange.Cells[i, 3] != null && excelRange.Cells[i, 3].Value2 != null)
                    {
                        try { 
                        //double val = Convert.ToDouble(excelRange.Cells[i, 3].Value2.ToString());
                            DateTime date = DateTime.Now;// DateTime.FromOADate(val);

                          DateTime.TryParse(excelRange.Cells[i, 3].Value2.ToString(), out date);
                            onewben.DateOfBirth = date;
                        }
                        catch {; }
                    }
                    //   onewben.DateOfBirth = (excelRange.Cells[i, 3].Value2.ToString());
                    onewben.Relationship = (excelRange.Cells[i, 8].Value2.ToString());
                    onewben.Telephone  = (excelRange.Cells[i, 5].Value2.ToString());
                    onewben.Town = (excelRange.Cells[i, 6].Value2.ToString());
                    onewben.CreatedBy  = (excelRange.Cells[i, 2].Value2.ToString());
                   onewben.Benid = onewben.AddEditBen(ref error);





                }



            }




            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(excelWorksheet);

            //close and release
            // excelWorksheet.Close();
            Marshal.ReleaseComObject(excelWorksheet);

            //quit and release
            // ExcelApp.Quit();
            // Marshal.ReleaseComObject(ExcelApp);







        }

        private void frmSunData_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            readExcelBeneficiary();
        }

        private void btnMigrateOfficials_Click(object sender, EventArgs e)
        {

        }
    }
    }

