using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcepApp = Microsoft.Office.Interop.Excel;

namespace ReadExcel
{
    public partial class frmMigrateMembersandshares : Form
    {
        public frmMigrateMembersandshares()
        {
            InitializeComponent();
        }
        Classes.Member omember = new Classes.Member();
        Classes.Member onewmember = null;
        Classes.Transaction oTransaction = new Classes.Transaction();
        Classes.Transaction onewTransactions = null;
        string FileName = "";

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            txtFileName.Text = openFileDialog .FileName;
            FileName = openFileDialog .FileName;
            
        }

        private void btnReadFile_Click(object sender, EventArgs e)
        {
            Application.DoEvents();
            int memberid = 0;
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook   xlworkbook = excepApp.Workbooks.Open (FileName);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for(int i=4;i<=xlRange.Rows.Count;i++)
            {
                onewTransactions = new Classes.Transaction();
                double openingbalance = 0;
                label1.Text = i.ToString() + "Of " + xlRange.Rows.Count;
                
                if (xlRange.Cells[i, 3] != null && xlRange.Cells[i, 3].Value2 != null)
                {
                    try
                    {

                    
                    onewTransactions.Membername = xlRange.Cells[i, 3].Value2.ToString();
                    }catch {; }
                    }

                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null && xlRange.Cells[i, 2].ToString() !="" && xlRange.Cells[i, 2].Value2.ToString() != "")
                {
                    try
                    {
                        onewTransactions.StaffNumber = xlRange.Cells[i, 2].Value2.ToString();
                    }
                    catch {; }
                }
                Application.DoEvents();
                onewTransactions.MemberID = onewTransactions.AddEdditBarabaramembers(ref error);
                memberid = onewTransactions.MemberID;
                for(int x=1; x<=xlRange.Columns.Count;x++)
                {
                    label2.Text = x.ToString() + " Of " + xlRange.Columns.Count;
                    if(x==4)
                    {
                   
                    onewTransactions = new Classes.Transaction();
                    onewTransactions.MemberID = memberid;
                        if (xlRange.Cells[i, x] != null && xlRange.Cells[i, x].Value2 != null)
                        {
                            try
                            {
                                onewTransactions.openingbalance = Convert.ToDouble(xlRange.Cells[i, x].Value2.ToString());
                            }
                            catch {; }
                        }
                            
                        openingbalance = onewTransactions.openingbalance;
                        //double val = Convert.ToDouble(xlRange .Cells[3, x].Value2.ToString());
                        //DateTime date = DateTime.FromOADate(val);
                        string date = "20181231";
                        onewTransactions.openingbalance = openingbalance;
                      
                        onewTransactions.TransDate1  = date;

                        onewTransactions.LoanAmount = "0";
                        if(memberid >0)
                        {
                            onewTransactions.TransID = onewTransactions.AddEdditBarabaraShareTransactions(ref error);

                        }
                    
                    }
                    else if (x > 4 && x<19)
                    {
                        string date = "";
                        DateTime tDate = DateTime.Now;
                        if (x == 5) date = "2019/01/01";
                        if (x == 6) date = "20190201";
                        if (x == 7) date = "20190301";
                        if (x == 8) date = "20190401";
                        if (x == 9) date = "20190501";
                        if (x == 10) date = "20190601";
                        if (x == 11) date = "20190701";
                        if (x == 12) date = "201908801";
                        if (x == 13) date = "20190901";
                        if (x == 14) date = "20191001";
                        if (x == 15) date = "20191101";
                        if (x == 16) date = "20191201";
                        if (x == 17) date = "20200101";
                        if (x == 18) date = "20200201";
                        DateTime.TryParse(date, out tDate);
                        onewTransactions = new Classes.Transaction();
                        onewTransactions.MemberID = memberid;
                        onewTransactions.openingbalance = 0 ;
                        //double val = Convert.ToDouble(date);
                        //DateTime date4 = DateTime.FromOADate(val);
                        onewTransactions.TransDate1      = date  ;
                        if(xlRange.Cells[i, x] !=null && xlRange.Cells[i, x].Value2 !=null)
                        {
                            try
                            {
                                onewTransactions.Amount = Convert.ToDouble(xlRange.Cells[i, x].Value2.ToString());
                            }
                            catch {; }
                        }
                        
                        onewTransactions.TransID = onewTransactions.AddEdditBarabaraShareTransactions(ref error);
                    }
                }
            }
        }

        private void btnMetro_Click(object sender, EventArgs e)
        {
            Link myLink = new Link();
            myLink.GetMetroResults();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            Application.DoEvents();
            int memberid = 0;
            string error = "";
            ExcepApp.Application excepApp = new ExcepApp.Application();
            ExcepApp.Workbook xlworkbook = excepApp.Workbooks.Open(FileName);
            ExcepApp.Worksheet xlworksheet = xlworkbook.Sheets[1];
            ExcepApp.Range xlRange = xlworksheet.UsedRange;
            for (int i = 7; i <= 164; i++)
            {
                onewTransactions = new Classes.Transaction();
                
                label1.Text = i.ToString() + "Of " + 164;
                Application.DoEvents();
                int employerid = 0;
                try
                {
                    if ((xlRange.Cells[i, 2].Value2 == "" || xlRange.Cells[i, 2].Value2 == null))
                    {

                        goto NextRow;
                    }
                }
                catch {; }
               
                if ((xlRange.Cells[i, 2] .Value2.ToString()== "PAYROLL" && xlRange.Cells[i, 2].Value2 != null))
                {
                    employerid = 1;
                    goto NextRow;
                }
                if ((xlRange.Cells[i, 2].Value2.ToString() == "TOTAL" && xlRange.Cells[i, 2].Value2 != null))
                {
                    
                    goto NextRow;
                }

                if ((xlRange.Cells[i, 2].Value2.ToString() == "DIASPORA-KENYA" && xlRange.Cells[i, 2].Value2 != null))
                {
                    employerid = 2;
                    goto NextRow;
                }
                if ((xlRange.Cells[i, 2].Value2.ToString() == " DIASPORA - TZ  Ex 20" && xlRange.Cells[i, 2].Value2 != null))
                {
                    employerid = 3;
                    goto NextRow;
                }
                if ((xlRange.Cells[i, 2].Value2.ToString() == "DIASPORA-UG Ex 35" && xlRange.Cells[i, 2].Value2 != null))
                {
                    employerid = 4;
                    goto NextRow;
                }
                if ((xlRange.Cells[i, 2].Value2.ToString() == "DIASPORA-RW Ex 8" && xlRange.Cells[i, 2].Value2 != null))
                {
                    employerid = 5;
                    goto NextRow;
                }
                if ((xlRange.Cells[i, 2].Value2.ToString() == "FROZEN" && xlRange.Cells[i, 2].Value2 != null))
                {
                    employerid = 6;
                    goto NextRow;
                }

                if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                {
                    try
                    {


                        onewTransactions.Membername = xlRange.Cells[i, 2].Value2.ToString();
                    }
                    catch {; }
                }

               
               
                if (i > 6 && i < 92) employerid = 1;
                if (i > 94 && i < 106) employerid = 2;
                if (i > 108 && i < 117) employerid = 3;
                if (i > 119 && i < 134) employerid = 4;
                if (i > 136 && i < 138) employerid = 5;
                if (i > 140 && i < 165) employerid = 6;
                onewTransactions.EmployerId = employerid;
                onewTransactions.MemberID = onewTransactions.AddEdditLongHornmembers(ref error);
                
                memberid = onewTransactions.MemberID;
                for (int x = 3; x <= 5; x++)
                {
                    label2.Text = x.ToString() + " Of " + xlRange.Columns.Count;


                    int sharetypeid = 0;
                    if(x==3)
                    {
                        sharetypeid = 1;//for shares 
                    }
                    else if (x == 4)
                    {
                        sharetypeid = 2;//for shares 
                    }
                    else
                    {
                        sharetypeid = 3;
                    }
                        onewTransactions = new Classes.Transaction();
                        onewTransactions.MemberID = memberid;
                    onewTransactions.ShareTypeId = sharetypeid;                     
                        
                        if (xlRange.Cells[i, x] != null && xlRange.Cells[i, x].Value2 != null)
                        {
                            try
                            {
                                onewTransactions.Amount = Convert.ToDouble(xlRange.Cells[i, x].Value2.ToString());
                            }
                            catch {; }
                        }

                        onewTransactions.TransID = onewTransactions.AddEdditLongHornShareTransactions(ref error);

                    }

                for (int x = 6; x <= 9; x++)
                {
                    int Loantypeid = 0;
                    label2.Text = x.ToString() + " Of " + xlRange.Columns.Count;


                    if (x == 6)
                    {
                        Loantypeid = 1;//for Loans
                    }
                    else if (x == 7)
                    {
                        Loantypeid = 2;//for Loans 
                    }
                    else if (x == 8)
                    {
                        Loantypeid = 3;//for Loans 
                    }
                    else
                    {
                        Loantypeid = 4;
                    }
                    onewTransactions = new Classes.Transaction();
                    onewTransactions.MemberID = memberid;
                    onewTransactions.LoanTypeId = Loantypeid;

                    if (xlRange.Cells[i, x] != null && xlRange.Cells[i, x].Value2 != null)
                    {
                        try
                        {
                            onewTransactions.Amount = Convert.ToDouble(xlRange.Cells[i, x].Value2.ToString());
                        }
                        catch {; }
                    }

                    onewTransactions.LoanID = onewTransactions.AddEdditLongHornLoans(ref error);

                }
                NextRow:;
                }
            button1.Enabled = true;
            }

        private void frmMigrateMembersandshares_Load(object sender, EventArgs e)
        {
            btnMetro.Enabled = false;
            btnReadFile.Enabled = false;
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
    }
    }

