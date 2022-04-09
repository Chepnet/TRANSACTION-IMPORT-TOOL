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
    public partial class frmReadExcel : Form
    {
        payments payment = null;
        Transaction Member = null;
        Transaction onewLoan = null;
        Transaction onewrepayment = null;
        CopyLoan ocopyLoan = new CopyLoan();
        CopyLoan onewcopyloan = null;
        public frmReadExcel()
        {
            InitializeComponent();
        }

        string filename = "";

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDailog1 = new OpenFileDialog();
            openFileDailog1.ShowDialog();
            txtfile.Text = openFileDailog1.FileName;
            filename = openFileDailog1.FileName;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            
            readExcelMembersLoansRepayment();

            
            button2.Enabled = true;
        }
        private void readExcel()
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            int count = excelWorkbook.Sheets.Count;
            for (int sheetnum = 1; sheetnum <= count; sheetnum++)
            {
                ExcelApp.Worksheet excelWorksheet = excelWorkbook.Sheets[sheetnum];

                ExcelApp.Range excelRange = excelWorksheet.UsedRange;

                decimal dvalue = numericUpDown1.Value;

                int[] Cols = { 3, 5, 6 }; //Columns to loop
                                          //C, E, F
                string data = string.Empty;
                int rowcount = excelRange.Rows.Count;
                int colcount = excelRange.Columns.Count;
                int monthcol = 0, amountcol = 0, interestcol = 0, principalcol = 0, balancecol = 0;
                int monthrow = 0;
                int monthcol1 = 0, amountcol1 = 0, interestcol1 = 0, principalcol1 = 0, balancecol1 = 0;
                int monthrow1 = 0;

                string strdata = "", strName = "", strStaffNo = "", strMemberNo = "";
                bool foundtheLoan = false;
                bool foundtheLoan1 = false;
                string errm = "";
                Member memb = null;
                string mcode = "";
                Transaction trans = null;
                Transaction trans1 = null;
                //try
                //{
                //foreach (ExcelApp.Worksheet sheet in excelWorkbook.Worksheets)
                //ExcelApp.Worksheet sheet = excelWorkbook.Worksheets[cvalue];


                string test = "";
                try
                {
                    test = (string)excelRange.Cells[24, 12].Value2.ToString();

                }
                catch
                {
                    ;
                }
                string loantype = "";
                string Longloantype = "";
                if (excelWorksheet != null)
                {
                    for (int i = 1; i <= rowcount; i++)
                    {
                        //foundtheLoan = false;
                        for (int x = 1; x <= colcount; x++)
                        {

                            if (excelRange.Cells[i, x] != null && excelRange.Cells[i, x].Value2 != null)
                            {

                                data = excelRange.Cells[i, x].Value2.ToString();

                                if (data == "NAME")
                                {
                                    try
                                    {
                                        if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strName = excelRange.Cells[i, x + 1].Value2;
                                        }

                                        if (strName == "")
                                            strName = excelRange.Cells[i, x + 2].Value2;
                                        trans.Membername = strName;
                                        goto NextRow;
                                    }
                                    catch {; }

                                }
                                if (data == "STAFF NO.")
                                {
                                    try
                                    {
                                        strStaffNo = excelRange.Cells[i, x + 1].Value2.ToString();
                                        if (strStaffNo == "")
                                            strStaffNo = excelRange.Cells[i, x + 2].Value2.ToString();
                                        trans.StaffNumber = strStaffNo;

                                        goto NextRow;
                                    }
                                    catch { goto NextRow; ; }
                                }
                                if (data == "SACCO NO.")
                                {
                                    try
                                    {
                                        if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strMemberNo = excelRange.Cells[i, x + 1].Value2.ToString();
                                        }
                                        else if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strMemberNo = excelRange.Cells[i, x + 2].Value2.ToString();
                                        }
                                        trans.MemberNumber = strMemberNo;

                                    }
                                    catch {; }

                                    goto NextRow;
                                }



                                strdata = data.ToLower().Trim();
                                if (strdata == "month")
                                {
                                    int a = 4;
                                }
                                switch (strdata)
                                {
                                    case "month":
                                        if (monthcol == 0)
                                        {
                                            trans = new Transaction();
                                            trans.SheetNo = sheetnum;
                                            trans.Membername = strName;
                                            trans.MemberNumber = strMemberNo;
                                            trans.StaffNumber = strStaffNo;
                                            trans.strAmount = "";
                                            trans.strPrincipal = "";
                                            trans.strInterest = "";
                                            trans.strBalance = "";

                                            monthcol = x; //in case we are only reading one loan
                                            amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                            if (excelRange.Cells[i + 1, balancecol].Value2 != null)
                                            {
                                                trans.LoanAmount = excelRange.Cells[i + 1, balancecol].Value2.ToString();
                                            }

                                            trans.strBalance = trans.LoanAmount;
                                            trans.LoanTypeName = loantype;
                                            foundtheLoan = true;
                                            monthrow = i;

                                            goto NextRow; //i++;

                                            //we may have to initiate a new loan/transaction here
                                        }
                                        else
                                        if (monthcol != 0 && monthcol == x)
                                        {
                                            trans = new Transaction();
                                            trans.SheetNo = sheetnum;
                                            trans.Membername = strName;
                                            trans.MemberNumber = strMemberNo;
                                            trans.StaffNumber = strStaffNo;
                                            trans.strAmount = "";
                                            trans.strPrincipal = "";
                                            trans.strInterest = "";
                                            trans.strBalance = "";

                                            monthcol = x; //in case we are only reading one loan
                                            amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                            if (excelRange.Cells[i + 1, balancecol].Value2 != null)
                                            {
                                                trans.LoanAmount = excelRange.Cells[i + 1, balancecol].Value2.ToString();
                                                trans.strBalance = trans.LoanAmount;
                                                if (loantype == "")
                                                {
                                                    trans.LoanTypeName = "Long Term Loan";
                                                }
                                                else
                                                {
                                                    trans.LoanTypeName = loantype;
                                                }

                                                foundtheLoan = true;
                                                monthrow = i;

                                            }


                                            goto NextRow; //i++;
                                        }
                                        else if (monthcol < x)
                                        {
                                            trans1 = new Transaction();
                                            trans1.SheetNo = sheetnum;
                                            trans1.Membername = strName;
                                            trans1.MemberNumber = strMemberNo;
                                            trans1.StaffNumber = strStaffNo;
                                            trans1.strAmount = "";
                                            trans1.strPrincipal = "";
                                            trans1.strInterest = "";
                                            trans1.strBalance = "";

                                            monthcol1 = x; //in case we are only reading one loan
                                            amountcol1 = x + 1; interestcol1 = x + 2; principalcol1 = x + 3; balancecol1 = x + 4;
                                            trans1.LoanAmount = excelRange.Cells[i + 1, balancecol1].Value2.ToString();
                                            trans1.strBalance = trans.LoanAmount;
                                            if (loantype == "")
                                            {
                                                trans1.LoanTypeName = "Short Term";
                                            }
                                            else
                                            {
                                                trans1.LoanTypeName = loantype;
                                            }

                                            foundtheLoan1 = true;
                                            monthrow = i;


                                            goto NextRow; //i++;
                                        }
                                        break;
                                    case "top up":
                                        // if (monthcol == 0)

                                        trans = new Transaction();
                                        trans.SheetNo = sheetnum;
                                        trans.Membername = strName;
                                        trans.MemberNumber = strMemberNo;
                                        trans.StaffNumber = strStaffNo;

                                        trans.strAmount = "";
                                        trans.strPrincipal = "";
                                        trans.strInterest = "";
                                        trans.strBalance = "";

                                        monthcol = x; //in case we are only reading one loan
                                        amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                        if (excelRange.Cells[i, balancecol].Value2 == null)
                                            trans.LoanAmount = "0";
                                        else
                                        {
                                            trans.LoanAmount = excelRange.Cells[i, balancecol].Value2.ToString();
                                            trans.strBalance = trans.LoanAmount;
                                        }

                                        try
                                        {
                                            //trans.MonthName 
                                            double val = double.Parse(excelRange.Cells[i - 1, x].Value2.ToString());
                                            trans.TransDate = DateTime.FromOADate(val);
                                            x = x;
                                        }
                                        catch { }
                                        trans.LoanTypeName = "Top up of " + trans.LoanAmount;
                                        foundtheLoan = true;
                                        monthrow = i;
                                        trans.TransID = trans.topupTrans(ref errm);

                                        goto NextRow; //i++;

                                    //we may have to initiate a new loan/transaction here

                                    //break;
                                    case "top-up":
                                        // if (monthcol == 0)
                                        {
                                            trans = new Transaction();
                                            trans.SheetNo = sheetnum;
                                            trans.Membername = strName;
                                            trans.MemberNumber = strMemberNo;
                                            trans.StaffNumber = strStaffNo;

                                            trans.strAmount = "";
                                            trans.strPrincipal = "";
                                            trans.strInterest = "";
                                            trans.strBalance = "";

                                            monthcol = x; //in case we are only reading one loan
                                            amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                            //trans.LoanAmount = excelRange.Cells[i, balancecol].Value2 != null ? excelRange.Cells[i, balancecol].Value2.ToString() : null;

                                            double myval = 0;
                                            for (int n = balancecol + 1; n >= 1; n--)
                                            {
                                                trans.LoanAmount = "";
                                                try
                                                {
                                                    trans.LoanAmount = (string)excelRange.Cells[i, n].Value2.ToString();
                                                    trans.strBalance = trans.LoanAmount;
                                                }
                                                catch
                                                {
                                                    ;
                                                }
                                                double.TryParse(trans.LoanAmount, out myval);
                                                if (myval > 0)
                                                    break;
                                            }

                                            if (trans.LoanAmount == null)
                                            {
                                                trans.LoanAmount = "0";
                                                trans.strBalance = "0";
                                            }


                                            try
                                            {
                                                //trans.MonthName 
                                                double val = double.Parse(excelRange.Cells[i - 1, x].Value2.ToString());
                                                trans.TransDate = DateTime.FromOADate(val);
                                                x = x;
                                            }
                                            catch { }
                                            trans.LoanTypeName = "Top up of " + trans.LoanAmount;
                                            foundtheLoan = true;
                                            monthrow = i;
                                            trans.TransID = trans.topupTrans(ref errm);
                                            i = i + 2;
                                            // goto NextRow; //i++;

                                            //we may have to initiate a new loan/transaction here
                                        }
                                        break;


                                        //case "loan
                                }


                                if (foundtheLoan)
                                {
                                    if (x == monthcol)
                                    {
                                        try
                                        {
                                            //trans.MonthName 
                                            double val = double.Parse(excelRange.Cells[i, x].Value2.ToString());
                                            trans.TransDate = DateTime.FromOADate(val);
                                            x = x;
                                        }
                                        catch { }
                                    }
                                    if (x == amountcol)
                                    {
                                        trans.strAmount = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //trans.strAmount = excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strAmount == null)
                                        //    trans.strAmount = "0";
                                    }
                                    if (x == interestcol)
                                    {
                                        trans.strInterest = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strInterest == null)
                                        //    trans.strInterest = "0";
                                    }
                                    if (x == principalcol)
                                    {
                                        trans.strPrincipal = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strPrincipal == null)
                                        //    trans.strPrincipal = "0";
                                    }
                                    if (x == balancecol)
                                    {
                                        trans.strBalance = excelRange.Cells[i, x].Value2.ToString();
                                        trans.TransID = trans.AddEdditTrans(ref errm);
                                        //save since we now have all records then go to nextRow
                                        //goto NextRow;
                                    }
                                }
                                bool skip = false;
                                if (strdata.Contains("loan"))
                                {

                                    if (strdata.Contains("statement"))
                                        skip = false;
                                    else
                                    {
                                        if (monthcol1 == 0)
                                        {
                                            trans1 = new Transaction();
                                            trans1.LoanTypeName = data.ToString();
                                            loantype = trans1.LoanTypeName;
                                            skip = true;

                                        }
                                        else
                                        {
                                            trans1.LoanTypeName = data.ToString();
                                            loantype = trans1.LoanTypeName;
                                            skip = true;

                                        }

                                    }

                                }
                                if (foundtheLoan1)
                                {
                                    if (x == monthcol1)
                                    {
                                        try
                                        {
                                            //trans.MonthName 
                                            double val = double.Parse(excelRange.Cells[i, x].Value2.ToString());
                                            trans1.TransDate = DateTime.FromOADate(val);
                                            x = x;
                                        }
                                        catch { }
                                    }
                                    if (x == amountcol1)
                                    {
                                        trans1.strAmount = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //trans.strAmount = excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strAmount == null)
                                        //    trans.strAmount = "0";
                                    }
                                    if (x == interestcol1)
                                    {
                                        trans1.strInterest = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strInterest == null)
                                        //    trans.strInterest = "0";
                                    }
                                    if (x == principalcol1)
                                    {
                                        trans1.strPrincipal = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strPrincipal == null)
                                        //    trans.strPrincipal = "0";
                                    }
                                    if (x == balancecol1)
                                    {
                                        trans1.strBalance = excelRange.Cells[i, x].Value2.ToString();
                                        trans1.TransID = trans1.AddEdditTrans(ref errm);
                                        //save since we now have all records then go to nextRow
                                        //goto NextRow;
                                    }
                                }
                            }

                        }

                        //this is where we go to the next row

                        NextRow:;
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
                //}
                //catch (Exception ex)
                //{
                //    string abc = ex.Message.ToString();
                //    string a = "";
                //    ;
                //}            




            }
        }
        private void readExcelIncludingMembers()
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            int count = excelWorkbook.Sheets.Count;
            int memberId = 0;

            for (int sheetnum = 1; sheetnum <= count; sheetnum++)
            {
                string loantype = "";
                ExcelApp.Worksheet excelWorksheet = excelWorkbook.Sheets[sheetnum];

                ExcelApp.Range excelRange = excelWorksheet.UsedRange;

                decimal dvalue = numericUpDown1.Value;

                int[] Cols = { 3, 5, 6 }; //Columns to loop
                                          //C, E, F
                string data = string.Empty;
                int rowcount = excelRange.Rows.Count;
                int colcount = excelRange.Columns.Count;//7; // to be checked later
                int monthcol = 0, amountcol = 0, interestcol = 0, principalcol = 0, balancecol = 0;
                int monthrow = 0;
                int monthcol1 = 0, amountcol1 = 0, interestcol1 = 0, principalcol1 = 0, balancecol1 = 0;
                int monthrow1 = 0;

                string strdata = "", strName = "", strStaffNo = "", strMemberNo = "";
                bool foundtheLoan = false;
                bool foundtheLoan1 = false;
                string errm = "";
                Member memb = null;
                string mcode = "";
                Transaction trans = null;
                Transaction trans1 = null;
                trans = new Transaction();
                trans1 = new Transaction();

                Member = null;
                onewLoan = null;

                if (excelWorksheet != null)
                {
                    string PreviousLoanAmount = "";
                    for (int i = 1; i <= rowcount; i++)
                    {
                        if (i == 53)
                        {
                            i = i;
                        }
                        //foundtheLoan = false;
                        for (int x = 1; x <= colcount; x++)
                        {

                            if (excelRange.Cells[i, x] != null && excelRange.Cells[i, x].Value2 != null)
                            {

                                data = excelRange.Cells[i, x].Value2.ToString();

                                if (data == "NAME")
                                {
                                    try
                                    {
                                        if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strName = excelRange.Cells[i, x + 1].Value2;
                                        }

                                        if (strName == "")
                                            strName = excelRange.Cells[i, x + 2].Value2;
                                        trans.Membername = strName;
                                        goto NextRow;
                                    }
                                    catch {; }

                                }
                                if (data == "STAFF NO.")
                                {
                                    try
                                    {
                                        strStaffNo = excelRange.Cells[i, x + 1].Value2.ToString();
                                        if (strStaffNo == "")
                                            strStaffNo = excelRange.Cells[i, x + 2].Value2.ToString();
                                        trans.StaffNumber = strStaffNo;

                                        goto NextRow;
                                    }
                                    catch { goto NextRow; ; }
                                }
                                if (data == "SACCO NO.")
                                {
                                    try
                                    {
                                        if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strMemberNo = excelRange.Cells[i, x + 1].Value2.ToString();
                                        }
                                        else if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strMemberNo = excelRange.Cells[i, x + 2].Value2.ToString();
                                        }
                                        trans.MemberNumber = strMemberNo;

                                    }
                                    catch {; }

                                    if (Member == null)
                                    {
                                        Member = new Transaction();

                                        Member.Membername = strName;
                                        Member.MemberNumber = strMemberNo;
                                        Member.StaffNumber = strStaffNo;
                                        Member.MemberID = Member.AddEdditBarabaramembers(ref errm);
                                        memberId = Member.MemberID;
                                    }

                                    goto NextRow;
                                }



                                strdata = data.ToLower().Trim();
                                if (strdata == "month")
                                {
                                    int a = 4;
                                }
                                switch (strdata)
                                {
                                    case "month":
                                        //if (monthcol == 0)
                                        //{
                                        trans = new Transaction();
                                        trans1 = new Transaction();
                                        trans.SheetNo = sheetnum;
                                        trans.Membername = strName;
                                        trans.MemberNumber = strMemberNo;
                                        trans.StaffNumber = strStaffNo;
                                        trans.strAmount = "";
                                        trans.strPrincipal = "";
                                        trans.strInterest = "";
                                        trans.strBalance = "";

                                        monthcol = x; //in case we are only reading one loan
                                        amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                        if (excelRange.Cells[i + 1, balancecol].Value2 != null)
                                        {
                                            trans.LoanAmount = excelRange.Cells[i + 1, balancecol].Value2.ToString();
                                        }

                                        trans.strBalance = trans.LoanAmount;
                                        trans.LoanTypeName = loantype;
                                        foundtheLoan = true;
                                        monthrow = i;

                                        onewLoan = new Transaction();

                                        onewLoan.MemberID = memberId;// Member.MemberID;

                                        onewLoan.LoanAmount = trans.LoanAmount;

                                        onewLoan.LoanID = onewLoan.AddEdditBarabaraLoans(ref errm);


                                        goto NextRow; //i++;

                                        //we may have to initiate a new loan/transaction here
                                        //}
                                        //else
                                        if (monthcol != 0 /*&& monthcol == x*/)
                                        {
                                            trans = new Transaction();
                                            trans.SheetNo = sheetnum;
                                            trans.Membername = strName;
                                            trans.MemberNumber = strMemberNo;
                                            trans.StaffNumber = strStaffNo;
                                            trans.strAmount = "";
                                            trans.strPrincipal = "";
                                            trans.strInterest = "";
                                            trans.strBalance = "";

                                            monthcol = x; //in case we are only reading one loan
                                            amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                            if (excelRange.Cells[i + 1, balancecol].Value2 != null)
                                            {
                                                trans.LoanAmount = excelRange.Cells[i + 1, balancecol].Value2.ToString();
                                                trans.strBalance = trans.LoanAmount;
                                                if (loantype == "")
                                                {
                                                    trans.LoanTypeName = "Long Term Loan";
                                                }
                                                else
                                                {
                                                    trans.LoanTypeName = loantype;
                                                }

                                                foundtheLoan = true;
                                                monthrow = i;

                                            }
                                            //  PreviousLoanAmount = trans.LoanAmount;

                                            goto NextRow; //i++;
                                        }
                                        else if (monthcol < x)
                                        {
                                            trans1 = new Transaction();
                                            trans1.SheetNo = sheetnum;
                                            trans1.Membername = strName;
                                            trans1.MemberNumber = strMemberNo;
                                            trans1.StaffNumber = strStaffNo;
                                            trans1.strAmount = "";
                                            trans1.strPrincipal = "";
                                            trans1.strInterest = "";
                                            trans1.strBalance = "";

                                            monthcol1 = x; //in case we are only reading one loan
                                            amountcol1 = x + 1; interestcol1 = x + 2; principalcol1 = x + 3; balancecol1 = x + 4;
                                            trans1.LoanAmount = excelRange.Cells[i + 1, balancecol1].Value2.ToString();
                                            trans1.strBalance = trans1.LoanAmount;

                                            if (loantype == "")
                                            {
                                                trans1.LoanTypeName = "Short Term";
                                            }
                                            else
                                            {
                                                trans1.LoanTypeName = loantype;
                                            }

                                            foundtheLoan1 = true;
                                            monthrow = i;
                                            //if (PreviousLoanAmount != trans1.LoanAmount)
                                            //{
                                            //onewLoan = null;
                                            //if (onewLoan == null)
                                            //{
                                            //    onewLoan = new Classes.Transaction();
                                            //    if (Member != null)
                                            //    {
                                            //        onewLoan.MemberID = Member.MemberID;
                                            //    }

                                            //    onewLoan.LoanAmount = trans1.LoanAmount;
                                            //    onewLoan.LoanID = onewLoan.AddEdditBarabaraLoans(ref errm);
                                            //    PreviousLoanAmount = trans1.LoanAmount;
                                            //}

                                            //}

                                            goto NextRow; //i++;

                                        }
                                        break;
                                    case "top up":
                                        // if (monthcol == 0)

                                        trans = new Transaction();
                                        trans.SheetNo = sheetnum;
                                        trans.Membername = strName;
                                        trans.MemberNumber = strMemberNo;
                                        trans.StaffNumber = strStaffNo;

                                        trans.strAmount = "";
                                        trans.strPrincipal = "";
                                        trans.strInterest = "";
                                        trans.strBalance = "";

                                        monthcol = x; //in case we are only reading one loan
                                        amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                        if (excelRange.Cells[i, balancecol].Value2 == null)
                                            trans.LoanAmount = "0";
                                        else
                                        {
                                            trans.LoanAmount = excelRange.Cells[i, balancecol].Value2.ToString();
                                            trans.strBalance = trans.LoanAmount;
                                        }

                                        try
                                        {
                                            //trans.MonthName 
                                            double val = double.Parse(excelRange.Cells[i - 1, x].Value2.ToString());
                                            trans.TransDate = DateTime.FromOADate(val);
                                            x = x;
                                        }
                                        catch { }
                                        trans.LoanTypeName = "Top up of " + trans.LoanAmount;
                                        foundtheLoan = true;
                                        monthrow = i;
                                        trans.TransID = trans.topupTrans(ref errm);


                                        goto NextRow; //i++;

                                    case "top-up":
                                        // if (monthcol == 0)
                                        {
                                            trans = new Transaction();
                                            trans.SheetNo = sheetnum;
                                            trans.Membername = strName;
                                            trans.MemberNumber = strMemberNo;
                                            trans.StaffNumber = strStaffNo;

                                            trans.strAmount = "";
                                            trans.strPrincipal = "";
                                            trans.strInterest = "";
                                            trans.strBalance = "";

                                            monthcol = x; //in case we are only reading one loan
                                            amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                            //trans.LoanAmount = excelRange.Cells[i, balancecol].Value2 != null ? excelRange.Cells[i, balancecol].Value2.ToString() : null;

                                            double myval = 0;
                                            for (int n = balancecol + 1; n >= 1; n--)
                                            {
                                                trans.LoanAmount = "";
                                                try
                                                {
                                                    trans.LoanAmount = (string)excelRange.Cells[i, n].Value2.ToString();
                                                    trans.strBalance = trans.LoanAmount;
                                                    PreviousLoanAmount = trans.LoanAmount;
                                                }
                                                catch
                                                {
                                                    ;
                                                }
                                                double.TryParse(trans.LoanAmount, out myval);
                                                if (myval > 0)
                                                    break;
                                            }

                                            if (trans.LoanAmount == null)
                                            {
                                                trans.LoanAmount = "0";
                                                trans.strBalance = "0";
                                            }


                                            try
                                            {
                                                //trans.MonthName 
                                                double val = double.Parse(excelRange.Cells[i - 1, x].Value2.ToString());
                                                trans.TransDate = DateTime.FromOADate(val);
                                                x = x;
                                            }
                                            catch { }
                                            trans.LoanTypeName = "Top up of " + trans.LoanAmount;
                                            foundtheLoan = true;
                                            monthrow = i;
                                            trans.TransID = trans.topupTrans(ref errm);
                                            i = i + 2;
                                            // goto NextRow; //i++;

                                            //we may have to initiate a new loan/transaction here
                                        }
                                        break;


                                        //case "loan
                                }


                                if (foundtheLoan)
                                {
                                    onewrepayment = new Transaction();
                                    // onewrepayment = new Classes.Transaction();
                                    if (x == monthcol)
                                    {
                                        try
                                        {
                                            //trans.MonthName 
                                            double val = double.Parse(excelRange.Cells[i, x].Value2.ToString());
                                            trans.TransDate = DateTime.FromOADate(val);
                                            onewrepayment.TransDate = DateTime.FromOADate(val);
                                            x = x;
                                        }
                                        catch { }
                                    }
                                    if (x == amountcol)
                                    {
                                        trans.strAmount = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //trans.strAmount = excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strAmount == null)
                                        //    trans.strAmount = "0";

                                    }
                                    if (x == interestcol)
                                    {
                                        trans.strInterest = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strInterest == null)
                                        //    trans.strInterest = "0";

                                    }
                                    if (x == principalcol)
                                    {
                                        trans.strPrincipal = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strPrincipal == null)
                                        //    trans.strPrincipal = "0";

                                    }
                                    if (x == balancecol)
                                    {
                                        trans.strBalance = excelRange.Cells[i, x].Value2.ToString();

                                        trans.TransID = trans.AddEdditTrans(ref errm);
                                        //save since we now have all records then go to nextRow
                                        //goto NextRow;
                                    }


                                }
                                bool skip = false;
                                if (strdata.Contains("loan"))
                                {

                                    if (strdata.Contains("statement"))
                                        skip = false;
                                    else
                                    {
                                        if (monthcol1 == 0)
                                        {
                                            trans1 = new Transaction();
                                            trans1.LoanTypeName = data.ToString();
                                            loantype = trans1.LoanTypeName;
                                            skip = true;

                                        }
                                        else
                                        {
                                            trans1.LoanTypeName = data.ToString();
                                            loantype = trans1.LoanTypeName;
                                            skip = true;

                                        }

                                    }

                                }
                                if (foundtheLoan1)
                                {
                                    onewrepayment = new Transaction();
                                    if (x == monthcol1)
                                    {
                                        try
                                        {
                                            //trans.MonthName 
                                            double val = double.Parse(excelRange.Cells[i, x].Value2.ToString());
                                            trans1.TransDate = DateTime.FromOADate(val);
                                            onewrepayment.TransDate = trans1.TransDate;
                                            x = x;
                                        }
                                        catch { }
                                    }
                                    if (x == amountcol1)
                                    {
                                        trans1.strAmount = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //trans.strAmount = excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strAmount == null)
                                        //    trans.strAmount = "0";
                                        onewrepayment.TransDate = trans1.TransDate;
                                    }
                                    if (x == interestcol1)
                                    {
                                        trans1.strInterest = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strInterest == null)
                                        //    trans.strInterest = "0";
                                        onewrepayment.strInterest = trans1.strInterest;
                                    }
                                    if (x == principalcol1)
                                    {
                                        trans1.strPrincipal = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        //if (trans.strPrincipal == null)
                                        //    trans.strPrincipal = "0";
                                        onewrepayment.strPrincipal = trans1.strPrincipal;
                                    }
                                    if (x == balancecol1)
                                    {
                                        trans1.strBalance = excelRange.Cells[i, x].Value2.ToString();
                                        trans1.TransID = trans1.AddEdditTrans(ref errm);
                                        //save since we now have all records then go to nextRow
                                        //goto NextRow;
                                        onewrepayment.strBalance = trans1.strBalance;
                                    }
                                    onewrepayment.TransDate = trans1.TransDate;
                                    onewrepayment.MemberID = Member.MemberID;
                                    onewrepayment.LoanID = onewLoan.LoanID;
                                    onewrepayment.strInterest = trans1.strInterest;
                                    onewrepayment.strPrincipal = trans1.strPrincipal;
                                    onewrepayment.LoanAmount = trans1.LoanAmount;
                                    onewrepayment.strBalance = trans1.strBalance;
                                    onewrepayment.strAmount = trans1.strAmount;
                                    onewrepayment.LoanRepaymentId = onewrepayment.AddEdditBarabaraLoanrepayments(ref errm);


                                }
                            }



                            ///PreviousLoanAmount = trans.LoanAmount;
                        }
                        //trans = new Transaction();
                        //this is where we go to the next row
                        if (onewLoan != null)
                        {
                            onewrepayment = new Transaction();
                            onewrepayment.TransDate = trans.TransDate;
                            onewrepayment.MemberID = memberId;// Member.MemberID;
                            onewrepayment.LoanID = onewLoan.LoanID;
                            onewrepayment.strInterest = trans.strInterest;
                            onewrepayment.strPrincipal = trans.strPrincipal;
                            onewrepayment.LoanAmount = trans.LoanAmount;
                            onewrepayment.strBalance = trans.strBalance;
                            onewrepayment.strAmount = trans.strAmount;
                            onewrepayment.LoanRepaymentId = onewrepayment.AddEdditBarabaraLoanrepayments(ref errm);
                        }


                        NextRow:;



                    }
                    Member = null;
                    //if (Member == null)
                    //{
                    //    Member = new Classes.Transaction();

                    //    Member.Membername = strName;
                    //    Member.MemberNumber = strMemberNo;
                    //    Member.StaffNumber = strStaffNo;
                    //    Member.MemberID = Member.AddEdditBarabaramembers(ref errm);
                    //}



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
                //}
                //catch (Exception ex)
                //{
                //    string abc = ex.Message.ToString();
                //    string a = "";
                //    ;
                //}            




            }
        }
        private void readExcelMembersLoansRepayments()
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            int count = excelWorkbook.Sheets.Count;//1;
            int memberId = 0;

            for (int sheetnum = 1; sheetnum <= count; sheetnum++)
            {
                string loantype = "";
                ExcelApp.Worksheet excelWorksheet = excelWorkbook.Sheets[sheetnum];

                ExcelApp.Range excelRange = excelWorksheet.UsedRange;

                decimal dvalue = numericUpDown1.Value;

                int[] Cols = { 3, 5, 6 }; //Columns to loop
                                          //C, E, F
                string data = string.Empty;
                int rowcount = excelRange.Rows.Count;
                int colcount = excelRange.Columns.Count;//7; // to be checked later
                int monthcol = 0, amountcol = 0, interestcol = 0, principalcol = 0, balancecol = 0;
                int monthrow = 0;
             
                string strdata = "", strName = "", strStaffNo = "", strMemberNo = "";
                bool foundtheLoan = false;
                string errm = "";
                              
                Transaction trans = null;
                               
                if (excelWorksheet != null)
                {
                    
                    for (int i = 1; i <= rowcount; i++)
                    {
                        string  amountpaid = "";
                        for (int x = 1; x <= colcount; x++)
                        {

                            if (excelRange.Cells[i, x] != null && excelRange.Cells[i, x].Value2 != null)
                            {

                                data = excelRange.Cells[i, x].Value2.ToString();

                                if (data == "NAME")
                                {
                                    try
                                    {
                                        if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strName = excelRange.Cells[i, x + 1].Value2;
                                        }

                                        if (strName == "")
                                            strName = excelRange.Cells[i, x + 2].Value2;
                                       // trans.Membername = strName;
                                        goto NextRow;
                                    }
                                    catch {; }

                                }
                                if (data == "STAFF NO.")
                                {
                                    try
                                    {
                                        strStaffNo = excelRange.Cells[i, x + 1].Value2.ToString();
                                        if (strStaffNo == "")
                                            strStaffNo = excelRange.Cells[i, x + 2].Value2.ToString();
                                       // trans.StaffNumber = strStaffNo;

                                        goto NextRow;
                                    }
                                    catch { goto NextRow; ; }
                                }
                                if (data == "SACCO NO.")
                                {
                                    try
                                    {
                                        if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strMemberNo = excelRange.Cells[i, x + 1].Value2.ToString();
                                        }
                                        else if (excelRange.Cells[i, x + 1].Value2.ToString() != "")
                                        {
                                            strMemberNo = excelRange.Cells[i, x + 2].Value2.ToString();
                                        }
                                       // trans.MemberNumber = strMemberNo;

                                    }
                                    catch {; }

                                    if (Member == null)
                                    {
                                        Member = new Transaction();

                                        Member.Membername = strName;
                                        Member.MemberNumber = strMemberNo;
                                        Member.StaffNumber = strStaffNo;
                                        Member.MemberID = Member.AddEdditBarabaramembers(ref errm);
                                        memberId = Member.MemberID;
                                    }

                                    goto NextRow;
                                }

                                

                                strdata = data.ToLower().Trim();
                                if (strdata == "month")
                                {
                                    int a = 4;
                                }
                                switch (strdata)
                                {
                                    case "month":
                                        //if (monthcol == 0)
                                        //{
                                        trans = new Transaction();
                                        trans.SheetNo = sheetnum;
                                        trans.Membername = strName;
                                        trans.MemberNumber = strMemberNo;
                                        trans.StaffNumber = strStaffNo;
                                        trans.strAmount = "";
                                        trans.strPrincipal = "";
                                        trans.strInterest = "";
                                        trans.strBalance = "";

                                        monthcol = x; //in case we are only reading one loan
                                        amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                        if (excelRange.Cells[i + 1, balancecol].Value2 != null)
                                        {
                                            trans.LoanAmount = excelRange.Cells[i + 1, balancecol].Value2.ToString();
                                        }

                                        trans.strBalance = trans.LoanAmount;
                                        amountpaid = trans.LoanAmount;
                                        trans.LoanTypeName = loantype;
                                        if (trans.LoanTypeName == "")
                                        {
                                            trans.LoanTypeName = "Normal Loans";
                                        }
                                        foundtheLoan = true;
                                        monthrow = i;

                                        onewLoan = new Transaction();

                                        onewLoan.MemberID = memberId;// Member.MemberID;

                                        onewLoan.LoanAmount = trans.LoanAmount;

                                        onewLoan.LoanID = onewLoan.AddEdditBarabaraLoans(ref errm);
                                        // trans1 = new Transaction();


                                        goto NextRow; //i++;


                                        if (monthcol != 0 /*&& monthcol == x*/)
                                        {
                                            trans = new Transaction();
                                            trans.SheetNo = sheetnum;
                                            trans.Membername = strName;
                                            trans.MemberNumber = strMemberNo;
                                            trans.StaffNumber = strStaffNo;
                                            trans.strAmount = "";
                                            trans.strPrincipal = "";
                                            trans.strInterest = "";
                                            trans.strBalance = "";

                                            monthcol = x; //in case we are only reading one loan
                                            amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                            if (excelRange.Cells[i + 1, balancecol].Value2 != null)
                                            {
                                                trans.LoanAmount = excelRange.Cells[i + 1, balancecol].Value2.ToString();
                                                trans.strBalance = trans.LoanAmount;
                                                foundtheLoan = true;
                                                monthrow = i;

                                            }

                                            goto NextRow; //i++;
                                        }

                                        break;
                                    case "top up":
                                        // if (monthcol == 0)

                                        trans = new Transaction();
                                        trans.SheetNo = sheetnum;
                                        trans.Membername = strName;
                                        trans.MemberNumber = strMemberNo;
                                        trans.StaffNumber = strStaffNo;

                                        trans.strAmount = "";
                                        trans.strPrincipal = "";
                                        trans.strInterest = "";
                                        trans.strBalance = "";

                                        monthcol = x; //in case we are only reading one loan
                                        amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                        if (excelRange.Cells[i, balancecol].Value2 == null)
                                            trans.LoanAmount = "0";
                                        else
                                        {
                                            trans.LoanAmount = excelRange.Cells[i, balancecol].Value2.ToString();
                                            trans.strBalance = trans.LoanAmount;
                                            amountpaid = "-" + (trans.LoanAmount);
                                        }

                                        try
                                        {
                                            //trans.MonthName 
                                            double val = double.Parse(excelRange.Cells[i - 1, x].Value2.ToString());
                                            trans.TransDate = DateTime.FromOADate(val);
                                            x = x;
                                        }
                                        catch { }
                                        trans.LoanTypeName = "Top up of " + trans.LoanAmount;
                                        foundtheLoan = true;
                                        monthrow = i;
                                        trans.TransID = trans.topupTrans(ref errm);
                               

                                        goto NextRow; //i++;
                                

                                    case "top-up":
                                        // if (monthcol == 0)
                                        {
                                            trans = new Transaction();
                                            trans.SheetNo = sheetnum;
                                            trans.Membername = strName;
                                            trans.MemberNumber = strMemberNo;
                                            trans.StaffNumber = strStaffNo;

                                            trans.strAmount = "";
                                            trans.strPrincipal = "";
                                            trans.strInterest = "";
                                            trans.strBalance = "";

                                            monthcol = x; //in case we are only reading one loan
                                            amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                            //trans.LoanAmount = excelRange.Cells[i, balancecol].Value2 != null ? excelRange.Cells[i, balancecol].Value2.ToString() : null;

                                            double myval = 0;
                                            for (int n = balancecol + 1; n >= 1; n--)
                                            {
                                                trans.LoanAmount = "";
                                                try
                                                {
                                                    trans.LoanAmount = (string)excelRange.Cells[i, n].Value2.ToString();
                                                    trans.strBalance = trans.LoanAmount;
                                                    amountpaid = "-" + (trans.LoanAmount);
                                                }
                                                catch
                                                {
                                                    ;
                                                }
                                                double.TryParse(trans.LoanAmount, out myval);
                                                if (myval > 0)
                                                    break;
                                            }

                                            if (trans.LoanAmount == null)
                                            {
                                                trans.LoanAmount = "0";
                                                trans.strBalance = "0";
                                            }


                                            try
                                            {
                                                //trans.MonthName 
                                                double val = double.Parse(excelRange.Cells[i - 1, x].Value2.ToString());
                                                trans.TransDate = DateTime.FromOADate(val);
                                                x = x;
                                            }
                                            catch { }
                                            trans.LoanTypeName = "Top up of " + trans.LoanAmount;
                                            foundtheLoan = true;
                                            monthrow = i;
                                            trans.TransID = trans.topupTrans(ref errm);
                                            i = i + 2;
                                        }
                                        goto NextRow;


                                            case "REFINANCING":

                                        {
                                            trans = new Transaction();
                                            trans.SheetNo = sheetnum;
                                            trans.Membername = strName;
                                            trans.MemberNumber = strMemberNo;
                                            trans.StaffNumber = strStaffNo;

                                            trans.strAmount = "";
                                            trans.strPrincipal = "";
                                            trans.strInterest = "";
                                            trans.strBalance = "";

                                            monthcol = x; //in case we are only reading one loan
                                            amountcol = x + 1; interestcol = x + 2; principalcol = x + 3; balancecol = x + 4;
                                            //trans.LoanAmount = excelRange.Cells[i, balancecol].Value2 != null ? excelRange.Cells[i, balancecol].Value2.ToString() : null;

                                            double myval = 0;
                                            for (int n = balancecol + 1; n >= 1; n--)
                                            {
                                                trans.LoanAmount = "";
                                                try
                                                {
                                                    trans.LoanAmount = (string)excelRange.Cells[i, n].Value2.ToString();
                                                    trans.strBalance = trans.LoanAmount;
                                                    amountpaid = "-" + (trans.LoanAmount);
                                                }
                                                catch
                                                {
                                                    ;
                                                }
                                                double.TryParse(trans.LoanAmount, out myval);
                                                if (myval > 0)
                                                    break;
                                            }

                                            if (trans.LoanAmount == null)
                                            {
                                                trans.LoanAmount = "0";
                                                trans.strBalance = "0";
                                            }


                                            try
                                            {
                                                //trans.MonthName 
                                                double val = double.Parse(excelRange.Cells[i - 1, x].Value2.ToString());
                                                trans.TransDate = DateTime.FromOADate(val);
                                                x = x;
                                            }
                                            catch { }
                                            trans.LoanTypeName = "Refinancing of " + trans.LoanAmount;
                                            foundtheLoan = true;
                                            monthrow = i;
                                            trans.TransID = trans.topupTrans(ref errm);
                                            i = i + 2;

                                        }
                               
                                        break;
                                }




                                if (foundtheLoan)
                                {
                                   // onewrepayment = new Transaction();
                                  if (x == monthcol)
                                    {
                                        try
                                        {
                                            double val = double.Parse(excelRange.Cells[i, x].Value2.ToString());
                                            trans.TransDate = DateTime.FromOADate(val);
                                           // onewrepayment.TransDate  = DateTime.FromOADate(val);
                                            x = x;
                                        }
                                        catch { }
                                    }
                                    if (x == amountcol)
                                    {
                                        trans.strAmount = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                        amountpaid = trans.strAmount;                                        
                                    }
                                    if (x == interestcol)
                                    {
                                        trans.strInterest = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                                                              
                                    }
                                    if (x == principalcol)
                                    {
                                        trans.strPrincipal = excelRange.Cells[i, x].Value2.ToString() == null ? "" : excelRange.Cells[i, x].Value2.ToString();
                                                                                
                                    }
                                    if (x == balancecol)
                                    {
                                        trans.strBalance = excelRange.Cells[i, x].Value2.ToString();
                                      
                                        trans.TransID = trans.AddEdditTrans(ref errm);
                                       
                                    }


                                }
                                
                               
                            }
                        }

                        if (onewLoan!=null)
                        {
                        onewrepayment = new Transaction();

                            if(trans!=null)
                            {
                              
                        onewrepayment.TransDate = trans.TransDate;
                        onewrepayment.MemberID = memberId;// Member.MemberID;
                        onewrepayment.LoanID = onewLoan.LoanID;
                        onewrepayment.strInterest = trans.strInterest;
                        onewrepayment.strPrincipal = trans.strPrincipal;
                        onewrepayment.LoanAmount = trans.LoanAmount; ;
                        onewrepayment.strBalance = trans.strBalance;
                        onewrepayment.strAmount = amountpaid; 
                        onewrepayment.LoanRepaymentId = onewrepayment.AddEdditBarabaraLoanrepayments(ref errm);
                            }
                        }
                       
                        NextRow:;
                        
                    }
                    Member = null;
                    //if (Member == null)
                    //{
                    //    Member = new Classes.Transaction();

                    //    Member.Membername = strName;
                    //    Member.MemberNumber = strMemberNo;
                    //    Member.StaffNumber = strStaffNo;
                    //    Member.MemberID = Member.AddEdditBarabaramembers(ref errm);
                    //}
                   
                   

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
                //}
                //catch (Exception ex)
                //{
                //    string abc = ex.Message.ToString();
                //    string a = "";
                //    ;
                //}            

              
                    
                
            }
        }
        private void readExcelMembersLoansRepayment()
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
            ExcelApp.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            ExcelApp.Range excelRange = excelWorksheet.UsedRange;
            string data = string.Empty;
            int rowcount = excelRange.Rows.Count;
            int colcount = excelRange.Columns.Count;//7; // to be checked later
            int refnocol = 0, MobileNo = 0, fulldatecol = 0, principlecol = 0, creditamountcol = 0, Interestcol = 0, mpesafeecol = 0, duedatecol = 0;
            int empfirstnamecol = 0, trnaturecol = 0, empsurnamecol = 0, trtimecol = 0, paymentscol = 0, insurancecol = 0, companynamecol = 0;

            string strdata = "";
            bool foundtheLoan = false;
            string errm = "";

            Application.DoEvents();
            

            if (excelWorksheet != null)
            {
                
                for (int i = 1; i <= rowcount; i++)
                {
                    onewcopyloan = new Classes.CopyLoan();
                    lblread.Text = i.ToString() + " Of " + rowcount.ToString();
                    string amountpaid = "";
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
                                goto NextRow;

                            }





                        }
                    


                        strdata = data.ToLower().Trim();
                        DateTime fulldate = DateTime.Now;
                        DateTime duedate = DateTime.Now;
                        switch (strdata)
                        {
                           
                            case "refnumber":

                                 
                                
                                refnocol = x; //in case we are only reading one loan
                                refnocol = x; MobileNo = x + 3; fulldatecol = x +8 ; principlecol = x + 11; creditamountcol = x + 14;
                                Interestcol = x + 16; mpesafeecol = x + 18; duedatecol = x + 21;
                                empfirstnamecol = x + 24; trnaturecol = x + 27; empsurnamecol = x + 28; trtimecol = x + 30; paymentscol = x + 34; insurancecol = x + 37;
                                companynamecol = x + 39;
                                goto NextRow;
                        }
                        //i++;





                     
                        
                        if (refnocol != 0)
                        {
                            if (i == 1) goto NextRow;
                            if (x == refnocol)
                                    {
                                        try
                                        {
                                    //trans.MonthName 

                                    onewcopyloan.RefNo = (excelRange.Cells[i, x].Value2.ToString());
                                            x = x;
                                        }
                                        catch { }
                                    }
                             if (x == MobileNo )
                                    {
                                try
                                {
                                    onewcopyloan.MobileNo = (excelRange.Cells[i, x].Value2.ToString());
                                }
                                catch {; }
                                         
                               
                                    }
                            if (x == fulldatecol )
                            {
                                try
                                {
									 double val = (excelRange.Cells[i, x].Value2);
									DateTime dat=DateTime.FromOADate(val);
									 onewcopyloan.Fuldate = dat.ToString ();
									
									//DateTime.TryParse((excelRange.Cells[i, x].Value2), out fulldate);
									////double val = (excelRange.Cells[i , x].Value2.ToString());
									//onewcopyloan.Fuldate  = (excelRange.Cells[i, x].Value2);
									//double d = double.TryParse(excelRange.Cells[i, x].Value2, out d);
									//double fulldat = double.Parse(excelRange.Cells[i, x].Value2);
									//DateTime conv = DateTime.FromOADate(fulldat);
                                    //onewcopyloan.Fulldate = conv ;
                                }
                                catch (Exception ex) {ex.Message.ToString (); }
                              

                            }
                            if (x == principlecol )
                            {
                                try
                                {
                                    onewcopyloan.Principal = Double.Parse((excelRange.Cells[i, x].Value2.ToString()));
                                }
                                catch {; }
                                

                            }
                            if (x == creditamountcol )
                            {
                                try
                                {
                                    onewcopyloan.CreditAmount = Double.Parse((excelRange.Cells[i, x].Value2.ToString()));
                                }
                                catch {; }
                               

                            }
                            if (x == Interestcol )
                            {
                                try
                                {
                                    onewcopyloan.Interest = Double.Parse((excelRange.Cells[i, x].Value2.ToString()));
                                }
                                catch {; }


                            }
                            if (x == mpesafeecol )
                            {
                                try
                                {
                                    onewcopyloan.Mpesafee  = Double.Parse((excelRange.Cells[i, x].Value2.ToString()));
                                }
                                catch {; }


                            }
                            if (x == duedatecol )
                            {
                                try
                                {
									double val = (excelRange.Cells[i, x].Value2);
									DateTime dat = DateTime.FromOADate(val);
									onewcopyloan.Duedate  = dat.ToString();
									
                                }
                                catch (Exception ex) {ex.Message.ToString (); }


                            }
                            if (x == empfirstnamecol )
                            {
                                try
                                {
                                    onewcopyloan.Empfirstname  = (excelRange.Cells[i, x].Value2.ToString());
                                }
                                catch {; }


                            }
                            if (x == trnaturecol )
                            {
                                try
                                {
                                    onewcopyloan.Trnature  = (excelRange.Cells[i, x].Value2.ToString());
                                }
                                catch {; }


                            }
                            if (x == empsurnamecol )
                            {
                                try
                                {
                                    onewcopyloan.Empsurname  = (excelRange.Cells[i+1, x].Value2.ToString());
                                }
                                catch {; }


                            }
                            if (x == trtimecol )
                            {
                                try
                                {
                                    onewcopyloan.Trdetails = (excelRange.Cells[i+1, x].Value2.ToString());
                                }
                                catch {; }


                            }
                            if (x == paymentscol )
                            {
                                try
                                {
                                    onewcopyloan.Payments = (excelRange.Cells[i+1, x].Value2.ToString());
                                }
                                catch {; }


                            }
                            if (x == insurancecol )
                            {
                                try
                                {
                                    onewcopyloan.Insurance = (excelRange.Cells[i+1, x].Value2.ToString());
                                }
                                catch {; }


                            }
                            if (x == companynamecol )
                            {
                                try
                                {
                                    onewcopyloan.Schemename  = (excelRange.Cells[i+1, x].Value2.ToString());
                                }
                                catch {; }


                            }

                            




                            //goto NextRow;
                        }
                       
                            
                       

                    }

                    if(onewcopyloan.MobileNo !="")
                    {
                        onewcopyloan.Id = onewcopyloan.AddLoan(ref errm);
                    }
                    

                    NextRow:;



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
            }

		private void readExcel2019LOANBALANCES()
		{
			ExcelApp.Application excelApp = new ExcelApp.Application();
			ExcelApp.Workbook excelWorkbook = excelApp.Workbooks.Open(filename);
			ExcelApp.Worksheet excelWorksheet = excelWorkbook.Sheets[1];
			ExcelApp.Range excelRange = excelWorksheet.UsedRange;
			string data = string.Empty;
			int rowcount = excelRange.Rows.Count;
			int colcount = excelRange.Columns.Count;//7; // to be checked later
			int schemenocol = 0, InsuredCol = 0, advancescol = 0, loanfeescol = 0, totaldebitscol = 0, total = 0, balancecol = 0, idnumbercol = 0;
			int staffno = 0;


			string strdata = "";
			bool foundtheLoan = false;
			string errm = "";

			Application.DoEvents();


			if (excelWorksheet != null)
			{

				for (int i = 1; i <= rowcount; i++)
				{
					onewcopyloan = new Classes.CopyLoan();
					lblread.Text = i.ToString() + " Of " + rowcount.ToString();
					string amountpaid = "";
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
								goto NextRow;

							}





						}



						strdata = data.ToLower().Trim();
						DateTime fulldate = DateTime.Now;
						DateTime duedate = DateTime.Now;
						switch (strdata)
						{

							case "scheme":



								schemenocol  = x; //in case we are only reading one loan
								schemenocol  = x; InsuredCol = x + 3;advancescol = x + 8;loanfeescol = x + 11;totaldebitscol = x + 14;total = x + 17;balancecol = x + 21;idnumbercol = x + 24;staffno = x + 26;
								goto NextRow;
						}
						//i++;







						if (schemenocol  != 0)
						{
							if (i == 1) goto NextRow;
							if (x == schemenocol )
							{
								try
								{
									//trans.MonthName 

									onewcopyloan.PhoneNo   = (excelRange.Cells[i, x].Value2.ToString());
									x = x;
								}
								catch { }
							}
							if (x == InsuredCol)
							{
								try
								{
									onewcopyloan.Names = (excelRange.Cells[i, x].Value2.ToString());
								}
								catch {; }


							}
							if (x == advancescol)
							{
								try
								{
									
									onewcopyloan.Advance  = Convert.ToDouble(excelRange.Cells[i, x].Value2.ToString());

									//DateTime.TryParse((excelRange.Cells[i, x].Value2), out fulldate);
									////double val = (excelRange.Cells[i , x].Value2.ToString());
									//onewcopyloan.Fuldate  = (excelRange.Cells[i, x].Value2);
									//double d = double.TryParse(excelRange.Cells[i, x].Value2, out d);
									//double fulldat = double.Parse(excelRange.Cells[i, x].Value2);
									//DateTime conv = DateTime.FromOADate(fulldat);
									//onewcopyloan.Fulldate = conv ;
								}
								catch (Exception ex) { ex.Message.ToString(); }


							}
							if (x == loanfeescol )
							{
								try
								{
									onewcopyloan.LoanFees  = Double.Parse((excelRange.Cells[i, x].Value2.ToString()));
								}
								catch {; }


							}
							if (x == totaldebitscol )
							{
								try
								{
									onewcopyloan.TotalDebits  = Double.Parse((excelRange.Cells[i, x].Value2.ToString()));
								}
								catch {; }


							}
							if (x == total )
							{
								try
								{
									onewcopyloan.Credit  = Double.Parse((excelRange.Cells[i, x].Value2.ToString()));
								}
								catch {; }


							}
							if (x == balancecol )
							{
								try
								{
									onewcopyloan.Balance  = Double.Parse((excelRange.Cells[i, x].Value2.ToString()));
								}
								catch {; }


							}
							if (x == idnumbercol )
							{
								try
								{
									onewcopyloan.IdNumber  = excelRange.Cells[i, x].Value2.ToString();

								}
								catch (Exception ex) { ex.Message.ToString(); }


							}
							if (x == staffno)
							{
								try
								{
									onewcopyloan.StaffNo  = (excelRange.Cells[i, x].Value2.ToString());
								}
								catch {; }


							}
							
							//goto NextRow;
						}




					}

					if (onewcopyloan.IdNumber != "")
					{
						onewcopyloan.Id = onewcopyloan.AddDecLoan(ref errm);
					}


				NextRow:;



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
		}




		private void openFileDialog4_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void frmReadExcel_Load(object sender, EventArgs e)
        {

        }

		private void btn2019_Click(object sender, EventArgs e)
		{
			readExcel2019LOANBALANCES();
		}

        private void txtfile_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
    }


