using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel;
using BrightIdeasSoftware;

namespace ReadExcel
{
    public partial class frmImportFileFormat : Form
    {
        public frmImportFileFormat()
        {
            InitializeComponent();
        }
        Classes.FileImportFormat oFileImportFormat = new Classes.FileImportFormat();
        Classes.FileImportFormat oNewFileImportFormat = null;
        Classes.PayModes oPayModes = new Classes.PayModes();
        Classes.PayModes oNewPayModes = null;

        Classes.ShareTransactions oShareTransactions = new Classes.ShareTransactions();
        Classes.ShareTransactions oNewShareTransactions = null;
        Classes.ProductSetup oProductSetup = new Classes.ProductSetup();
        Classes.ProductSetup oNewProductSetup = null;
        Classes.ShareAndSavings oSavings = new Classes.ShareAndSavings();
        Classes.ShareAndSavings oNewSavings = new Classes.ShareAndSavings();
        //Classes.LoanRepayments oLoanRepayments = new Classes.LoanRepayments();
        //Classes.LoanRepayments oNewLoanRepayments = null;
        Classes.Loans oLoan = new Classes.Loans();
        Classes.Loans oNewLoan = null;
        Classes.Member oMember = new Classes.Member();
        Classes.Member oNewMember = null;
        Classes.ImportFileNames oImportFileNames = new Classes.ImportFileNames();
        Classes.ImportFileNames oNewImportFileNames = null;
        Classes.LoanTypes oLoanTypes = new Classes.LoanTypes();
        Classes.LoanTypes oNewLoanTypes = null;
        Classes.ShareTypes oShareTypes = new Classes.ShareTypes();
        Classes.ShareTypes oNewShareTypes = null;
        Classes.Banks oBank = new Classes.Banks();
        Classes.Banks oNewBank = null;
        Classes.Serials oSerials = new Classes.Serials();
        Classes.Serials oNewSerials = null;
        Classes.Receipt oReceipt = new Classes.Receipt();
        Classes.Receipt oNewReceipt = null;
        Classes.LoanRepayments oLoanRepayment = new Classes.LoanRepayments();
        Classes.LoanRepayments oNewLoanRepayment = null;



        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
         
        }

        private void frmImportFileFormat_Load(object sender, EventArgs e)
        {
            //we have gotten the name of the file and this will help us to get the file format to be looped
            cmbFormat.Items.Clear();
            ArrayList myList = new ArrayList();
            myList = oImportFileNames  .GetImportFileNames();
            foreach (Classes.ImportFileNames  oimportname in myList)
            {
                // string fileformat = oimport.FormatName;
                cmbFormat.Items.Add(new ItemData.itemData(oimportname.ImportFileName , oimportname));
            }
        }
        private void ReadAndMigrateExcelFile()
        {
            Application.DoEvents();
            int s = 1;
            if (chkIsHeader.Checked)
                s = 2;
            ExcelApp.Application excelapp = new ExcelApp.Application();
            ExcelApp.Workbook xlWorkbook = excelapp.Workbooks.Open(txtFilePath.Text);
            ExcelApp.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            ExcelApp.Range xlRange = xlWorksheet.UsedRange;
            if (xlWorksheet != null)
            {
                if (oNewImportFileNames == null)
                {
                    MessageBox.Show("File Import Format Name  Is Required", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    cmbFormat.Focus();
                    return;

                }
                ArrayList newList = new ArrayList();
                ArrayList LoanList = new ArrayList();
                ArrayList FileList = new ArrayList();
                for (int i = s; i <= xlRange.Rows.Count; i++)
                {
                    bool exist = true;
                    string membercode = "";
                    oNewMember = new Classes.Member();
                  
                    FileList = oFileImportFormat.GetFileImportFormatsByImportFileName(oNewImportFileNames.ImportFileNameId);
                    lblReadData.Text = i.ToString() + " Of  " + xlRange.Rows.Count;
                    string error = "";
                    for (int j = 1; j < xlRange.Columns.Count; j++)//looping through positions to the arraylist
                    {
                        //we might need to check for the 
                        oNewFileImportFormat = oFileImportFormat.proc_GetFileFormatDetails(oNewImportFileNames.ImportFileNameId, j);
                        if(oNewFileImportFormat !=null)
                        {

                      
                        //foreach (Classes.FileImportFormat oimport in FileList) //
                        //{
                        //    if (oimport.Position == j)
                        //    {
                        // to obtain memberno
                        if (xlRange.Cells[i, oNewFileImportFormat.Position] != null && xlRange.Cells[i, oNewFileImportFormat.Position].Value2 != null)
                        {
                            try
                            {
                                membercode = xlRange.Cells[i, oNewFileImportFormat.Position].Value2.ToString();
                                oNewMember.Mcode = oMember.getMemberNo(membercode);
                                oNewMember.MemberName = oMember.getMemberName(membercode);
                                if (oNewMember.Mcode == "")
                                {
                                    exist = false;

                                }
                            }
                            catch {; }
                        }

                        DateTime TransDate = DateTime.Now;

                        oNewSavings = new Classes.ShareAndSavings();
                        oNewLoan = new Classes.Loans();

                       

                         
                        if (!oNewFileImportFormat.IsLoan)
                        { 
                        if (oNewSavings != null)
                        {


                            if (xlRange.Cells[i, oNewFileImportFormat.Position] != null && xlRange.Cells[i, oNewFileImportFormat.Position].Value2 != null)
                            {
                                try
                                {
                                        oNewSavings.Amount = xlRange.Cells[i, oNewFileImportFormat.Position].Value2.ToString();
                                       
                                        if(oNewMember!=null)
                                            { 
                                            oNewSavings.MemberNo  = oNewMember.Mcode;
                                                oNewSavings.MemberName = oNewMember.MemberName;
                                            }
                                            oNewSavings.ShareTypeId = oNewFileImportFormat.ProductId;

                                        newList.Add(oNewSavings);
                                }
                                catch {; }
                            }

                        }
                        }
                        else
                        {
                                if (oNewLoan != null)
                                {


                                    if (xlRange.Cells[i, oNewFileImportFormat.Position] != null && xlRange.Cells[i, oNewFileImportFormat.Position].Value2 != null)
                                    {
                                        try
                                        {
                                            oNewLoan.Loantypeid  = xlRange.Cells[i, oNewFileImportFormat.ProductId].Value2.ToString();
                                            oNewLoan.MemberNo = oNewMember.Mcode;

                                            oNewLoan.MemberName = oNewMember.MemberName; 
                                        //oNewLoan.LoanId = oNewLoan.LoanId;
                                        oNewLoan.LoanId = oLoan.getLoanId(membercode, oNewFileImportFormat.ProductId);
                                            oNewLoan.Loanamountpaid = xlRange.Cells[i, oNewFileImportFormat.Position].Value2.ToString();

                                            oNewLoan.InterestAmount = oLoan.getInterestAmount(membercode, oNewFileImportFormat.ProductId);
                                           

                                            //We need to check the memberno from the database and existence of that loan by sending the memberno and loantypeid,return member name ,mcode,interest amount,

                                            LoanList.Add(oNewLoan);
                                        }
                                        catch {; }
                                    }

                                }
                            
                        }
                           }
                        //}

                    }

                    // onewMember.ImportMembers(ref error); we shall not insert we shall display on the list

                }

                objListLoans.SetObjects(LoanList );
                objShareTransactionss.SetObjects(newList);
            }
        }
        private void ReadAndMigrateExcelFile2()
        {
            Application.DoEvents();
            int s = 1;
            if (chkIsHeader.Checked)
                s = 2;
            ExcelApp.Application excelapp = new ExcelApp.Application();
            ExcelApp.Workbook xlWorkbook = excelapp.Workbooks.Open(txtFilePath.Text);
            ExcelApp.Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            ExcelApp.Range xlRange = xlWorksheet.UsedRange;
            if (xlWorksheet != null)
            {
                if (oNewImportFileNames == null)
                {
                    MessageBox.Show("File Import Format Name  Is Required", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    cmbFormat.Focus();
                    return;

                }
                ArrayList newList = new ArrayList();
                ArrayList LoanList = new ArrayList();
                ArrayList FileList = new ArrayList();//Items to be looped
                ArrayList InvalidSavingsList = new ArrayList();
                ArrayList InvalidLoansList = new ArrayList();


                double Totalsavings = 0;
                double Totalloanrepayments = 0;
                double totalinterest = 0;
                double totalprincipal = 0;
                double totalinvalidsavings = 0;
                double totalinvalidloans = 0;
                double sumtotal = 0;
                for (int i = s; i <= xlRange.Rows.Count; i++)
                {
                    bool exist = true;
                    string membercode = "";
                    oNewMember = new Classes.Member();

                    //FileList = oFileImportFormat.GetFileImportFormatsByImportFileName(oNewImportFileNames.ImportFileNameId);
                    lblReadData.Text = i.ToString() + " Of  " + xlRange.Rows.Count;
                    string error = "";
                    for (int j = 1; j < xlRange.Columns.Count; j++)//looping through positions on  the arraylist
                    {
                        //we might need to check for the 
                        oNewFileImportFormat = oFileImportFormat.proc_GetFileFormatDetails(oNewImportFileNames.ImportFileNameId, j);
                        if (oNewFileImportFormat != null)
                        {
                           
                            //{
                            if(oNewFileImportFormat.ProductId==0)
                            {

                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                    try
                                {
                                    membercode = xlRange.Cells[i, j].Value2.ToString();
                                        oNewMember = oMember.GetMemberByCode(membercode);
                                        if(oNewMember !=null)
                                        { 
                                            if(oNewMember.Mcode!="")
                                            {
                                                oNewMember.Mcode = oNewMember.Mcode;
                                                oNewMember.MemberName = oNewMember.FullName;
                                            }
                                            else
                                            {
                                                oNewMember.Mcode = membercode;
                                                
                                            }
                                      
                                   
                                        }
                                    }
                                catch {; }
                            }
                            else
                            {

                            

                            //DateTime TransDate =DateTime.Now;
                            //        DateTime.TryParse("20220101", out TransDate);

                            oNewSavings = new Classes.ShareAndSavings();
                            oNewLoan = new Classes.Loans();




                            if (!oNewFileImportFormat.IsLoan)
                            {
                                if (oNewSavings != null)
                                {


                                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                    {
                                        try
                                        {
                                                    double amount = 0;
                                                    double.TryParse((xlRange.Cells[i, j].Value2.ToString()), out amount);
                                                    oNewSavings.Amount = amount;

                                            if (oNewMember != null)
                                            {
                                                oNewSavings.MemberNo = oNewMember.Mcode;
                                                oNewSavings.MemberName = oNewMember.FullName;
                                                    if (oNewMember.Mcode == "")
                                                    {
                                                        oNewMember.Mcode = membercode;

                                                    }
                                                }
                                            oNewSavings.ShareTypeId = oNewFileImportFormat.ProductId;
                                                
                                                oNewSavings.Sharename = oNewFileImportFormat.ProductName;
                                                    if(oNewSavings.Amount>0)
                                                    {
                                                    if(oNewMember.Mcode != "")
                                                    {
                                                        Totalsavings += oNewSavings.Amount;
                                                        newList.Add(oNewSavings);
                                                    }
                                                    else
                                                    {
                                                        InvalidSavingsList .Add(oNewSavings);
                                                        totalinvalidsavings += oNewSavings.Amount;
                                                    }
                                                }
                                               
                                           
                                        }
                                        catch {; }
                                    }

                                }
                            }
                            else
                            {
                                if (oNewLoan != null)
                                {


                                    if (xlRange.Cells[i,j] != null && xlRange.Cells[i, j].Value2 != null)
                                    {
                                        try
                                        {
                                                    oNewLoan.Loantypeid = oNewFileImportFormat.ProductId;
                                                oNewLoan.LoanTypeName = oNewFileImportFormat.ProductName;
                                            oNewLoan.MemberNo = oNewMember.Mcode;

                                            oNewLoan.MemberName = oNewMember.FullName;
                                                //oNewLoan.LoanId = oNewLoan.LoanId;
                                                //  oNewLoan.LoanId = oLoan.getLoanId(oNewMember.Mcode, oNewFileImportFormat.ProductId);
                                                // oNewLoan.InterestAmount = oLoan.getInterestAmount(membercode, oNewFileImportFormat.ProductId);



                                                double amountpaid = 0;
                                                double.TryParse((xlRange.Cells[i, j].Value2.ToString()), out amountpaid);
                                                oNewLoan.Loanamountpaid = amountpaid;

                                                //  oNewLoan.PrincipalAmount = (amountpaid - oNewLoan.InterestAmount);


                                                //We need to check the memberno from the database and existence of that loan by sending the memberno and loantypeid,return member name ,mcode,interest amount,
                                                if (oNewLoan.Loanamountpaid >0)
                                                    {
                                                    if(oNewMember.Mcode!="")
                                                    {
                                                        LoanList.Add(oNewLoan);
                                                        Totalloanrepayments += oNewLoan.Loanamountpaid;
                                                    }
                                                    else
                                                    {
                                                        InvalidLoansList.Add(oNewLoan);
                                                        totalinvalidloans  += oNewLoan.Loanamountpaid;
                                                    }
                                                  
                                                    //totalinterest +=oNewLoan.InterestAmount;
                                                    //totalprincipal +=oNewLoan.PrincipalAmount;

                                                        
                                                    }
                                                   
                                        }
                                        catch {; }
                                    }

                                }

                            }
                            }
                        }
                        //}

                    }

                    // onewMember.ImportMembers(ref error); we shall not insert we shall display on the list

                }
                sumtotal = Totalloanrepayments + Totalsavings;
                lblTotalsavingsValue .Text  = Totalsavings.ToString();
                //lblTotalPrincipal.Text = totalprincipal.ToString();
                lblTotalAMOUNT.Text = Totalloanrepayments.ToString();
                lblinvalidsaving.Text = totalinvalidsavings.ToString();
                lblInvalidsavings.Text = totalinvalidloans.ToString();
                //lblInterest.Text = totalinterest.ToString();
                lblTotals.Text = sumtotal.ToString();
                
                


                objListLoans.SetObjects(LoanList);
                objShareTransactionss.SetObjects(newList);
                objListInvalidLoans.SetObjects(InvalidLoansList);
                objListInvalidsavings.SetObjects(InvalidSavingsList);
                chkCheckAll.Checked = true;
            }
        }
        private void cmbFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            object obj = ((ItemData.itemData)(cmbFormat.SelectedItem))._itemData;
            oNewImportFileNames  = (Classes.ImportFileNames )obj;//this is the selected filename and we shall use the filename id to search for the fileformat list to be looped 
       
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();
            txtFilePath.Text = openFileDialog1.FileName;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void objListLoans_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        public void ClearText()
        {
            oNewFileImportFormat = null;
            oNewImportFileNames = null;
            txtAmount.Text = "";
            txtbankGl.Text = "";
            txtDocumentNo.Text = "";
            txtFilePath.Text = "";
            txtModeofpayment.Text = "";
            oNewBank = null;
            objListLoans.Clear();
            objShareTransactionss.Clear();
            lblTotals.Text = "0";
            //lblTotalsavings.Text = "";
            lblTotalAMOUNT.Text = "";
           
            lblTotalsavingsValue.Text = "0";
            lblReadData.Text = "0";


        }

        private void btnPost_Click(object sender, EventArgs e)
        {
            if(objListLoans.Items.Count==0&&objShareTransactionss.Items.Count==0)
            {
                 MessageBox.Show("Import the data to be posted", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                   
                    return;

                
            }
            if (txtDocumentNo.Text .Trim()=="")
            {
                MessageBox.Show("Document No required", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtDocumentNo.Focus();
                return;


            }
            if (txtAmount.Text.Trim() == "")
            {
                MessageBox.Show("Enter Amount Received In Bank", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtDocumentNo.Focus();
                return;


            }
            if (oNewPayModes == null)
            {
                MessageBox.Show("PayModes Is Required", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtModeofpayment.Focus();
                return;


            }
            if (txtbankGl.Text  == "")
            {
                MessageBox.Show("Debit GL Is Required", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtbankGl.Focus();
                return;


            }
            if (objListInvalidLoans.Items.Count >0)
            {
                MessageBox.Show("Please validate MemberNo and Re-Import the data before posting ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabPage3.Focus();
                objListInvalidLoans.Focus();
               
                return;


            }
            if (objListInvalidsavings .Items.Count > 0)
            {
                MessageBox.Show("Please validate MemberNo and Re-Import the data before posting ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabPage3.Focus();
                objListInvalidLoans.Focus();

                return;


            }
            string error = "";
            double amountmigrated = 0;
            double.TryParse(lblTotals.Text, out amountmigrated);
            double amountReceived = 0;
            double.TryParse(txtAmount .Text , out amountReceived);
            if (amountReceived!= amountmigrated)
            {
                MessageBox.Show("Amount Mismatch", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtAmount.Focus();
                return;
            }
            string createdBy = "Import";
            int serialid = oSerials.GetSerialId (createdBy);
            int receiptId = oReceipt.GetReceiptId(createdBy, amountReceived, oNewPayModes.PaymentModeName, dtPTransDate.Value, serialid);


            for (int i = 0; i < objShareTransactionss.Items.Count; i++)
            {
                if (oNewSavings == null)
                    oNewSavings = new Classes.ShareAndSavings();

                if (objShareTransactionss .Items[i].Checked)
                {
                    Classes.ShareAndSavings osavingshare = (Classes.ShareAndSavings)objShareTransactionss.GetModelObject(i);

                   

                    if (oNewSavings != null)
                    {
                        oNewSavings.Amount = osavingshare.Amount  ;
                        oNewSavings.ChequeNumber  =txtDocumentNo.Text ;
                       oNewSavings.SerialID = serialid;
                        oNewSavings.ProductTypeId = 1;
                        oNewSavings.ShareTypeId = osavingshare.ShareTypeId;
                        oNewSavings.TransDate = dtPTransDate.Value;
                        oNewSavings.GLDR = txtbankGl.Text;
                        oNewSavings.ReceiptId = receiptId;
                        oNewSavings.MemberNo = osavingshare.MemberNo;
                        oNewSavings.TransId = oNewSavings.PostSharesAndSavings(ref error);
                       

                    }
                }
                oNewSavings = null;

            }
            for (int i = 0; i < objListLoans .Items.Count; i++)
            {
                if (oNewLoanRepayment == null)
                    oNewLoanRepayment = new Classes.LoanRepayments ();

                if (objListLoans .Items[i].Checked)
                {
                    Classes.Loans olon = (Classes.Loans)objListLoans.GetModelObject(i);



                    if (oNewLoanRepayment != null)
                    {
                        oNewLoanRepayment.PaymentAmount  = olon.Loanamountpaid ;
                        oNewLoanRepayment.ChequeNo = txtDocumentNo.Text;
                        oNewLoanRepayment.SerialId = serialid;
                        oNewLoanRepayment.ProductTypeId = 2;
                        oNewLoanRepayment.ProductId = olon.Loantypeid;
                        oNewLoanRepayment.PaymentDate = dtPTransDate.Value;
                       oNewLoanRepayment.GLDR = txtbankGl.Text;
                        oNewLoanRepayment.ReceiptId = receiptId;
                        oNewLoanRepayment.MemberNo = olon.MemberNo;

                        oNewLoanRepayment.RepaymentId = oNewLoanRepayment .PostLoanRepayments (ref error);


                    }
                }
                oNewLoanRepayment  = null;

            }
            if (error == "")
            {
                MessageBox.Show("Process succeeded", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearText();

            }
            else
            {
                MessageBox.Show(error, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void btnImpo_Click(object sender, EventArgs e)
        {
            btnImpo.Enabled = false;
            Cursor = Cursors.WaitCursor;
            ReadAndMigrateExcelFile2();

            btnImpo.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void btnPayMode_Click(object sender, EventArgs e)
        {
            frmSearchPayModes frm = new frmSearchPayModes();
            frm.ShowDialog();
            oNewPayModes = oPayModes.GetPayMode(frm.selInt);
            if(oNewPayModes!=null)
            {
                txtModeofpayment.Text = oNewPayModes.PaymentModeName;
                //retrieve banks 
                if(oNewPayModes .PaymentModeName .Trim().ToLower ()=="cheque")
                {
                    ArrayList myList = new ArrayList();
                    myList = oBank.GetBanks();
                    foreach (Classes.Banks obankname in myList)
                    {
                        // string fileformat = oimport.FormatName;
                        cmbBank.Items.Add(new ItemData.itemData(obankname.Bankname, obankname));
                    }

                }
            }

        }

        private void cmbBank_SelectedIndexChanged(object sender, EventArgs e)
        {
            object obj = ((ItemData.itemData)(cmbBank .SelectedItem))._itemData;
            oNewBank = (Classes.Banks )obj;//this is the selected filename and we shall use the filename id to search for the fileformat list to be looped 
            if(oNewBank!=null)
            {
                txtbankGl.Text = oNewBank.BankGLCode.ToString();
            }
        }

        private void objShareTransactionss_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void chkCheckAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (OLVListItem olv in objShareTransactionss .Items)
            {
                olv.Checked = chkCheckAll.Checked;
            }
            foreach (OLVListItem olv in objListLoans.Items)
            {
                olv.Checked = chkCheckAll.Checked;
            }
        }
    }
}
