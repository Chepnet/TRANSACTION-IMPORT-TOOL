using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;

namespace ReadExcel.Classes
{
    class Transaction
    {
        public int TransID { get; set; }
        public int LoanRepaymentId { get; set; }
        public int MemberID { get; set; }
        public int LoanTypeId { get; set; }
        public int LoanID { get; set; }
        public int SheetNo { get; set; }
        public double openingbalance { get; set; }
        public string  Membername { get; set; }
        public string MemberNumber { get; set; }
        public string StaffNumber { get; set; }
        public string LoanTypeName { get; set; }
        public string LoanAmount { get; set; }
        public double Amount { get; set; }
        public string MonthName { get; set; }
        public int PeriodMonth { get; set; }
        public int PeriodYear { get; set; }

        public int EmployerId { get; set; }
        public DateTime TransDate { get; set; }
        public DateTime ValueDate { get; set; }
        public string  TransDate1 { get; set; }
        public string  strAmount { get; set; }
        public string  strInterest { get; set; }
        public string strPrincipal { get; set; }
        public string  strBalance { get; set; }

        public int ShareTypeId { get; set; }
        public bool IsTopTrans { get; set; }
        public double Debit { get; set; }
        public double Credit { get; set; }
        public double Balance { get; set; }
        public string Particulars { get; set; }
        string errmsg = "";

        public int AddEdditBarabaramembers(ref string errm)
        {
            int Memberid = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_AddBarabaraMembers",
              "@memberid", this.MemberID  ,//MEMBERID
             "@mpayroll", this.StaffNumber  ,//STAFF NUMBER
            "@mcode", ""  ,
            "@msurname", this.Membername       );
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    Memberid  = int.Parse(rd["MemberId"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (Memberid);
        }

        public int AddEdditLongHornmembers(ref string errm)
        {
            int Memberid = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "pro_AddEditMembers",
             "@memberid", this.MemberID,//MEMBERID
             "@employerId",this.EmployerId,
            "@membername", this.Membername);
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    Memberid = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (Memberid);
        }
        public int AddEdditLongHornLoans(ref string errm)
        {
            int Loanid = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "pro_MigrateLoans",
               //"@LoanId", this.LoanID,
             "@memberid", this.MemberID,//MEMBERID
             "@LoantypeId", this.LoanTypeId,
            "@LoanAmount", this.Amount);
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    //Loanid = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (Loanid);
        }
        public int AddEdditBarabaraLoans(ref string errm)
        {
            int Loanid = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_AddBarabaraLoans",
              "@LoanId", this.LoanID ,
              "@memberid", this.MemberID ,
             //   "@loantypeid", this.lo,
                "@loanamount", this.LoanAmount  );
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    Loanid = int.Parse(rd["LoanId"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (Loanid);
        }
        public int AddEdditLongHornShareTransactions(ref string errm)
        {

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "pro_AddEditSharetransactions",
              "@TransId", this.TransID,
              "@memberid", this.MemberID,
              "@TransDate",this.TransDate,
              "@amount", this.Amount);
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    TransID  = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (TransID);
        }
        public int AddEdditBarabaraShareTransactions(ref string errm)
        {
            
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_migrateBarabaraSharetransactions",
              "@TransId", this.TransID ,
              "@memberid", this.MemberID,
              "@openingbalance",this.openingbalance ,
              "@transDate", this.TransDate1  ,
              "@amount", this.Amount);
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                     //TransID  = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (TransID );
        }
        public int AddEdditBarabaraLoanrepayments(ref string errm)
        {
            int Loanrepaymentid = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_AddBarabaraLoanrepayment",
              "@RepaymentId", this.LoanRepaymentId,
                "@memberid", this.MemberID ,
                "@LoanId", this.LoanID , 
                "@PaymentDate", this.TransDate ,
                "@PaymentAmount", this.strAmount ,
                "@Principal", this.strPrincipal ,
                "@Interest", this.strInterest ,
                "@LoanAmount", this.LoanAmount,
                 "@LoanBalance", this.strBalance );
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    Loanrepaymentid = int.Parse(rd["LoanRepaymentId"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (Loanrepaymentid);
        }

        public int AddEdditTrans( ref string errm){
         int done = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_AddEdditbarabaratrans",
            "@sheetno", this.SheetNo,
            "@membername", this.Membername,
            "@membernumber", this.MemberNumber,
             "@staffnumber", this.StaffNumber,
            "@loantypename", this.LoanTypeName,
            "@loanamount", this.LoanAmount,
            "@transdate", this.TransDate,
            "@stramount", this.strAmount,
            "@strinterest", this.strInterest,
            "@strprincipal", this.strPrincipal,
            "@strbalance", this.strBalance,
            "@istoptrans", this.IsTopTrans
            );
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    done = int.Parse(rd["TRANSID"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch { ;}
            }
            return (done);
        }
        public int topupTrans(ref string errm)
        {
            int done = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_AddTopupbarabaratrans",
            "@sheetno", this.SheetNo,
            "@membername", this.Membername,
            "@membernumber", this.MemberNumber,
             "@staffnumber", this.StaffNumber,
            "@loantypename", "Top Up Loan",
            "@loanamount", this.LoanAmount,
            "@transdate", this.TransDate,
            "@stramount", this.strAmount,
            "@strinterest", this.strInterest,
            "@strprincipal", this.strPrincipal,
            "@strbalance", this.strBalance,
            "@istoptrans", this.IsTopTrans
            );
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    done = int.Parse(rd["TRANSID"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch { ;}
            }
            return (done);
        }
        public int AddEdditBarabaradirectdeposit(ref string errm)
        {

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_Insertdirectdeposit",
                "@transId", this.TransID ,
                                "@transdate", this.TransDate1 ,
                                "@memberno", this.MemberNumber ,
                                "@payroll", this.StaffNumber ,
                                "@membername", this.Membername ,
                                "@amount", this.Amount 
                                
                                );
            errm = errmsg;
            if (errmsg == "")
            {
                //if (rd.Read())
                //{
                //    TransID = int.Parse(rd["Id"].ToString());
                //}
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (TransID);
        }
        public int AddEdditbarabaraTrans(ref string errm)
        {

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_migrateBARABARAloantopup",
                                 "@loanId", this.TransID,
                                "@memberno", this.MemberNumber ,
                                "@membername", this.Membername,
                                "@Amount", this.Amount,
                                "@loantypename", this.LoanTypeName ,
                                "@repaymentperiod", this.ShareTypeId ,
                                "@loanrepayamount", this.Credit 
                                );
            errm = errmsg;
            if (errmsg == "")
            {
                //if (rd.Read())
                //{
                //    TransID = int.Parse(rd["Id"].ToString());
                //}
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (TransID);
        }

        public int AddEdditUniqueEquityTrans(ref string errm)
        {

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_AddEditEquityTrans",
                                 "@TransId", this.TransID,
                                "@TransDate", this.TransDate,
                                "@TransParticulars", this.Particulars,
                                "@ValueDate", this.ValueDate,
                                "@Debit", this.Debit,
                                "@Credit", this.Credit,
                                "@Balance",this.Balance
                                );
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    TransID = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch {; }
            }
            return (TransID);
        }

    }
}
