using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;

namespace ReadExcel.Classes
{
    class Loan
    {
        
        private string _memberno = "";
        private string   _loantransdate ="" ;
        private double _loanamount = 0;
        private double _interest = 0;
        private double _sumothercharges = 0;
        private double _originalamount = 0;
        private int loanid = 0;
        private int _remainingperiod = 0;
        private int _loanno = 0;
        private string _refno = "";
        private int transid = 0;
        private int loantypeid = 0;
        private double monthlyinterest = 0;
        private string _IDNO = "";
        private double period = 0;
        private double _paymentamount = 0;


        private string _membername = "";
        public string MemberNo { get { return _memberno ; } set { _memberno  = value; } }
        public string MemberName { get { return _membername; } set { _membername = value; } }
        public string    Loantransdate { get { return _loantransdate; } set { _loantransdate = value; } }

        public double Loanamount { get { return _loanamount; } set { _loanamount = value; } }
        public double Paymentamount { get { return _paymentamount; } set { _paymentamount = value; } }
        public double Interest { get { return _interest ; } set { _interest  = value; } }
        public double SumOtherCharges { get { return _sumothercharges ; } set { _sumothercharges  = value; } }
        public double OriginalAmount { get { return _originalamount ; } set { _originalamount  = value; } }

        public int LoanId { get { return loanid ; } set { loanid = value; } }
        public int TransId { get { return transid ; } set { transid  = value; } }
        public string IDNo { get { return _IDNO; } set { _IDNO = value; } }
        public int LoanTypeId { get { return loantypeid; } set { loantypeid = value; } }

        public double Period { get { return period; } set { period = value; } }
        public double  MonthlyInterest { get { return monthlyinterest; } set { monthlyinterest = value; } }
        public int RemainingPeriod { get { return _remainingperiod ; } set { _remainingperiod = value; } }
        public int LoanNo { get { return _loanno ; } set { _loanno  = value; } }

        public string RefNo { get { return _refno ; } set { _refno  = value; } }

        string err = "";
        public int AddEditLoan(ref string error)
        {
            int id = 0;
            Link myLink = new ReadExcel.Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_migrateMajicareLoans", "@LoanId", this.LoanId,
                "@memberno", this.MemberNo,
                "@Transdate", this.Loantransdate,
                "@LoanAmount", this.Loanamount,
                "@MemberName", this.MemberName,
                "@IdNo", this.IDNo,
                "@Interest", 0,
                "@OriginalLoanAmount", 0,
                "@SumOtherCharges", 0,
                "@Remainingperiod", this.RemainingPeriod ,
                "@LoanTypeId", this.LoanNo );
            error = err;
            if(err=="")
            {
                if(rd.Read ())
                {
                    //id = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); }
                catch {; }
            }
            return id;
        }
        public int AddEditShareTransactions(ref string error)
        {
            int id = 0;
            Link myLink = new ReadExcel.Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_MigrateMajicareShareTransactions", "@TransId", this.TransId ,
                "@memberno", this.MemberNo,
                "@Membername",this.MemberName,
                 "@IDNO", this.IDNo,
                "@transdate", this.Loantransdate,
                "@Amount", this.Loanamount );
            error = err;
            if (err == "")
            {
                if (rd.Read())
                {
                    //id = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); }
                catch {; }
            }
            return id;
        }

        public int AddEditNascaLoan(ref string error)
        {
            int id = 0;
            Link myLink = new ReadExcel.Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_MigrateNascaLoans", "@LoanId", this.LoanId,
                "@mcode", this.MemberNo,
                "@LoanTransDate",this.Loantransdate,
                 "@MonthlyInterest", this.MonthlyInterest ,
                 "@Period",this.Period,
                "@LoanAmount", this.Loanamount,
                "@LoanTypeId", this.LoanTypeId,
                "@Description", this.MemberName,
                "@Interestrate",this.Interest

              );
            error = err;
            if (err == "")
            {
                if (rd.Read())
                {
                    id = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); }
                catch {; }
            }
            return id;
        }
        public int AddEditBarabaraLoan(ref string error)
        {
            int id = 0;
            Link myLink = new ReadExcel.Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_AddBarabaraLoans", "@LoanId", this.LoanId,
                "@mcode", this.MemberNo,
                "@LoanAmount", this.Loanamount,
                "@LoanType", this.RefNo,
                "@CustomerName", this.MemberName

              );
            error = err;
            if (err == "")
            {
                if (rd.Read())
                {
                    id = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); }
                catch {; }
            }
            return id;
        }
        public int AddEditKRBLoan(ref string error)
        {
            int id = 0;
            Link myLink = new ReadExcel.Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_migrateKRBLoans", "@LoanId", this.LoanId,
                "@mcode", this.MemberNo,
                "@TransDate", this.Loantransdate,
                 "@Period", this.Period,
                "@LoanAmount", this.Loanamount,
                "@LoanTypeId", this.LoanTypeId,
                "@Description", this.MemberName,
                "@PaidAmount", this.Paymentamount,
                "@RepayAmount", this.OriginalAmount

              );
            error = err;
            if (err == "")
            {
                if (rd.Read())
                {
                    id = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); }
                catch {; }
            }
            return id;
        }

    }
}
