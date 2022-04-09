using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class Loans
    {
        private int _loanId = 0;
        private int _memberid = 0;
        private int _loantypeid = 0;
        private string _loancode = "";
        private string _loanmanualref = "";
        private DateTime _loantransdate = DateTime.Today;
        private DateTime _loanappdate = DateTime.Today;
        private DateTime _loaneffectdate = DateTime.Today;
        private DateTime _loanenddate = DateTime.Today;
        private int _purposeid = 0;
        private double _loangrosspay = 0;
        private double _loannettpay = 0;
        private int _loanrepayperiod = 0;
        private double _loanamount = 0;
        private double _loanrepayamount = 0;
        private double _loaninterestrate = 0;
        private double _loaninterestamount = 0;
        private string _interesttype = "";
        private double _maximumloan = 0;
        private string _loanstatus = "";
        private double _loannettamount = 0;
        private double _loanamountpaid = 0;
        private double _approvedamount = 0;
        private string _lstatus = "";
        private bool _rejected = false;
        private bool _written = false;
        private bool _approved = false;
        private bool _readycheque = false;
        private bool _collected = false;
        private int _repaymode = 0;
        private DateTime _approveddate = DateTime.Today;
        private string _approvalrejectionremarks = "";
        private DateTime _rejecteddate = DateTime.Today;
        private string _chequetype = "";
        private DateTime _writingdate = DateTime.Today;
        private string _writingremarks = "";
        private double _membershares = 0;
        private string _posted = "";
        private DateTime _readydate = DateTime.Today;
        private string _readychequeno = "";
        private string _readyremarks = "";
        private double _collectedchequeamount = 0;
        private DateTime _collecteddate = DateTime.Today;
        private string _collectionremarks = "";
        private double _loanBalance = 0;
        private string _paymode = "";
        private double _newloanrepayamount = 0;
        private double _interestPaid = 0;
        private double _loanOrigAmount = 0;
        private double _amountQualified = 0;
        private double _totalSHares = 0;
        private double _freeShares = 0;
        private double _monthlychargesloaded = 0;
        private double _monthlychargesseparate = 0;
        private double _annualchargesloaded = 0;
        private double _annualchargesseparate = 0;
        private bool _newLoan = false;
        private bool _existingLoan = false;
        private double _monthOpeningBal = 0;
        private double _loanPenalty = 0;
        private bool _intFirst = false;
        private double _intPeriod = 0;
        private int _loanpenaltyrate = 0;
        private double _sumInterest = 0;
        private string _payType = "";
        private bool _affectsDR = false;
        private double _sumOtherCharges = 0;
        private string _volNo = "";
        private int _gracePeriod = 0;
        private bool _usedPartialDisbursement = false;
        private double _firstPartialAmount = 0;
        private string _createdBy = "";
        private DateTime _createdOn = DateTime.Today;
        private string _modifiedBy = "";
        private DateTime _modifiedOn = DateTime.Today;
        private bool _scheduled = false;
        private int _donorID = 0;
        private int _interestBalance=0;
        private double _interestamount = 0;
        private int _creditOfficerId=0;
        private string _membername = "";
        private string _memberno= "";
        private double _principalamount = 0;
        private string _loanTypename = "";
        public int LoanId { get { return _loanId; } set { _loanId = value; } }
        public int Memberid { get { return _memberid; } set { _memberid = value; } }
        public int Loantypeid { get { return _loantypeid; } set { _loantypeid = value; } }
        public string Loancode { get { return _loancode; } set { _loancode = value; } }
        public string Loanmanualref { get { return _loanmanualref; } set { _loanmanualref = value; } }
        public DateTime Loantransdate { get { return _loantransdate; } set { _loantransdate = value; } }
        public DateTime Loanappdate { get { return _loanappdate; } set { _loanappdate = value; } }
        public DateTime Loaneffectdate { get { return _loaneffectdate; } set { _loaneffectdate = value; } }
        public DateTime Loanenddate { get { return _loanenddate; } set { _loanenddate = value; } }
        public int Purposeid { get { return _purposeid; } set { _purposeid = value; } }
        public double Loangrosspay { get { return _loangrosspay; } set { _loangrosspay = value; } }
        public double Loannettpay { get { return _loannettpay; } set { _loannettpay = value; } }
        public int Loanrepayperiod { get { return _loanrepayperiod; } set { _loanrepayperiod = value; } }
        public double Loanamount { get { return _loanamount; } set { _loanamount = value; } }
        public double Loanrepayamount { get { return _loanrepayamount; } set { _loanrepayamount = value; } }
        public double Loaninterestrate { get { return _loaninterestrate; } set { _loaninterestrate = value; } }
        public double Loaninterestamount { get { return _loaninterestamount; } set { _loaninterestamount = value; } }
        public string Interesttype { get { return _interesttype; } set { _interesttype = value; } }
        public double Maximumloan { get { return _maximumloan; } set { _maximumloan = value; } }
        public string Loanstatus { get { return _loanstatus; } set { _loanstatus = value; } }
        public double Loannettamount { get { return _loannettamount; } set { _loannettamount = value; } }
        public double Loanamountpaid { get { return _loanamountpaid; } set { _loanamountpaid = value; } }
        public double Approvedamount { get { return _approvedamount; } set { _approvedamount = value; } }
        public string Lstatus { get { return _lstatus; } set { _lstatus = value; } }
        public bool Rejected { get { return _rejected; } set { _rejected = value; } }
        public bool Written { get { return _written; } set { _written = value; } }
        public bool Approved { get { return _approved; } set { _approved = value; } }
        public bool Readycheque { get { return _readycheque; } set { _readycheque = value; } }
        public bool Collected { get { return _collected; } set { _collected = value; } }
        public int Repaymode { get { return _repaymode; } set { _repaymode = value; } }
        public DateTime Approveddate { get { return _approveddate; } set { _approveddate = value; } }
        public string Approvalrejectionremarks { get { return _approvalrejectionremarks; } set { _approvalrejectionremarks = value; } }
        public DateTime Rejecteddate { get { return _rejecteddate; } set { _rejecteddate = value; } }
        public string Chequetype { get { return _chequetype; } set { _chequetype = value; } }
        public DateTime Writingdate { get { return _writingdate; } set { _writingdate = value; } }
        public string Writingremarks { get { return _writingremarks; } set { _writingremarks = value; } }
        public double Membershares { get { return _membershares; } set { _membershares = value; } }
        public string Posted { get { return _posted; } set { _posted = value; } }
        public DateTime Readydate { get { return _readydate; } set { _readydate = value; } }
        public string Readychequeno { get { return _readychequeno; } set { _readychequeno = value; } }
        public string Readyremarks { get { return _readyremarks; } set { _readyremarks = value; } }
        public double Collectedchequeamount { get { return _collectedchequeamount; } set { _collectedchequeamount = value; } }
        public DateTime Collecteddate { get { return _collecteddate; } set { _collecteddate = value; } }
        public string Collectionremarks { get { return _collectionremarks; } set { _collectionremarks = value; } }
        public double LoanBalance { get { return _loanBalance; } set { _loanBalance = value; } }
        public string Paymode { get { return _paymode; } set { _paymode = value; } }
        public double Newloanrepayamount { get { return _newloanrepayamount; } set { _newloanrepayamount = value; } }
        public double InterestPaid { get { return _interestPaid; } set { _interestPaid = value; } }
        public double LoanOrigAmount { get { return _loanOrigAmount; } set { _loanOrigAmount = value; } }
        public double AmountQualified { get { return _amountQualified; } set { _amountQualified = value; } }
        public double TotalSHares { get { return _totalSHares; } set { _totalSHares = value; } }
        public double FreeShares { get { return _freeShares; } set { _freeShares = value; } }
        public double Monthlychargesloaded { get { return _monthlychargesloaded; } set { _monthlychargesloaded = value; } }
        public double Monthlychargesseparate { get { return _monthlychargesseparate; } set { _monthlychargesseparate = value; } }
        public double Annualchargesloaded { get { return _annualchargesloaded; } set { _annualchargesloaded = value; } }
        public double Annualchargesseparate { get { return _annualchargesseparate; } set { _annualchargesseparate = value; } }
        public bool NewLoan { get { return _newLoan; } set { _newLoan = value; } }
        public bool ExistingLoan { get { return _existingLoan; } set { _existingLoan = value; } }
        public double MonthOpeningBal { get { return _monthOpeningBal; } set { _monthOpeningBal = value; } }
        public double LoanPenalty { get { return _loanPenalty; } set { _loanPenalty = value; } }
        public bool IntFirst { get { return _intFirst; } set { _intFirst = value; } }
        public double IntPeriod { get { return _intPeriod; } set { _intPeriod = value; } }
        public int Loanpenaltyrate { get { return _loanpenaltyrate; } set { _loanpenaltyrate = value; } }
        public double SumInterest { get { return _sumInterest; } set { _sumInterest = value; } }
        public string PayType { get { return _payType; } set { _payType = value; } }
        public bool AffectsDR { get { return _affectsDR; } set { _affectsDR = value; } }
        public double SumOtherCharges { get { return _sumOtherCharges; } set { _sumOtherCharges = value; } }
        public string VolNo { get { return _volNo; } set { _volNo = value; } }
        public int GracePeriod { get { return _gracePeriod; } set { _gracePeriod = value; } }
        public bool UsedPartialDisbursement { get { return _usedPartialDisbursement; } set { _usedPartialDisbursement = value; } }
        public double FirstPartialAmount { get { return _firstPartialAmount; } set { _firstPartialAmount = value; } }
        public string CreatedBy { get { return _createdBy; } set { _createdBy = value; } }
        public DateTime CreatedOn { get { return _createdOn; } set { _createdOn = value; } }
        public string ModifiedBy { get { return _modifiedBy; } set { _modifiedBy = value; } }
        public string MemberNo { get { return _memberno; } set { _memberno = value; } }
        public DateTime ModifiedOn { get { return _modifiedOn; } set { _modifiedOn = value; } }
        public bool Scheduled { get { return _scheduled; } set { _scheduled = value; } }
        public int DonorID { get { return _donorID; } set { _donorID = value; } }
        public int InterestBalance { get {return  _interestBalance;}
        set{_interestBalance =value;}}
        public double InterestAmount { get { return _interestamount; } set { _interestamount = value; } }
        public double PrincipalAmount { get { return _principalamount; } set { _principalamount = value; } }
        public string MemberName { get { return _membername; } set { _membername = value; } }
        public int CreditOfficerId
        { get { return _creditOfficerId; } set{ _creditOfficerId = value; } }
        string err = "";
        public string LoanTypeName { get { return _loanTypename; } set { _loanTypename = value; } }
        LoanTypes oLoanType = new LoanTypes();
        public string getMemberNo(  string  MemberNo,int loanTypeId )
        {
            MemberNo = "";
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_CheckMemberNoExistenceandLoan", 
               "@mcode", MemberNo ,
               "@LoanTypeId", loanTypeId);
            if (err == "")
            {
                if (rd.Read())
                {
                  
                   MemberNo = rd["MemberNo"].ToString();
                  
                }
                try { rd.Close(); }
                catch {; }

            }
            return MemberNo ;
        }
        public double  getInterestAmount( string MemberNo, int loanTypeId)
        {
            InterestAmount = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_CheckMemberNoExistenceandLoan", 
               "@mcode", MemberNo,
               "@LoanTypeId", loanTypeId);
            if (err == "")
            {
                if (rd.Read())
                {

                    this.InterestAmount = double.Parse(rd["InterestAmount"].ToString());
                    this.PrincipalAmount = Loanamountpaid - this.InterestAmount;

                }
                try { rd.Close(); }
                catch {; }

            }
            return InterestAmount;
        }
        //public double getLoan(ref string MemberNo, int loanTypeId, DateTime transDate)
        //{
        //    InterestAmount = 0;
        //    Link myLink = new Link();
        //    DbDataReader rd = myLink.GetDBResults(ref err, "proc_CheckMemberNoExistenceandLoan", "@TransDate", transDate,
        //       "@mcode", MemberNo,
        //       "@LoanTypeId", loanTypeId);
        //    if (err == "")
        //    {
        //        if (rd.Read())
        //        {

        //            this.InterestAmount = double.Parse(rd["InterestAmount"].ToString());

        //        }
        //        try { rd.Close(); }
        //        catch {; }

        //    }
        //    return InterestAmount;
        //}
        public string getMemberName( string MemberNo, int loanTypeId)
        {
            
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_CheckMemberNoExistenceandLoan",
               "@mcode", MemberNo,
               "@LoanTypeId", loanTypeId);
            if (err == "")
            {
                if (rd.Read())
                {

                    this.MemberName =rd["MemberName"].ToString();

                }
                try { rd.Close(); }
                catch {; }

            }
            return MemberName ;
        }
        public int getLoanId(string MemberNo, int loanTypeId)
        {

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_CheckMemberNoExistenceandLoan",
               "@mcode", MemberNo,
               "@LoanTypeId", loanTypeId);
            if (err == "")
            {
                if (rd.Read())
                {

                    this.LoanId = int.Parse(rd["LoanId"].ToString());
                    if(this.LoanId>0)
                    {
                        LoanTypes myLoanType = oLoanType.GetLoanType(loanTypeId);
                        this.LoanTypeName = myLoanType.LoanTypename;
                    }

                }
                try { rd.Close(); }
                catch {; }

            }
            return LoanId;
        }
    }
}
