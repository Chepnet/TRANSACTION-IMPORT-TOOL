using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;

namespace ReadExcel.Classes
{
    class CopyLoan
    {
        private int _id = 0;
        private string _refNo = "";
        private string _mobileNo = "";
        private DateTime  _fulldate = DateTime.Now ;
        private double _principal = 0;
		private string  _fuldate = "";
		private double _creditAmount = 0;
        private double _interest = 0;
        private double _mpesafee = 0;
        private string _duedate = "";
        private string _empfirstname = "";
        private string _trnature = "";
        private string _empsurname = "";
        private string _trdetails = "";
        private double _payments = 0;
        private double _insurance = 0;
        private string _schemename = "";
        private int _schemeId = 0;
        private int _memberid = 0;

		private string _phoneNo = "";
		private string _names = "";
		private double _advance = 0;
		private double _loanFees = 0;
		private double _totalDebits = 0;
		private double _credit = 0;
		private double _balance = 0;
		//private int _id = 0;
		private string _scheme = "";
		//private int _schemeId = 0;
		private string _idNumber = "";
		private string _staffNo = "";
		private double _rate = 0;
		private int _memberId = 0;
		public string PhoneNo { get { return _phoneNo; } set { _phoneNo = value; } }
		public string Names { get { return _names; } set { _names = value; } }
		public double Advance { get { return _advance; } set { _advance = value; } }
		public double LoanFees { get { return _loanFees; } set { _loanFees = value; } }
		public double TotalDebits { get { return _totalDebits; } set { _totalDebits = value; } }
		public double Credit { get { return _credit; } set { _credit = value; } }
		public double Balance { get { return _balance; } set { _balance = value; } }
		//public int Id { get { return _id; } set { _id = value; } }
		public string Scheme { get { return _scheme; } set { _scheme = value; } }
		//public int SchemeId { get { return _schemeId; } set { _schemeId = value; } }
		public string IdNumber { get { return _idNumber; } set { _idNumber = value; } }
		public string StaffNo { get { return _staffNo; } set { _staffNo = value; } }
		public double Rate { get { return _rate; } set { _rate = value; } }
		public int MemberId { get { return _memberId; } set { _memberId = value; } }


		public int Id { get { return _id; } set { _id = value; } }
        public string RefNo { get { return _refNo; } set { _refNo = value; } }
        public string MobileNo { get { return _mobileNo; } set { _mobileNo = value; } }
        public DateTime   Fulldate { get { return _fulldate; } set { _fulldate = value; } }
		public string  Fuldate { get { return _fuldate; } set { _fuldate = value; } }
		public double Principal { get { return _principal; } set { _principal = value; } }
        public double CreditAmount { get { return _creditAmount; } set { _creditAmount = value; } }
        public double Interest { get { return _interest; } set { _interest = value; } }
        public double Mpesafee { get { return _mpesafee; } set { _mpesafee = value; } }
        public string  Duedate { get { return _duedate; } set { _duedate = value; } }
        public string Empfirstname { get { return _empfirstname; } set { _empfirstname = value; } }
        public string Trnature { get { return _trnature; } set { _trnature = value; } }
        public string Empsurname { get { return _empsurname; } set { _empsurname = value; } }
        public string Trdetails { get { return _trdetails; } set { _trdetails = value; } }
        public double Payments { get { return _payments; } set { _payments = value; } }
        public double Insurance { get { return _insurance; } set { _insurance = value; } }
        public string Schemename { get { return _schemename; } set { _schemename = value; } }
        public int SchemeId { get { return _schemeId; } set { _schemeId = value; } }
        public int Memberid { get { return _memberid; } set { _memberid = value; } }



        string err = "";
        public int AddLoan(ref string error)
        {
            int id = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_AddCopyLoan", "@Id", this.Id,
                "@refNo", this.RefNo,
                "@MobileNo", this.MobileNo,
                "@fulldate", this.Fuldate ,
                "@Principal", this.Principal,
                "@CreditAmount", this.CreditAmount,
                "@Interest", this.Interest,
                "@Mpesafee", this.Mpesafee,
                "@duedate", this.Duedate,
                "@empfirstname", this.Empfirstname,
                "@trnature", this.Trnature,
                "@empsurname", this.Empsurname,
                "@trdetails", this.Trdetails,
                "@payments", this.Payments,
                "@insurance", this.Insurance,
                "@schemename", this.Schemename
              );
            err = error;
            if (err == "")
            {
                if(rd.Read ())
                {
                    id = int.Parse(rd["CopyId"].ToString());
                }
                try { rd.Close(); }
                catch {; }
            }
                
            return id;

        }
		public int AddDecLoan(ref string error)
		{
			int id = 0;
			Link myLink = new Link();
			DbDataReader rd = myLink.GetDBResults(ref err, "sp_MigrateTotblloanCopy", "@PhoneNo", this.PhoneNo,
			"@Names", this.Names,
			"@Advance", this.Advance,
			"@LoanFees", this.LoanFees,
			"@TotalDebits", this.TotalDebits,
			"@Credit", this.Credit,
			"@Balance", this.Balance,
			"@Id", this.Id,
			"@Scheme", this.Scheme,
			"@SchemeId", this.SchemeId,
			"@IdNumber", this.IdNumber,
			"@staffNo", this.StaffNo,
			"@Rate", this.Rate,
			"@MemberId", this.MemberId );
			err = error;
			if (err == "")
			{
				if (rd.Read())
				{
					id = int.Parse(rd["CopyId"].ToString());
				}
				try { rd.Close(); }
				catch {; }
			}

			return id;

		}

	}
}
