using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;


namespace ReadExcel.Classes
{
    class LoanRepayments
    {
        private int _repaymentId = 0;
        private int _memberid = 0;
        private int _loanId = 0;
        private int _serialId = 0;
        private int _repaymentNo = 0;
        private DateTime _paymentDate = DateTime.Today;
        private double _paymentAmount = 0;
        private double _principal = 0;
        private double _interest = 0;
        private string _payMode = "";
        private string _voucherNo = "";
        private double _newInterest = 0;
        private double _loanAmount = 0;
        private double _otherCharges = 0;
        private double _extraInterest = 0;
        private bool _affectsDR = false;
        private double _loanBalance = 0;
        private double _monthOpeningBal = 0;
        private double _loanPenalty = 0;
        private int _transPer = 0;
        private int _transYear = 0;
        private int _bankId = 0;
        private bool _isReversal = false;
        private int _origPaymentNo = 0;
        private DateTime _systemEntryDate = DateTime.Today;
        private string _receiptNo = "";
        private string _createdBy = "";
        private DateTime _createdOn = DateTime.Today;
        private string _modifiedBy = "";
        private DateTime _modifiedOn = DateTime.Today;
        private bool _repaidbyGuarantors = false;
        private string _machineName = "";
        private string _bankAccount = "";
        private string _chequeNo = "";
        private string _paidBy = "";
        private string _remarks = "";
        private bool _isInterestReceivableTrx = false;
        private bool _isEOMPenalty = false;
        private int _previousRepaymentId = 0;
        private string _sourceBranch = "";
        private bool _isLienTrx = false;
        private double _distributedPrincipal = 0;
        private double _distributedInterest = 0;
        private double _distributedPenalty = 0;
        private int _productTypeid = 0;
        private int _productid = 0;
        private string _gLDR = "";
        private int _receiptId = 0;
        private string _memberno = "";

        public int ProductTypeId { get { return _productTypeid; } set { _productTypeid = value; } }
        public int ProductId { get { return _productid; } set { _productid = value; } }
        public int RepaymentId { get { return _repaymentId; } set { _repaymentId = value; } }
        public int Memberid { get { return _memberid; } set { _memberid = value; } }
        public int LoanId { get { return _loanId; } set { _loanId = value; } }
        public int SerialId { get { return _serialId; } set { _serialId = value; } }
        public int RepaymentNo { get { return _repaymentNo; } set { _repaymentNo = value; } }
        public DateTime PaymentDate { get { return _paymentDate; } set { _paymentDate = value; } }
        public double PaymentAmount { get { return _paymentAmount; } set { _paymentAmount = value; } }
        public double Principal { get { return _principal; } set { _principal = value; } }
        public double Interest { get { return _interest; } set { _interest = value; } }
        public string PayMode { get { return _payMode; } set { _payMode = value; } }
        public string VoucherNo { get { return _voucherNo; } set { _voucherNo = value; } }
        public double NewInterest { get { return _newInterest; } set { _newInterest = value; } }
        public double LoanAmount { get { return _loanAmount; } set { _loanAmount = value; } }
        public double OtherCharges { get { return _otherCharges; } set { _otherCharges = value; } }
        public double ExtraInterest { get { return _extraInterest; } set { _extraInterest = value; } }
        public bool AffectsDR { get { return _affectsDR; } set { _affectsDR = value; } }
        public double LoanBalance { get { return _loanBalance; } set { _loanBalance = value; } }
        public double MonthOpeningBal { get { return _monthOpeningBal; } set { _monthOpeningBal = value; } }
        public double LoanPenalty { get { return _loanPenalty; } set { _loanPenalty = value; } }
        public int TransPer { get { return _transPer; } set { _transPer = value; } }
        public int TransYear { get { return _transYear; } set { _transYear = value; } }
        public int BankId { get { return _bankId; } set { _bankId = value; } }
        public bool IsReversal { get { return _isReversal; } set { _isReversal = value; } }
        public int OrigPaymentNo { get { return _origPaymentNo; } set { _origPaymentNo = value; } }
        public DateTime SystemEntryDate { get { return _systemEntryDate; } set { _systemEntryDate = value; } }
        public string ReceiptNo { get { return _receiptNo; } set { _receiptNo = value; } }
        public string CreatedBy { get { return _createdBy; } set { _createdBy = value; } }
        public DateTime CreatedOn { get { return _createdOn; } set { _createdOn = value; } }
        public string ModifiedBy { get { return _modifiedBy; } set { _modifiedBy = value; } }
        public DateTime ModifiedOn { get { return _modifiedOn; } set { _modifiedOn = value; } }
        public bool RepaidbyGuarantors { get { return _repaidbyGuarantors; } set { _repaidbyGuarantors = value; } }
        public string MachineName { get { return _machineName; } set { _machineName = value; } }
        public string BankAccount { get { return _bankAccount; } set { _bankAccount = value; } }
        public string ChequeNo { get { return _chequeNo; } set { _chequeNo = value; } }
        public string PaidBy { get { return _paidBy; } set { _paidBy = value; } }
        public string Remarks { get { return _remarks; } set { _remarks = value; } }
        public bool IsInterestReceivableTrx { get { return _isInterestReceivableTrx; } set { _isInterestReceivableTrx = value; } }
        public bool IsEOMPenalty { get { return _isEOMPenalty; } set { _isEOMPenalty = value; } }
        public int PreviousRepaymentId { get { return _previousRepaymentId; } set { _previousRepaymentId = value; } }
        public string SourceBranch { get { return _sourceBranch; } set { _sourceBranch = value; } }
        public bool IsLienTrx { get { return _isLienTrx; } set { _isLienTrx = value; } }
        public double DistributedPrincipal { get { return _distributedPrincipal; } set { _distributedPrincipal = value; } }
        public double DistributedInterest { get { return _distributedInterest; } set { _distributedInterest = value; } }
        public double DistributedPenalty { get { return _distributedPenalty; } set { _distributedPenalty = value; } }
        public string GLDR { get { return _gLDR; } set { _gLDR = value; } }
        public int ReceiptId { get { return _receiptId; } set { _receiptId = value; } }
        public string MemberNo { get { return _memberno; } set { _memberno = value; } }
        string err = "";
        public int PostLoanRepayments(ref string error)
        {
            int id = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_importTransactions",
                    "@productCategoryId", this.ProductTypeId,
                    "@productTypeId", this.ProductId,
                    "@mcode", this.MemberNo,
                    "@DebitGL", this.GLDR,
                    "@paymode", this.PayMode ,
                    "@chequenumber", this.ChequeNo,
                    "@transamount", this.PaymentAmount,
                    "@transDate", this.PaymentDate ,
                    "@serialId", this.SerialId,
                    "@receiptid", this.ReceiptId,
                    "@createdby", "Test"

                   );



            if (err == "")
            {
                if (rd.Read())
                {
                    id = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch (Exception ex)
                {
                    error = ex.Message.ToString();

                }
            }

            error = err;
            return id;

        }
    }
}
