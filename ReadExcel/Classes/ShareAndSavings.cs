using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;

namespace ReadExcel.Classes
{
    class ShareAndSavings
    {
        private int _transId = 0;
        private DateTime _transDate = DateTime.Today;
        private int _transPer = 0;
        private int _transYear = 0;
        private int _serialID = 0;
        private string _branchCode = "";
        private int _memberid = 0;
        private int _shareTypeId = 0;
        private int _memberShareId = 0;
        private double _openingBalance = 0;
        private bool _debitRaiseAffected = false;
        private DateTime _effectDate = DateTime.Today;
        private double _amount = 0;
        private string _receiptMode = "";
        private int _bankId = 0;
        private string _chequeNumber = "";
        private string _description = "";
        private string _remarks = "";
        private int _sourceMemberId = 0;
        private int _sourceMemberShareId = 0;
        private string _gLCR = "";
        private string _gLDR = "";
        private double _commission = 0;
        private bool _commPaidCash = false;
        private string _voucherNo = "";
        private bool _isReversal = false;
        private bool _reversed = false;
        private double _runningBalance = 0;
        private string _receiptNo = "";
        private string _createdBy = "";
        private DateTime _createdOn = DateTime.Today;
        private string _modifiedBy = "";
        private DateTime _modifiedOn = DateTime.Today;
        private int _prevId = 0;
        private string _machineName = "";
        private int _previousTransId = 0;
        private string _sourceBranch = "";
        private string _membername = "";
        private string _memberno = "";
        private string _sharename = "";
        private int _receiptId = 0;
        private int _productTypeid = 0;

        public int ProductTypeId { get { return _productTypeid; } set { _productTypeid = value; } }
        public int TransId { get { return _transId; } set { _transId = value; } }
        public DateTime TransDate { get { return _transDate; } set { _transDate = value; } }
        public int TransPer { get { return _transPer; } set { _transPer = value; } }
        public int TransYear { get { return _transYear; } set { _transYear = value; } }
        public int SerialID { get { return _serialID; } set { _serialID = value; } }
        public string BranchCode { get { return _branchCode; } set { _branchCode = value; } }
        public int Memberid { get { return _memberid; } set { _memberid = value; } }
        public int ShareTypeId { get { return _shareTypeId; } set { _shareTypeId = value; } }
        public int MemberShareId { get { return _memberShareId; } set { _memberShareId = value; } }
        public double OpeningBalance { get { return _openingBalance; } set { _openingBalance = value; } }
        public bool DebitRaiseAffected { get { return _debitRaiseAffected; } set { _debitRaiseAffected = value; } }
        public DateTime EffectDate { get { return _effectDate; } set { _effectDate = value; } }
        public double Amount { get { return _amount; } set { _amount = value; } }
        public string ReceiptMode { get { return _receiptMode; } set { _receiptMode = value; } }
        public int BankId { get { return _bankId; } set { _bankId = value; } }
        public string ChequeNumber { get { return _chequeNumber; } set { _chequeNumber = value; } }
        public string Description { get { return _description; } set { _description = value; } }
        public string Remarks { get { return _remarks; } set { _remarks = value; } }
        public int SourceMemberId { get { return _sourceMemberId; } set { _sourceMemberId = value; } }
        public int SourceMemberShareId { get { return _sourceMemberShareId; } set { _sourceMemberShareId = value; } }
        public string GLCR { get { return _gLCR; } set { _gLCR = value; } }
        public string GLDR { get { return _gLDR; } set { _gLDR = value; } }
        public double Commission { get { return _commission; } set { _commission = value; } }
        public bool CommPaidCash { get { return _commPaidCash; } set { _commPaidCash = value; } }
        public string VoucherNo { get { return _voucherNo; } set { _voucherNo = value; } }
        public bool IsReversal { get { return _isReversal; } set { _isReversal = value; } }
        public bool Reversed { get { return _reversed; } set { _reversed = value; } }
        public double RunningBalance { get { return _runningBalance; } set { _runningBalance = value; } }
        public string ReceiptNo { get { return _receiptNo; } set { _receiptNo = value; } }
        public string CreatedBy { get { return _createdBy; } set { _createdBy = value; } }
        public DateTime CreatedOn { get { return _createdOn; } set { _createdOn = value; } }
        public string ModifiedBy { get { return _modifiedBy; } set { _modifiedBy = value; } }
        public DateTime ModifiedOn { get { return _modifiedOn; } set { _modifiedOn = value; } }
        public int PrevId { get { return _prevId; } set { _prevId = value; } }
        public string MachineName { get { return _machineName; } set { _machineName = value; } }
        public int PreviousTransId { get { return _previousTransId; } set { _previousTransId = value; } }
        public string SourceBranch { get { return _sourceBranch; } set { _sourceBranch = value; } }
        public string MemberName { get { return _membername; } set { _membername = value; } }
        public string MemberNo { get { return _memberno; } set { _memberno = value; } }
        public string Sharename { get { return _sharename; } set { _sharename = value; } }
        public int ReceiptId { get { return _receiptId; } set { _receiptId = value; } }
        string err = "";
        public int PostSharesAndSavings(ref string error)
        {
            int id = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_importTransactions",
                    "@productCategoryId", this.ProductTypeId ,
                    "@productTypeId", this.ShareTypeId,
                    "@mcode",this.MemberNo,
                    "@DebitGL", this.GLDR,
                    "@paymode", this.ReceiptMode,
                    "@chequenumber", this.ChequeNumber,
                    "@transamount", this.Amount,
                    "@transDate", this.TransDate,
                    "@serialId", this.SerialID,
                    "@receiptid", this.ReceiptId,
                    "@createdby","Test"

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
