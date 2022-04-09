using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class MemberRegistration
    {
        private int _memberRegistrationFeeId = 0;
        private int _memberRegistrationId = 0;
        private double _amount = 0;
        private DateTime _datePaid = DateTime.Today;
        private string _gLDR = "";
        private string _gLCR = "";
        private string _modeOfPayment = "";
        private DateTime _createdOn = DateTime.Today;
        private string _createdBy = "";
        private string _supFlag = "";
        private DateTime _supervisedOn = DateTime.Today;
        private string _supervisedBy = "";
        private string _modifiedBy = "";
        private DateTime _modifiedOn = DateTime.Today;
        private string _documentNo = "";

        public int MemberRegistrationFeeId { get { return _memberRegistrationFeeId; } set { _memberRegistrationFeeId = value; } }
        public int MemberRegistrationId { get { return _memberRegistrationId; } set { _memberRegistrationId = value; } }
        public double Amount { get { return _amount; } set { _amount = value; } }
        public DateTime DatePaid { get { return _datePaid; } set { _datePaid = value; } }
        public string GLDR { get { return _gLDR; } set { _gLDR = value; } }
        public string GLCR { get { return _gLCR; } set { _gLCR = value; } }
        public string ModeOfPayment { get { return _modeOfPayment; } set { _modeOfPayment = value; } }
        public DateTime CreatedOn { get { return _createdOn; } set { _createdOn = value; } }
        public string CreatedBy { get { return _createdBy; } set { _createdBy = value; } }
        public string SupFlag { get { return _supFlag; } set { _supFlag = value; } }
        public DateTime SupervisedOn { get { return _supervisedOn; } set { _supervisedOn = value; } }
        public string SupervisedBy { get { return _supervisedBy; } set { _supervisedBy = value; } }
        public string ModifiedBy { get { return _modifiedBy; } set { _modifiedBy = value; } }
        public DateTime ModifiedOn { get { return _modifiedOn; } set { _modifiedOn = value; } }
        public string DocumentNo { get { return _documentNo; } set { _documentNo = value; } }
        string err = "";

        public int AddEditmemberregfee(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_AddEditmemberregistration",
                    "@MemberRegistrationFeeId", this.MemberRegistrationFeeId,
                    "@MemberRegistrationId", this.MemberRegistrationId,
                    "@Amount", this.Amount,
                    "@DatePaid", this.DatePaid

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
                    ;
                }
            }

            error = err;



            return id;


        }
    }
}
