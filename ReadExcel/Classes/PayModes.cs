using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class PayModes
    {
        private int _paymentModeId = 0;
        private string _paymentModeName = "";
        private string _description = "";
      
       
        private bool _allowBackDatedTransactions = false;
        private bool _isInhouseClearingPaymode = false;


        public int PaymentModeId { get { return _paymentModeId; } set { _paymentModeId = value; } }
        public string PaymentModeName { get { return _paymentModeName; } set { _paymentModeName = value; } }
        public string Description { get { return _description; } set { _description = value; } }
       
        public bool AllowBackDatedTransactions { get { return _allowBackDatedTransactions; } set { _allowBackDatedTransactions = value; } }
        public bool IsInhouseClearingPaymode { get { return _isInhouseClearingPaymode; } set { _isInhouseClearingPaymode = value; } }

        string err = "";
        public ArrayList GetPayModes()
        {
            ArrayList myList = new ArrayList();
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_GetAllPaymentModes");
            if (err == "")
            {
                while (rd.Read())
                {
                    PayModes obj = new Classes.PayModes();
                    if (!String.IsNullOrEmpty(rd["PaymentModeId"].ToString())) obj.PaymentModeId = int.Parse(rd["PaymentModeId"].ToString());
                    if (!String.IsNullOrEmpty(rd["PaymentModeName"].ToString())) obj.PaymentModeName = rd["PaymentModeName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Description"].ToString())) obj.Description = rd["Description"].ToString();
 
                    if (!String.IsNullOrEmpty(rd["AllowBackDatedTransactions"].ToString())) obj.AllowBackDatedTransactions = bool.Parse(rd["AllowBackDatedTransactions"].ToString());
                    if (!String.IsNullOrEmpty(rd["IsInhouseClearingPaymode"].ToString())) obj.IsInhouseClearingPaymode = bool.Parse(rd["IsInhouseClearingPaymode"].ToString());

                    myList.Add(obj);
                }
                try { rd.Close(); }
                catch {; }
            }
            return myList;
        }
        public PayModes GetPayMode(int PayModesId)
        {
            PayModes obj = null;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_getPaymentMode", "@PaymentModeId", PayModesId);
            if (err == "")
            {
                if (rd.Read())
                {
                    obj = new Classes.PayModes();
                    if (!String.IsNullOrEmpty(rd["PaymentModeId"].ToString())) obj.PaymentModeId = int.Parse(rd["PaymentModeId"].ToString());
                    if (!String.IsNullOrEmpty(rd["PaymentModeName"].ToString())) obj.PaymentModeName = rd["PaymentModeName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Description"].ToString())) obj.Description = rd["Description"].ToString();

                    if (!String.IsNullOrEmpty(rd["AllowBackDatedTransactions"].ToString())) obj.AllowBackDatedTransactions = bool.Parse(rd["AllowBackDatedTransactions"].ToString());
                    if (!String.IsNullOrEmpty(rd["IsInhouseClearingPaymode"].ToString())) obj.IsInhouseClearingPaymode = bool.Parse(rd["IsInhouseClearingPaymode"].ToString());
                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;
        }
    }
}
