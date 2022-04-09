using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;
namespace ReadExcel.Classes
{
    class Receipt
    {
        private int _serialid = 0;
        private int _transId = 0;
        private DateTime _transDate = DateTime.Today;
        private int _receiptId = 0;
        private double _amount = 0;
        private string _receiptMode = "";
        public DateTime TransDate { get { return _transDate; } set { _transDate = value; } }
        public double Amount { get { return _amount; } set { _amount = value; } }
        public string ReceiptMode { get { return _receiptMode; } set { _receiptMode = value; } }
        public int Serialid { get { return _serialid; } set { _serialid = value; } }
        string err = "";
        public int GetReceiptId(string createdBy,double amount,string receiptmode,DateTime transdate,int serialid)
        {
            int receiptid = 0;
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "proc_newImportTrxnsReceipt ",
                "@createdby", createdBy,
                "@transamount",amount,
                "@paymode",receiptmode,
                "@transdate",transdate,
                "@serialid",serialid
                );
            if (err == "")
            {
                if (rd.Read())
                {
                    receiptid = int.Parse(rd["receiptid"].ToString());


                }
                try { rd.Close(); }
                catch {; }
            }
            return receiptid;

        }
    }
}
