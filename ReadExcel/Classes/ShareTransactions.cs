using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class ShareTransactions
    {
        private double _amount =0;
        private string _mcode = "";
        private string _mpayroll = "";
        private string _description = "";
        private int _transId = 0;
        private int _sharetypeId = 0;
        private string _transDate = "";
        public double Amount{ get { return _amount; } set { _amount = value; } }
        public string Mcode { get { return _mcode; } set { _mcode = value; } }
        public string Mpayroll { get { return _mpayroll; } set { _mpayroll = value; } }
        public string Description { get { return _description; } set { _description = value; } }
        public int TransId { get { return _transId; } set { _transId = value; } }
        public int SharetypeId { get { return _sharetypeId; } set { _sharetypeId = value; } }
        public string TransDate { get { return _transDate; } set { _transDate = value; } }

        public int AddEditBarabaraSharesandDeposit(ref string error)
        {
            int id = 0;
            string err = "";
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_migrateKRBSharesandDeposit",
                "@TransId", this.TransId,
                "@ShareTypeId", this.SharetypeId,
                    "@Mcode", this.Mcode,
                    "@mpayroll", "",
                     "@Amount", this.Amount,
                    "@Description", this.Description
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
