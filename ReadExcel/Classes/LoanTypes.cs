using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class LoanTypes
    {
        private int _loanTypeid = 0;
        private string _loanTypecode = "";
        private string _loanTypename = "";
        public int LoanTypeid { get { return _loanTypeid; } set { _loanTypeid = value; } }
        public string LoanTypecode { get { return _loanTypecode; } set { _loanTypecode = value; } }
        public string LoanTypename { get { return _loanTypename; } set { _loanTypename = value; } }

        string err = "";
        public ArrayList GetLoanTypes()
        {
            ArrayList myList = new ArrayList();
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "sp_GetAllLoanTypes");
            if (err == "")
            {
                while (rd.Read())
                {
                    LoanTypes obj = new Classes.LoanTypes();
                    if (!String.IsNullOrEmpty(rd["LoanTypeid"].ToString())) obj.LoanTypeid = int.Parse(rd["LoanTypeid"].ToString());
                    if (!String.IsNullOrEmpty(rd["LoanTypecode"].ToString())) obj.LoanTypecode = rd["LoanTypecode"].ToString();
                    if (!String.IsNullOrEmpty(rd["LoanTypename"].ToString())) obj.LoanTypename = rd["LoanTypename"].ToString();


                    myList.Add(obj);
                }
                try { rd.Close(); }
                catch {; }
            }
            return myList;

        }
        public LoanTypes  GetLoanType(int loanTypeId)
        {
            LoanTypes obj = null;
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "sp_getLoanType", " @LoanTypeId",loanTypeId);
            if (err == "")
            {
                if (rd.Read())
                {
                     obj = new Classes.LoanTypes();
                    if (!String.IsNullOrEmpty(rd["LoanTypeid"].ToString())) obj.LoanTypeid = int.Parse(rd["LoanTypeid"].ToString());
                    if (!String.IsNullOrEmpty(rd["LoanTypecode"].ToString())) obj.LoanTypecode = rd["LoanTypecode"].ToString();
                    if (!String.IsNullOrEmpty(rd["LoanTypename"].ToString())) obj.LoanTypename = rd["LoanTypename"].ToString();

                 
                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;

        }
    }
}
