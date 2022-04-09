using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
namespace ReadExcel.Classes
{
    class Banks
    {
        private int _bankid = 0;
        private string _bankcode = "";
        private string _bankrefcode = "";
        private int _groupid = 0;
        private string _bankname = "";
        private string _bankGLCode = "";
        private string _remarks = "";
        private bool _defaultMPESAac = false;
        private bool _isChequeClearingAccount = false;
        public int Bankid { get { return _bankid; } set { _bankid = value; } }
        public string Bankcode { get { return _bankcode; } set { _bankcode = value; } }
        public string Bankrefcode { get { return _bankrefcode; } set { _bankrefcode = value; } }
        public int Groupid { get { return _groupid; } set { _groupid = value; } }
        public string Bankname { get { return _bankname; } set { _bankname = value; } }
        public string BankGLCode { get { return _bankGLCode; } set { _bankGLCode = value; } }
        public string Remarks { get { return _remarks; } set { _remarks = value; } }
        public bool DefaultMPESAac { get { return _defaultMPESAac; } set { _defaultMPESAac = value; } }
        public bool IsChequeClearingAccount { get { return _isChequeClearingAccount; } set { _isChequeClearingAccount = value; } }

        string err = "";
        public ArrayList GetBanks()
        {
            ArrayList myList = new ArrayList();
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_GetAllBanks");
            if (err == "")
            {
                while (rd.Read())
                {
                    Banks obj = new Classes.Banks();
                    if (!String.IsNullOrEmpty(rd["bankid"].ToString())) obj.Bankid = int.Parse(rd["bankid"].ToString());
                    if (!String.IsNullOrEmpty(rd["bankcode"].ToString())) obj.Bankcode = rd["bankcode"].ToString();
                    if (!String.IsNullOrEmpty(rd["bankrefcode"].ToString())) obj.Bankrefcode = rd["bankrefcode"].ToString();
                    if (!String.IsNullOrEmpty(rd["groupid"].ToString())) obj.Groupid = int.Parse(rd["groupid"].ToString());
                    if (!String.IsNullOrEmpty(rd["bankname"].ToString())) obj.Bankname = rd["bankname"].ToString();
                    if (!String.IsNullOrEmpty(rd["BankGLCode"].ToString())) obj.BankGLCode = rd["BankGLCode"].ToString();
                    if (!String.IsNullOrEmpty(rd["Remarks"].ToString())) obj.Remarks = rd["Remarks"].ToString();
                    if (!String.IsNullOrEmpty(rd["DefaultMPESAac"].ToString())) obj.DefaultMPESAac = bool.Parse(rd["DefaultMPESAac"].ToString());
                  
                    if (!String.IsNullOrEmpty(rd["IsChequeClearingAccount"].ToString())) obj.IsChequeClearingAccount = bool.Parse(rd["IsChequeClearingAccount"].ToString());

                    myList.Add(obj);
                }
                try { rd.Close(); }
                catch {; }
            }
            return myList;
        }
        public Banks GetBank(int bankId)
        {
            Classes.Banks obj = null;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_getBank", "@BankId",bankId);
            if (err == "")
            {
                if (rd.Read())
                {
                    obj = new Classes.Banks();
                    if (!String.IsNullOrEmpty(rd["bankid"].ToString())) obj.Bankid = int.Parse(rd["bankid"].ToString());
                    if (!String.IsNullOrEmpty(rd["bankcode"].ToString())) obj.Bankcode = rd["bankcode"].ToString();
                    if (!String.IsNullOrEmpty(rd["bankrefcode"].ToString())) obj.Bankrefcode = rd["bankrefcode"].ToString();
                    if (!String.IsNullOrEmpty(rd["groupid"].ToString())) obj.Groupid = int.Parse(rd["groupid"].ToString());
                    if (!String.IsNullOrEmpty(rd["bankname"].ToString())) obj.Bankname = rd["bankname"].ToString();
                    if (!String.IsNullOrEmpty(rd["BankGLCode"].ToString())) obj.BankGLCode = rd["BankGLCode"].ToString();
                    if (!String.IsNullOrEmpty(rd["Remarks"].ToString())) obj.Remarks = rd["Remarks"].ToString();
                    if (!String.IsNullOrEmpty(rd["DefaultMPESAac"].ToString())) obj.DefaultMPESAac = bool.Parse(rd["DefaultMPESAac"].ToString());

                    if (!String.IsNullOrEmpty(rd["IsChequeClearingAccount"].ToString())) obj.IsChequeClearingAccount = bool.Parse(rd["IsChequeClearingAccount"].ToString());

                    
                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;
        }
    }
}
