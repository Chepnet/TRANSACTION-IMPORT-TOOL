using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class ShareTypes
    {
        private int _shareid = 0;
        private string _sharecode = "";
        private string _sharename = "";

        public int Shareid { get { return _shareid; } set { _shareid = value; } }
        public string Sharecode { get { return _sharecode; } set { _sharecode = value; } }
        public string Sharename { get { return _sharename; } set { _sharename = value; } }

        string err = "";
        public ArrayList GetShareTypes()
        {
            ArrayList myList = new ArrayList();
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "sp_GetAllShareTypes");
            if (err == "")
            {
                while (rd.Read())
                {
                    ShareTypes  obj = new Classes.ShareTypes();
                    if (!String.IsNullOrEmpty(rd["shareid"].ToString())) obj.Shareid = int.Parse(rd["shareid"].ToString());
                    if (!String.IsNullOrEmpty(rd["sharecode"].ToString())) obj.Sharecode = rd["sharecode"].ToString();
                    if (!String.IsNullOrEmpty(rd["sharename"].ToString())) obj.Sharename = rd["sharename"].ToString();


                    myList.Add(obj);
                }
                try { rd.Close(); }
                catch {; }
            }
            return myList;

        }
        public ShareTypes GetShareType(int shareTypeId)
        {
            ShareTypes obj = null;
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "sp_getShareType", "@ShareId", shareTypeId);
            if (err == "")
            {
                if (rd.Read())
                {
                    obj = new Classes.ShareTypes();
                    if (!String.IsNullOrEmpty(rd["shareid"].ToString())) obj.Shareid = int.Parse(rd["shareid"].ToString());
                    if (!String.IsNullOrEmpty(rd["sharecode"].ToString())) obj.Sharecode = rd["sharecode"].ToString();
                    if (!String.IsNullOrEmpty(rd["sharename"].ToString())) obj.Sharename = rd["sharename"].ToString();

                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;

        }

    }
}
