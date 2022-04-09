using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;

namespace ReadExcel.Classes
{
    class Serials
    {
        private int _serialid = 0;
        
        public int Serialid { get { return _serialid; } set { _serialid = value; } }
        string err = "";
        public int GetSerialId(string createdBy)
        {
            int serialid = 0;
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "sp_NewSerial", "@createdby", createdBy);
            if (err == "")
            {
                if (rd.Read())
                {
                    serialid =  int.Parse(rd["SerialId"].ToString());


                }
                try { rd.Close(); }
                catch {; }
            }
            return serialid;

        }

    }
}
