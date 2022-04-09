using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class ImportFileNames
    {
        private int _importFileNameId = 0;
        private string _importFileName = "";
        private string _remarks = "";

        public int ImportFileNameId { get { return _importFileNameId; } set { _importFileNameId = value; } }
        public string ImportFileName { get { return _importFileName; } set { _importFileName = value; } }
        public string Remarks { get { return _remarks; } set { _remarks = value; } }
        string err = "";
        public ArrayList GetImportFileNames()
        {
            ArrayList myList = new ArrayList();
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "proc_getAllTransactionsImportFileNames");
            if (err == "")
            {
                while (rd.Read())
                {
                    ImportFileNames  obj = new Classes.ImportFileNames ();
                    if (!String.IsNullOrEmpty(rd["ImportFileNameId"].ToString())) obj.ImportFileNameId = int.Parse(rd["ImportFileNameId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ImportFileName"].ToString())) obj.ImportFileName = rd["ImportFileName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Remarks"].ToString())) obj.Remarks = rd["Remarks"].ToString();

                    myList.Add(obj);
                }
                try { rd.Close(); }
                catch {; }
            }
            return myList;

        }
        public ImportFileNames GetImportFileName(int importFileNameId)
        {
            Classes.ImportFileNames obj = null;
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "proc_getTransactionsImportFileName", "@ImportFileNameId", importFileNameId );
            if (err == "")
            {
                while (rd.Read())
                {
                    obj = new Classes.ImportFileNames();
                    if (!String.IsNullOrEmpty(rd["ImportFileNameId"].ToString())) obj.ImportFileNameId = int.Parse(rd["ImportFileNameId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ImportFileName"].ToString())) obj.ImportFileName = rd["ImportFileName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Remarks"].ToString())) obj.Remarks = rd["Remarks"].ToString();

                    
                }
                try { rd.Close(); }
                catch {; }
            }
            return obj ;

        }
        public int AddTransationImportFileName(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_AddEditImportFileNames",
                   "@ImportFileNameId", this.ImportFileNameId,
                    "@ImportFileName", this.ImportFileName,
                    "@Remarks", this.Remarks,
                    "@CreatedBy", "Test",
                    "@Machinename", Environment.MachineName
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
