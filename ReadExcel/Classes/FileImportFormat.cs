using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class FileImportFormat
    {
        private int _fileFormatId = 0;
        private int _productId = 0;
        private string _productName = "";
        private int _position = 0;
        private string _fileimportname = "";
        private int _productMigrationId = 0;
        private int _importFileNameId = 0;
        private bool _isLoan = false;
        public int FileFormatId { get { return _fileFormatId; } set { _fileFormatId = value; } }
        public int ProductId { get { return _productId; } set { _productId = value; } }
        public string ProductName { get { return _productName; } set { _productName = value; } }
        public int Position { get { return _position; } set { _position = value; } }
        public string FileImportName { get { return _fileimportname; } set { _fileimportname = value; } }
        public int ProductMigrationId { get { return _productMigrationId; } set { _productMigrationId = value; } }
        public bool IsLoan { get { return _isLoan; } set { _isLoan = value; } }
        public int ImportFileNameId { get { return _importFileNameId; } set { _importFileNameId = value; } }
        ImportFileNames oImportFileName = new ImportFileNames ();

        string err = "";
        public ArrayList GetFileImportFormatsByImportFileName(int importFileNameId)
        {
            ArrayList myList = new ArrayList();
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "proc_GetAllFileImportFormatsByImportFileName", "@ImportFileNameId",importFileNameId);
            if (err == "")
            {
                while (rd.Read())
                {
                    FileImportFormat obj = new Classes.FileImportFormat();
                    if (!String.IsNullOrEmpty(rd["FileFormatId"].ToString())) obj.FileFormatId = int.Parse(rd["FileFormatId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductId"].ToString())) obj.ProductId = int.Parse(rd["ProductId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductName"].ToString())) obj.ProductName = rd["ProductName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Position"].ToString())) obj.Position = int.Parse(rd["Position"].ToString());
                   if (!String.IsNullOrEmpty(rd["ImportFileNameId"].ToString())) obj.ImportFileNameId = int.Parse(rd["ImportFileNameId"].ToString());
                    if (!String.IsNullOrEmpty(rd["IsLoan"].ToString())) obj.IsLoan = bool.Parse(rd["IsLoan"].ToString());
                    if(obj.ImportFileNameId >0)
                    {
                        ImportFileNames myImportName = oImportFileName.GetImportFileName(obj.ImportFileNameId);
                        if(myImportName !=null)
                        {
                            obj.FileImportName = myImportName.ImportFileName;
                        }
                    }
                    myList.Add(obj);
                }
                try { rd.Close(); }
                catch {; }
            }
            return myList;

        }
        public ArrayList GetFileImportFormats()
        {
            ArrayList myList = new ArrayList();
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "proc_getAllFileFormat");
            if (err == "")
            {
                while (rd.Read())
                {
                    FileImportFormat obj = new Classes.FileImportFormat();
                    if (!String.IsNullOrEmpty(rd["FileFormatId"].ToString())) obj.FileFormatId = int.Parse(rd["FileFormatId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductId"].ToString())) obj.ProductId = int.Parse(rd["ProductId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductName"].ToString())) obj.ProductName = rd["ProductName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Position"].ToString())) obj.Position = int.Parse(rd["Position"].ToString());
                    if (!String.IsNullOrEmpty(rd["ImportFileNameId"].ToString())) obj.ImportFileNameId = int.Parse(rd["ImportFileNameId"].ToString());
                    if (!String.IsNullOrEmpty(rd["IsLoan"].ToString())) obj.IsLoan = bool.Parse(rd["IsLoan"].ToString());
                    if (obj.ImportFileNameId > 0)
                    {
                        ImportFileNames myImportName = oImportFileName.GetImportFileName(obj.ImportFileNameId);
                        if (myImportName != null)
                        {
                            obj.FileImportName = myImportName.ImportFileName;
                        }
                    }
                    myList.Add(obj);
                }
                try { rd.Close(); }
                catch {; }
            }
            return myList;

        }
        public FileImportFormat GetFileImportFormat( int fileImportFormat)
        {
            FileImportFormat obj = null;
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "proc_getAllFileFormat", "@FileFormatId", fileImportFormat);
            if (err == "")
            {
                while (rd.Read())
                {
                   obj = new Classes.FileImportFormat();
                    if (!String.IsNullOrEmpty(rd["FileFormatId"].ToString())) obj.FileFormatId = int.Parse(rd["FileFormatId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductId"].ToString())) obj.ProductId = int.Parse(rd["ProductId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductName"].ToString())) obj.ProductName = rd["ProductName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Position"].ToString())) obj.Position = int.Parse(rd["Position"].ToString());
                    if (!String.IsNullOrEmpty(rd["ImportFileNameId"].ToString())) obj.ImportFileNameId = int.Parse(rd["ImportFileNameId"].ToString());
                    if (!String.IsNullOrEmpty(rd["IsLoan"].ToString())) obj.IsLoan = bool.Parse(rd["IsLoan"].ToString());
                    if (obj.ImportFileNameId > 0)
                    {
                        ImportFileNames myImportName = oImportFileName.GetImportFileName(obj.ImportFileNameId);
                        if (myImportName != null)
                        {
                            obj.FileImportName = myImportName.ImportFileName;
                        }
                    }
                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;

        }

        public FileImportFormat proc_GetFileFormatDetails(int importFileNameId,int position)
        {
            FileImportFormat obj = null;
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "proc_GetFileFormatDetails", "@ImportFileNameId", importFileNameId, "@Position", position );
            if (err == "")
            {
                while (rd.Read())
                {
                    obj = new Classes.FileImportFormat();
                    if (!String.IsNullOrEmpty(rd["FileFormatId"].ToString())) obj.FileFormatId = int.Parse(rd["FileFormatId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductId"].ToString())) obj.ProductId = int.Parse(rd["ProductId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductName"].ToString())) obj.ProductName = rd["ProductName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Position"].ToString())) obj.Position = int.Parse(rd["Position"].ToString());
                    if (!String.IsNullOrEmpty(rd["ImportFileNameId"].ToString())) obj.ImportFileNameId = int.Parse(rd["ImportFileNameId"].ToString());
                    if (!String.IsNullOrEmpty(rd["IsLoan"].ToString())) obj.IsLoan = bool.Parse(rd["IsLoan"].ToString());
                    if (obj.ImportFileNameId > 0)
                    {
                        ImportFileNames myImportName = oImportFileName.GetImportFileName(obj.ImportFileNameId);
                        if (myImportName != null)
                        {
                            obj.FileImportName = myImportName.ImportFileName;
                        }
                    }
                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;

        }
        public int AddFileImportFormat(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_AddFileImportFormat",
                   "@FileFormatId", this.FileFormatId,
                   "@IsLoan", this.IsLoan,
                    "@ProductId", this.ProductId,
                    "@ProductName", this.ProductName,
                    "@Position", this.Position,
                    "@CreatedBy", "Test",
                    "@Machinename", Environment.MachineName,
                    "@ImportFileNameId", this.ImportFileNameId
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
