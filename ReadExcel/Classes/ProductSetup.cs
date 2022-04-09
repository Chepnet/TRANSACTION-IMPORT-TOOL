using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class ProductSetup
    {
        private int _productImpotId = 0;
        private string _productName = "";
        private string _description = "";
        private bool _isLoan = false;
        private int _productId = 0;
        private int _position = 1;
        public int ProductImportId { get { return _productImpotId; } set { _productImpotId = value; } }
        public string ProductName { get { return _productName; } set { _productName = value; } }
        public string Description { get { return _description; } set { _description = value; } }
        public bool IsLoan { get { return _isLoan; } set { _isLoan = value; } }
        public int ProductId { get { return _productId; } set { _productId = value; } }
        public int PositionId { get { return _position; } set { _position = value; } }
        string err = "";
        public ArrayList GetMigrationProducts()
        {
            ArrayList myList = new ArrayList();
            Link myLink = new Link();
            int position = 0;
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_getAllProductMigration");
            if (err == "")
            {
                while (rd.Read())
                {
                    ProductSetup obj = new Classes.ProductSetup();
                    if (!String.IsNullOrEmpty(rd["ProductMigrationId"].ToString())) obj.ProductImportId  = int.Parse(rd["ProductMigrationId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductName"].ToString())) obj.ProductName = rd["ProductName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Description"].ToString())) obj.Description = rd["Description"].ToString();
                    if (!String.IsNullOrEmpty(rd["IsLoan"].ToString())) obj.IsLoan = bool.Parse(rd["IsLoan"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductId"].ToString())) obj.ProductId = int.Parse(rd["ProductId"].ToString());
                    obj.PositionId = _position++;
                    myList.Add(obj);
                   

                }
                try { rd.Close(); }
                catch {; }
            }
            return myList;

        }
        public ProductSetup  GetMigrationProduct(int productMigrationId)
        {
            ProductSetup obj = null;
            Link myLink = new Link();

            DbDataReader rd = myLink.GetDBResults(ref err, "proc_getProductMigration", "@ProductMigrationId",productMigrationId);
            if (err == "")
            {
                if (rd.Read())
                {
                    obj = new Classes.ProductSetup();
                    if (!String.IsNullOrEmpty(rd["ProductMigrationId"].ToString())) obj.ProductImportId = int.Parse(rd["ProductMigrationId"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductName"].ToString())) obj.ProductName = rd["ProductName"].ToString();
                    if (!String.IsNullOrEmpty(rd["Description"].ToString())) obj.Description = rd["Description"].ToString();
                    if (!String.IsNullOrEmpty(rd["IsLoan"].ToString())) obj.IsLoan = bool.Parse(rd["IsLoan"].ToString());
                    if (!String.IsNullOrEmpty(rd["ProductId"].ToString())) obj.ProductId = int.Parse(rd["ProductId"].ToString());
                    


                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;

        }
        public int AddProduct(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_AddProductMigration",
                    "@ProductMigrationId", this.ProductImportId,
                    "@ProductId", this.ProductId,
                    "@ProductName", this.ProductName,
                    "@Description", this.Description,
                    "@IsLoan", this.IsLoan,
                    "@CreatedBy","Test",
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
