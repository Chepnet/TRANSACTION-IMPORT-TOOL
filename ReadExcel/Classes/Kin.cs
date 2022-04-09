using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class Kin
    {
        private int _kinid = 0;
        private int _memberid = 0;
        private string _relationship = "";
        private string _kinname = "";
        private string _kincode = "";
        private string _kinaddress = "";
        private string _telephone = "";
        private string _town = "";
        private string _age = "";
        private DateTime _createdOn = DateTime.Today;
        private string _createdBy = "";
        private DateTime _modifiedOn = DateTime.Today;
        private string _modifiedBy = "";
        private DateTime _dateOfBirth = DateTime.Today;
        public int Kinid { get { return _kinid; } set { _kinid = value; } }
        public int Memberid { get { return _memberid; } set { _memberid = value; } }
        public string Relationship { get { return _relationship; } set { _relationship = value; } }
        public string Kinname { get { return _kinname; } set { _kinname = value; } }
        public string Kincode { get { return _kincode; } set { _kincode = value; } }
        public string Kinaddress { get { return _kinaddress; } set { _kinaddress = value; } }
        public string Telephone { get { return _telephone; } set { _telephone = value; } }
        public string Town { get { return _town; } set { _town = value; } }
        public string Age { get { return _age; } set { _age = value; } }
        public DateTime CreatedOn { get { return _createdOn; } set { _createdOn = value; } }
        public string CreatedBy { get { return _createdBy; } set { _createdBy = value; } }
        public DateTime ModifiedOn { get { return _modifiedOn; } set { _modifiedOn = value; } }
        public string ModifiedBy { get { return _modifiedBy; } set { _modifiedBy = value; } }
        public DateTime DateOfBirth { get { return _dateOfBirth; } set { _dateOfBirth = value; } }
        string err = "";
        public int AddEditKin(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_AddEditKins",
                    "@memberid", this.Memberid,
                    "@kinid", this.Kinid ,
                     "@relationship", this.Relationship,
                    "@kinname", this.Kinname ,
                    "@kincode", this.Kincode,
                    "@kinaddress", this.Kinaddress ,
                    "@telephone", this.Telephone,
                    "@town", this.Town,
                    "@CreatedBy", this.CreatedBy
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
