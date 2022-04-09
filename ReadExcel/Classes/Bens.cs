using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class Bens
    {
        private int _benid = 0;
        private int _memberid = 0;
        private string _relationship = "";
        private string _benname = "";
        private string _bencode = "";
        private string _telephone = "";
        private string _town = "";
        private double _benvalue = 0;
        private DateTime _createdOn = DateTime.Today;
        private string _createdBy = "";
        private DateTime _modifiedOn = DateTime.Today;
        private string _modifiedBy = "";
        private DateTime _dateOfBirth = DateTime.Today;

        public int Benid { get { return _benid; } set { _benid = value; } }
        public int Memberid { get { return _memberid; } set { _memberid = value; } }
        public string Relationship { get { return _relationship; } set { _relationship = value; } }
        public string Benname { get { return _benname; } set { _benname = value; } }
        public string Bencode { get { return _bencode; } set { _bencode = value; } }
        public string Telephone { get { return _telephone; } set { _telephone = value; } }
        public string Town { get { return _town; } set { _town = value; } }
        public double Benvalue { get { return _benvalue; } set { _benvalue = value; } }
        public DateTime CreatedOn { get { return _createdOn; } set { _createdOn = value; } }
        public string CreatedBy { get { return _createdBy; } set { _createdBy = value; } }
        public DateTime ModifiedOn { get { return _modifiedOn; } set { _modifiedOn = value; } }
        public string ModifiedBy { get { return _modifiedBy; } set { _modifiedBy = value; } }
        public DateTime DateOfBirth { get { return _dateOfBirth; } set { _dateOfBirth = value; } }


        string err = "";
        public int AddEditBen(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_addeditben",
                    "@Benid", this.Benid,
                    "@memberid", this.Memberid,
                    "@relationship", this.Relationship,
                    "@Benname", this.Benname,
                    "@Bencode", this.Bencode,
                    "@telephone", this.Telephone,
                    "@town", this.Town,
                   
                    "@CreatedBy", this.CreatedBy,
                    "@DateOfBirth", this.DateOfBirth

                   
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
