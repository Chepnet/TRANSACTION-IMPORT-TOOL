using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel.Classes
{
    class GroupOfficial
    {
        private int _groupOfficialId = 0;
        private int _groupId = 0;
        private int _groupTitleId = 0;
        private string _officialName = "";
        private string _officialIDNo = "";
        private string _officialPhoneNo = "";
        private string _officialContact = "";
        private bool _isActive = false;
        private string _remarks = "";
        private string _employername = "";
        private string _title = "";
        private int _memberId = 0;
        private string _membername = "";
        public int GroupOfficialId { get { return _groupOfficialId; } set { _groupOfficialId = value; } }
        public int GroupId { get { return _groupId; } set { _groupId = value; } }
        public int GroupTitleId { get { return _groupTitleId; } set { _groupTitleId = value; } }
        public string OfficialName { get { return _officialName; } set { _officialName = value; } }
        public string OfficialIDNo { get { return _officialIDNo; } set { _officialIDNo = value; } }
        public string OfficialPhoneNo { get { return _officialPhoneNo; } set { _officialPhoneNo = value; } }
        public string OfficialContact { get { return _officialContact; } set { _officialContact = value; } }
        public bool IsActive { get { return _isActive; } set { _isActive = value; } }
        public string Remarks { get { return _remarks; } set { _remarks = value; } }
        public string Employername { get { return _employername; } set { _employername = value; } }
        public string GroupTitleName { get { return _title; } set { _title = value; } }
        public int MemberId { get { return _memberId; } set { _memberId = value; } }
        public string Membername { get { return _membername; } set { _membername = value; } }
        string err = "";
        
        Member omember = new Member();
        Member onewmember = null;
                public ArrayList GetGroupOfficials()
        {
            ArrayList MyList = new ArrayList();
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "pro_getAllGroupOfficials");
            if (err == "")
            {
                while (rd.Read())
                {
                    GroupOfficial obj = new Classes.GroupOfficial();
                    if (!String.IsNullOrEmpty(rd["GroupOfficialId"].ToString())) obj.GroupOfficialId = int.Parse(rd["GroupOfficialId"].ToString());
                    if (!String.IsNullOrEmpty(rd["GroupId"].ToString())) obj.GroupId = int.Parse(rd["GroupId"].ToString());
                    if (!String.IsNullOrEmpty(rd["MemberId"].ToString())) obj.MemberId = int.Parse(rd["MemberId"].ToString());

                    if (!String.IsNullOrEmpty(rd["GroupTitleId"].ToString())) obj.GroupTitleId = int.Parse(rd["GroupTitleId"].ToString());
                    if (!String.IsNullOrEmpty(rd["OfficialName"].ToString())) obj.OfficialName = rd["OfficialName"].ToString();
                    if (!String.IsNullOrEmpty(rd["OfficialIDNo"].ToString())) obj.OfficialIDNo = rd["OfficialIDNo"].ToString();
                    if (!String.IsNullOrEmpty(rd["OfficialPhoneNo"].ToString())) obj.OfficialPhoneNo = rd["OfficialPhoneNo"].ToString();
                    if (!String.IsNullOrEmpty(rd["OfficialContact"].ToString())) obj.OfficialContact = rd["OfficialContact"].ToString();
                    if (!String.IsNullOrEmpty(rd["IsActive"].ToString())) obj.IsActive = bool.Parse(rd["IsActive"].ToString());
                    if (!String.IsNullOrEmpty(rd["Remarks"].ToString())) obj.Remarks = rd["Remarks"].ToString();
                  
                 
                   
                    MyList.Add(obj);
                }
                try { rd.Close(); }
                catch {; }
            }
            return MyList;

        }
        public GroupOfficial GetGroupOfficial(int GroupOfficialId)
        {
            GroupOfficial obj = null;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "pro_getGroupOfficial", "@GroupOfficialId", GroupOfficialId);
            if (err == "")
            {
                if (rd.Read())
                {
                    obj = new Classes.GroupOfficial();
                    if (!String.IsNullOrEmpty(rd["GroupOfficialId"].ToString())) obj.GroupOfficialId = int.Parse(rd["GroupOfficialId"].ToString());
                    if (!String.IsNullOrEmpty(rd["GroupId"].ToString())) obj.GroupId = int.Parse(rd["GroupId"].ToString());
                    if (!String.IsNullOrEmpty(rd["MemberId"].ToString())) obj.MemberId = int.Parse(rd["MemberId"].ToString());
                    if (!String.IsNullOrEmpty(rd["GroupTitleId"].ToString())) obj.GroupTitleId = int.Parse(rd["GroupTitleId"].ToString());
                    if (!String.IsNullOrEmpty(rd["OfficialName"].ToString())) obj.OfficialName = rd["OfficialName"].ToString();
                    if (!String.IsNullOrEmpty(rd["OfficialIDNo"].ToString())) obj.OfficialIDNo = rd["OfficialIDNo"].ToString();
                    if (!String.IsNullOrEmpty(rd["OfficialPhoneNo"].ToString())) obj.OfficialPhoneNo = rd["OfficialPhoneNo"].ToString();
                    if (!String.IsNullOrEmpty(rd["OfficialContact"].ToString())) obj.OfficialContact = rd["OfficialContact"].ToString();
                    if (!String.IsNullOrEmpty(rd["IsActive"].ToString())) obj.IsActive = bool.Parse(rd["IsActive"].ToString());
                    if (!String.IsNullOrEmpty(rd["Remarks"].ToString())) obj.Remarks = rd["Remarks"].ToString();

                
                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;

        }
        public int AddEditGroupOfficial(bool delete, ref string error)
        {
            int id = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "pro_AddEditGroupOfficial", "@GroupOfficialId", this.GroupOfficialId,
                                "@GroupId", this.GroupId,
                                 "@MemberId", this.MemberId,
                                "@GroupTitleId", this.GroupTitleId,
                                "@OfficialName", this.OfficialName,
                                "@OfficialIDNo", this.OfficialIDNo,
                                "@OfficialPhoneNo", this.OfficialPhoneNo,
                                "@OfficialContact", this.OfficialContact,
                                "@IsActive", this.IsActive,
                                "@Remarks", this.Remarks,
                                "@MachineName", Environment.MachineName,
                                "@CreatedBy", "Test",
                                "@delete", delete);

            if (err == "")
            {
                if (rd.Read())
                {
                    id = int.Parse(rd["Id"].ToString());
                }
                try { rd.Close(); }
                catch {; }
            }
            err = error;
            return id;
        }
    }
}
