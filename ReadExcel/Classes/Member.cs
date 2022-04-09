using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.Common;
namespace ReadExcel.Classes
{
    class Member
    {
        private int _memberid = 0;
        private DateTime _mdate = DateTime.Now;
        private string _mpayroll = "";
        private string _mcode = "";
        private string _mtitle = "";
        private string _msurname = "";
        private string _mothername = "";
        private string _mfirstname = "";
        private string _mdob = "";
        private string _mgender = "";
        private string _mmarital = "";
        private string _mtel1 = "";
        private string _mtel2 = "";
        private string _mtel3 = "";
        private string _mcell = "";
        private string _mfax = "";
        private string _memail = "";
        private string _maddress = "";
        private string _mpostaladdress = "";
        private string _mcomments = "";
        private DateTime _mstartdate = DateTime.Today;
        private int _employerid = 0;
        private int _stationid = 0;
        private int _deptid = 0;
        private string _designationid = "";
        private string _menddate = "";
        private int _branchid = 0;
        private int _saccoofficial = 0;
        private string _mdesignationid = "";
        private DateTime _mwithdrawaldate = DateTime.Today;
        private string _statusid = "";
        private int _blocked = 0;
        private double _mgross = 0;
        private string _mnett = "";
        private string _gradeid = "";
        private double _regFee = 0;
        private DateTime _datePaid =DateTime.Now;
        private string _mstatus = "";
        private string _iDNO = "";
        private double _lOANBALatMIGRATION = 0;
        private string _fILENO = "";
        private string _createdBy = "";
        private DateTime _createdOn = DateTime.Today;
        private string _modifiedBy = "";
        private DateTime _modifiedOn = DateTime.Today;
        private int _levelofEducationID = 0;
        private int _businessSectorID = 0;
        private int _levelofIncomeID = 0;
        private int _documentTypeID = 0;
        private int _withdrawalReasonID = 0;
        private int _clientTypeid = 0;
        private string _companyId = "";
        private DateTime _supervisedOn = DateTime.Today;
        private string _sUpervisedBy = "";
        private int _previousMemberId = 0;
        private string _sourceBranch = "";
        private bool _new = false;

        private string _membername = "";
       
        public string MemberName { get { return _membername; } set { _membername = value; } }

        public int Memberid { get { return _memberid; } set { _memberid = value; } }
        public DateTime Mdate { get { return _mdate; } set { _mdate = value; } }
        public string Mpayroll { get { return _mpayroll; } set { _mpayroll = value; } }
        public string Mcode { get { return _mcode; } set { _mcode = value; } }
        public string Mtitle { get { return _mtitle; } set { _mtitle = value; } }
        public string Msurname { get { return _msurname; } set { _msurname = value; } }
        public string Mothername { get { return _mothername; } set { _mothername = value; } }
        public string Mfirstname { get { return _mfirstname; } set { _mfirstname = value; } }
        public string Mdob { get { return _mdob; } set { _mdob = value; } }
        public string Mgender { get { return _mgender; } set { _mgender = value; } }
        public string Mmarital { get { return _mmarital; } set { _mmarital = value; } }
        public string Mtel1 { get { return _mtel1; } set { _mtel1 = value; } }
        public string Mtel2 { get { return _mtel2; } set { _mtel2 = value; } }
        public string Mtel3 { get { return _mtel3; } set { _mtel3 = value; } }
        public string Mcell { get { return _mcell; } set { _mcell = value; } }
        public string Mfax { get { return _mfax; } set { _mfax = value; } }
        public string Memail { get { return _memail; } set { _memail = value; } }
        public string Maddress { get { return _maddress; } set { _maddress = value; } }
        public string Mpostaladdress { get { return _mpostaladdress; } set { _mpostaladdress = value; } }
        public string Mcomments { get { return _mcomments; } set { _mcomments = value; } }
        public DateTime Mstartdate { get { return _mstartdate; } set { _mstartdate = value; } }
        public int Employerid { get { return _employerid; } set { _employerid = value; } }
        public int Stationid { get { return _stationid; } set { _stationid = value; } }
        public int Deptid { get { return _deptid; } set { _deptid = value; } }
        public string Designationid { get { return _designationid; } set { _designationid = value; } }
        public string Menddate { get { return _menddate; } set { _menddate = value; } }
        public int Branchid { get { return _branchid; } set { _branchid = value; } }
        public int Saccoofficial { get { return _saccoofficial; } set { _saccoofficial = value; } }
        public string Mdesignationid { get { return _mdesignationid; } set { _mdesignationid = value; } }
        public DateTime Mwithdrawaldate { get { return _mwithdrawaldate; } set { _mwithdrawaldate = value; } }
        public string Statusid { get { return _statusid; } set { _statusid = value; } }
        public int Blocked { get { return _blocked; } set { _blocked = value; } }
        public double Mgross { get { return _mgross; } set { _mgross = value; } }
        public string Mnett { get { return _mnett; } set { _mnett = value; } }
        public string Gradeid { get { return _gradeid; } set { _gradeid = value; } }
        public double RegFee { get { return _regFee; } set { _regFee = value; } }
        public DateTime DatePaid { get { return _datePaid; } set { _datePaid = value; } }
        public string Mstatus { get { return _mstatus; } set { _mstatus = value; } }
        public string IDNO { get { return _iDNO; } set { _iDNO = value; } }
        public double LOANBALatMIGRATION { get { return _lOANBALatMIGRATION; } set { _lOANBALatMIGRATION = value; } }
        public string FILENO { get { return _fILENO; } set { _fILENO = value; } }
        public string CreatedBy { get { return _createdBy; } set { _createdBy = value; } }
        public DateTime CreatedOn { get { return _createdOn; } set { _createdOn = value; } }
        public string ModifiedBy { get { return _modifiedBy; } set { _modifiedBy = value; } }
        public DateTime ModifiedOn { get { return _modifiedOn; } set { _modifiedOn = value; } }
        public int LevelofEducationID { get { return _levelofEducationID; } set { _levelofEducationID = value; } }
        public int BusinessSectorID { get { return _businessSectorID; } set { _businessSectorID = value; } }
        public int LevelofIncomeID { get { return _levelofIncomeID; } set { _levelofIncomeID = value; } }
        public int DocumentTypeID { get { return _documentTypeID; } set { _documentTypeID = value; } }
        public int WithdrawalReasonID { get { return _withdrawalReasonID; } set { _withdrawalReasonID = value; } }
        public int ClientTypeid { get { return _clientTypeid; } set { _clientTypeid = value; } }
        public string CompanyId { get { return _companyId; } set { _companyId = value; } }
        public DateTime SupervisedOn { get { return _supervisedOn; } set { _supervisedOn = value; } }
        public string SUpervisedBy { get { return _sUpervisedBy; } set { _sUpervisedBy = value; } }
        public int PreviousMemberId { get { return _previousMemberId; } set { _previousMemberId = value; } }
        public string SourceBranch { get { return _sourceBranch; } set { _sourceBranch = value; } }
        public bool New { get { return _new; } set { _new = value; } }

        string err = "";
        public string FullName
        {
            get
            {
                return this.Mfirstname + " " + this.Mothername + " " + this.Msurname;
            }
        }
        public int AddEditMember(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_addeditMember",
                    "@memberid", this.Memberid,
                    "@mdate", "20210929",
                    "@msurname", this.Msurname,
                    "@mfirstname", this.Mfirstname,
                    "@mcell", this.Mcell,
                "@mgender", this.Mgender,
                "@mtel1", this.Mtel1,
                    "@IDNO", this.IDNO,
                    "@Datepaid",this.DatePaid,
                     "@Amount", this.RegFee
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
        public int AddEditBarabaraMember(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "sp_MigrateKrbMembers",
                     "@msurname", this.Msurname,
                    "@mfirstname", this.Mfirstname,
                    "@Mcode", this.Mcode,
                     "@Mpayroll", this.Mpayroll,
                     "@mothername", this.Mothername,
                        "@memail", this.Memail,
                        "@Blocked", this.Blocked,
                        "@employerid", this.Employerid,
                        "@Mdate", this.Mdate,
                        "@mcell", this.Mcell,
                         "@IdNo", this.IDNO,
                          "@mfax", this.Mfax
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
   
        public int MigrateExpense(ref string error)
        {
            int id = 0;

            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_Insertexpense",
                    "@memberid", this.Memberid,
                    "@Account", this.Employerid,
                    "@AccountName", this.Mcode,
                    "@New", this.New,
                    "@Amount",this.RegFee ,
                    "@TransDate", this.Mpayroll,

                     "@Parent", this.Stationid

                   );



            if (err == "")
            {
                if (rd.Read())
                {
                    //id = int.Parse(rd["Id"].ToString());
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
        public double  regfee { get; set; }

        string errmsg = "";
        public int getmember(ref string errm) {
            int done = 0;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref errmsg, "sp_getmembercode",
            "@mcode", this.Mcode 
            );
            errm = errmsg;
            if (errmsg == "")
            {
                if (rd.Read())
                {
                    done = int.Parse(rd["Memeberid"].ToString());
                }
                try { rd.Close(); rd.Dispose(); }
                catch { ;}
            }
            return (done);
        }
        public string getMemberNo(string MemberNo)
        {
            ////MemberNo = "";
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_CheckMemberNo", "@mcode", MemberNo
               );
            if (err == "")
            {
                if (rd.Read())
                {

                    this.Mcode = rd["MemberNo"].ToString();

                }
                try { rd.Close(); }
                catch {; }

            }
            return Mcode;
        }
        public Member GetMemberByCode(string  mcode)
        {
            Member obj = null;
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_getMemberbyCode", "@mcode", mcode);
            if (err == "")
            {
                if (rd.Read())
                {
                    obj = new Classes.Member();
                    if (!String.IsNullOrEmpty(rd["memberid"].ToString())) obj.Memberid = int.Parse(rd["memberid"].ToString());
                    if (!String.IsNullOrEmpty(rd["mdate"].ToString())) obj.Mdate = DateTime.Parse(rd["mdate"].ToString());
                    if (!String.IsNullOrEmpty(rd["mpayroll"].ToString())) obj.Mpayroll = rd["mpayroll"].ToString();
                    if (!String.IsNullOrEmpty(rd["mcode"].ToString())) obj.Mcode = rd["mcode"].ToString();
                    if (!String.IsNullOrEmpty(rd["mtitle"].ToString())) obj.Mtitle = rd["mtitle"].ToString();
                    if (!String.IsNullOrEmpty(rd["msurname"].ToString())) obj.Msurname = rd["msurname"].ToString();
                    if (!String.IsNullOrEmpty(rd["mothername"].ToString())) obj.Mothername = rd["mothername"].ToString();
                    if (!String.IsNullOrEmpty(rd["mfirstname"].ToString())) obj.Mfirstname = rd["mfirstname"].ToString();
                    //if (!String.IsNullOrEmpty(rd["mdob"].ToString())) obj.Mdob = DateTime.Parse(rd["mdob"].ToString());
                    if (!String.IsNullOrEmpty(rd["mgender"].ToString())) obj.Mgender = rd["mgender"].ToString();
                    if (!String.IsNullOrEmpty(rd["mmarital"].ToString())) obj.Mmarital = rd["mmarital"].ToString();
                    if (!String.IsNullOrEmpty(rd["mtel1"].ToString())) obj.Mtel1 = rd["mtel1"].ToString();
                    if (!String.IsNullOrEmpty(rd["mtel2"].ToString())) obj.Mtel2 = rd["mtel2"].ToString();
                    if (!String.IsNullOrEmpty(rd["mtel3"].ToString())) obj.Mtel3 = rd["mtel3"].ToString();
                    if (!String.IsNullOrEmpty(rd["mcell"].ToString())) obj.Mcell = rd["mcell"].ToString();
                    if (!String.IsNullOrEmpty(rd["mfax"].ToString())) obj.Mfax = rd["mfax"].ToString();
                    if (!String.IsNullOrEmpty(rd["memail"].ToString())) obj.Memail = rd["memail"].ToString();
                    if (!String.IsNullOrEmpty(rd["maddress"].ToString())) obj.Maddress = rd["maddress"].ToString();
                    if (!String.IsNullOrEmpty(rd["mpostaladdress"].ToString())) obj.Mpostaladdress = rd["mpostaladdress"].ToString();
                    if (!String.IsNullOrEmpty(rd["mcomments"].ToString())) obj.Mcomments = rd["mcomments"].ToString();
                    if (!String.IsNullOrEmpty(rd["mstartdate"].ToString())) obj.Mstartdate = DateTime.Parse(rd["mstartdate"].ToString());
                    if (!String.IsNullOrEmpty(rd["employerid"].ToString())) obj.Employerid = int.Parse(rd["employerid"].ToString());
                    if (!String.IsNullOrEmpty(rd["stationid"].ToString())) obj.Stationid = int.Parse(rd["stationid"].ToString());
                    if (!String.IsNullOrEmpty(rd["deptid"].ToString())) obj.Deptid = int.Parse(rd["deptid"].ToString());
                    if (!String.IsNullOrEmpty(rd["designationid"].ToString())) obj.Designationid = rd["designationid"].ToString();
                   // if (!String.IsNullOrEmpty(rd["menddate"].ToString())) obj.Menddate  = DateTime.Parse(rd["menddate"].ToString());
                    if (!String.IsNullOrEmpty(rd["branchid"].ToString())) obj.Branchid = int.Parse(rd["branchid"].ToString());
                    if (!String.IsNullOrEmpty(rd["saccoofficial"].ToString())) obj.Saccoofficial = int.Parse(rd["saccoofficial"].ToString());
                    if (!String.IsNullOrEmpty(rd["mdesignationid"].ToString())) obj.Mdesignationid = rd["mdesignationid"].ToString();
                    if (!String.IsNullOrEmpty(rd["mwithdrawaldate"].ToString())) obj.Mwithdrawaldate = DateTime.Parse(rd["mwithdrawaldate"].ToString());
                    if (!String.IsNullOrEmpty(rd["statusid"].ToString())) obj.Statusid = rd["statusid"].ToString();
                    if (!String.IsNullOrEmpty(rd["blocked"].ToString())) obj.Blocked = int.Parse(rd["blocked"].ToString());
                    if (!String.IsNullOrEmpty(rd["mgross"].ToString())) obj.Mgross = double.Parse(rd["mgross"].ToString());
                  //  if (!String.IsNullOrEmpty(rd["mnett"].ToString())) obj.Mnett = double.Parse(rd["mnett"].ToString());
                    if (!String.IsNullOrEmpty(rd["gradeid"].ToString())) obj.Gradeid = rd["gradeid"].ToString();
                    if (!String.IsNullOrEmpty(rd["RegFee"].ToString())) obj.RegFee = double.Parse(rd["RegFee"].ToString());
                    if (!String.IsNullOrEmpty(rd["DatePaid"].ToString())) obj.DatePaid = DateTime.Parse(rd["DatePaid"].ToString());
                    if (!String.IsNullOrEmpty(rd["mstatus"].ToString())) obj.Mstatus = rd["mstatus"].ToString();
                    if (!String.IsNullOrEmpty(rd["IDNO"].ToString())) obj.IDNO = rd["IDNO"].ToString();
                    if (!String.IsNullOrEmpty(rd["LOANBALatMIGRATION"].ToString())) obj.LOANBALatMIGRATION = double.Parse(rd["LOANBALatMIGRATION"].ToString());
                    if (!String.IsNullOrEmpty(rd["FILENO"].ToString())) obj.FILENO = rd["FILENO"].ToString();
                    if (!String.IsNullOrEmpty(rd["LevelofEducationID"].ToString())) obj.LevelofEducationID = int.Parse(rd["LevelofEducationID"].ToString());
                    if (!String.IsNullOrEmpty(rd["BusinessSectorID"].ToString())) obj.BusinessSectorID = int.Parse(rd["BusinessSectorID"].ToString());
                    if (!String.IsNullOrEmpty(rd["LevelofIncomeID"].ToString())) obj.LevelofIncomeID = int.Parse(rd["LevelofIncomeID"].ToString());
                    if (!String.IsNullOrEmpty(rd["DocumentTypeID"].ToString())) obj.DocumentTypeID = int.Parse(rd["DocumentTypeID"].ToString());
                    if (!String.IsNullOrEmpty(rd["WithdrawalReasonID"].ToString())) obj.WithdrawalReasonID = int.Parse(rd["WithdrawalReasonID"].ToString());

                }
                try { rd.Close(); }
                catch {; }
            }
            return obj;
        }
        public string getMemberName(string MemberNo)
        {

            //MemberNo = "";
            Link myLink = new Link();
            DbDataReader rd = myLink.GetDBResults(ref err, "proc_CheckMemberNo", "@mcode", MemberNo
               );
            if (err == "")
            {
                if (rd.Read())
                {

                    this.MemberName  = rd["MemberName"].ToString();

                }
                try { rd.Close(); }
                catch {; }

            }
            return MemberName;
        }
    }
}
