using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Collections;

namespace ReadExcel
{
   public class Link
    {
        DbDataReader rs = null;
        DbDataReader ResultSet = null;

        string msg = "";


      

        public int executeString(string sql)
        {
            int x = 1;
            SqlConnection conn = new SqlConnection(GlobalVariable.fetchedconnectionstring);
            SqlCommand cmd = new SqlCommand(sql, conn);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            try
            {
                x = cmd.ExecuteNonQuery();
            }
            catch
            {
                x = 0;
            }
            finally
            {
                conn.Close();
            }
            return x;
        }


        public void  GetMetroResults()
        {

            int i = 0;
            

            
                MetroPerson oconfig = new MetroPerson();
                oconfig = MetroPerson.Deserialize("metropol.xml");

                string svrname = oconfig.api_code.ToString();
                string dbname = oconfig.regoffice.ToString();
                string username = oconfig.identity_number;
                string pwd = oconfig.gender; // decript if twas encrypted

                
            //}

        }



        public DbDataReader GetDBResults(ref String errMsg, string StoredProcedure, params object[] ProcedureParameters)
        {

            int i = 0;
            errMsg = "";

            if (GlobalVariable.fetchedconnectionstring == "")
            {
                Configuration oconfig = new Configuration();
                oconfig = Configuration.Deserialize("config.xml");

                string svrname = oconfig.ServerName;
                string dbname = oconfig.DbName;
                string username = oconfig.UserName;
                string pwd = oconfig.Password; // decript if twas encrypted

                GlobalVariable.dbDatabaseName = dbname;
                GlobalVariable.dbPassword = pwd;
                GlobalVariable.dbServerName = svrname;
                GlobalVariable.dbUserName = username;

                GlobalVariable.fetchedconnectionstring = "Data Source=" + svrname + ";Initial Catalog=" + dbname + ";User ID=" + username + ";Password=" + pwd + ";Integrated Security=false; Connect Timeout=3000";
            }

            if (GlobalVariable.fetchedconnectionstring == "") return null;


            //string ParameterName = "";
            string SQLStatement = "";
            string DataProviderName = "System.Data.SqlClient";
            DbProviderFactory dpf = DbProviderFactories.GetFactory(DataProviderName);
            DbConnection objconnection = dpf.CreateConnection();
            objconnection.ConnectionString = GlobalVariable.fetchedconnectionstring; // "Data Source=USER-030C791BB9;Initial Catalog=Version2;Integrated Security=True";
            try
            {
                objconnection.Open();
            }
            catch (System.InvalidOperationException ex)
            {
                errMsg = ex.Message.ToString();
                return rs;
            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                errMsg = ex.Message.ToString();
                return rs;
            }
            catch (System.Exception ex)
            {
                errMsg = ex.Message.ToString();
                return rs;
            }

            DbCommand Command = dpf.CreateCommand();
            Command.Connection = objconnection;

            string mystring = "";

            //if (HttpContext.Current.Session["DBType"].ToString() == "MSSQL")
            //{
            foreach (object o in ProcedureParameters)
            {
                if (i % 2 == 0)
                    SQLStatement = SQLStatement + (string)o + "=";
                else
                {
                    if (o.GetType() == typeof(string))
                    {
                        mystring = (string)o;
                        mystring = mystring.Replace("'", "''");
                        SQLStatement = SQLStatement + " '" + mystring + "',";
                    }
                    else if (o.GetType() == typeof(DateTime))
                        SQLStatement = SQLStatement + " '" + o.ToString() + "',";
                    else if (o.GetType() == typeof(Int32))
                        SQLStatement = SQLStatement + o.ToString() + ",";
                    else if (o.GetType() == typeof(Double))
                        SQLStatement = SQLStatement + o.ToString() + ",";
                    else if (o.GetType() == typeof(Int32))
                        SQLStatement = SQLStatement + o.ToString() + ",";
                    else if (o.GetType() == typeof(Byte[]))
                        SQLStatement = SQLStatement + o.ToString() + ",";
                    else if (o.GetType() == typeof(Boolean))
                    {
                        if ((bool)o == true)
                            SQLStatement = SQLStatement + " 1" + ",";
                        else
                            SQLStatement = SQLStatement + " 0" + ",";
                    }
                    else
                        SQLStatement = SQLStatement + " '" + (string)o + "',";

                }
                i = i + 1;
            }
            if (SQLStatement.Length > 1)
                SQLStatement = SQLStatement.Substring(0, SQLStatement.Length - 1);



            StoredProcedure = "EXEC " + StoredProcedure + " " + SQLStatement;
            Command.CommandTimeout = 3600;
            Command.CommandType = CommandType.Text;
            Command.CommandText = StoredProcedure;
            try
            {

                ResultSet = Command.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (SqlException ex)
            {
                errMsg = ex.Message.ToString();
                return rs;
            }
            return ResultSet;
            //}

        }

        #region -- Configuration Class --
        /// <summary>
        /// This Configuration class is basically just a set of 
        /// properties with a couple of static methods to manage
        /// the serialization to and deserialization from a
        /// simple XML file.
        /// </summary>
        [Serializable]
        public class Configuration
        {
            string _servername;
            string _dbname;
            string _userid;
            string _pwd;

            public Configuration()
            {
                _servername = "";
                _dbname = "";
                _userid = "";
                _pwd = "";
            }
            public static void Serialize(string file, Configuration c)
            {
                System.Xml.Serialization.XmlSerializer xs
                   = new System.Xml.Serialization.XmlSerializer(c.GetType());
                StreamWriter writer = File.CreateText(file);
                xs.Serialize(writer, c);
                writer.Flush();
                writer.Close();
            }
            public static Configuration Deserialize(string file)
            {
                System.Xml.Serialization.XmlSerializer xs
                   = new System.Xml.Serialization.XmlSerializer(
                      typeof(Configuration));
                StreamReader reader = File.OpenText(file);
                Configuration c = (Configuration)xs.Deserialize(reader);
                reader.Close();
                return c;
            }

            public string ServerName
            {
                get { return _servername; }
                set { _servername = value; }
            }
            public string DbName
            {
                get { return _dbname; }
                set { _dbname = value; }
            }

            public string UserName
            {
                get { return _userid; }
                set { _userid = value; }
            }
            public string Password
            {
                get { return _pwd; }
                set { _pwd = value; }
            }

        }
        #endregion

        #region -- Configuration Class --
        /// <summary>
        /// This Configuration class is basically just a set of 
        /// properties with a couple of static methods to manage
        /// the serialization to and deserialization from a
        /// simple XML file.
        /// </summary>
        [Serializable]
        public class MetroPerson
        {
            public object api_code { get; set; }
            public object api_code_description { get; set; }
            public string citizenship { get; set; }
            public object clan { get; set; }
            public string date_of_birth { get; set; }
            public object date_of_death { get; set; }
            public object date_of_issue { get; set; }
            public string dob { get; set; }
            public object error { get; set; }
            public object error_message { get; set; }
            public object ethnic_group { get; set; }
            public object family { get; set; }
            public object fingerprint { get; set; }
            public string first_name { get; set; }
            public string gender { get; set; }
            public bool has_error { get; set; }
            public string id_number { get; set; }
            public string identity_number { get; set; }
            public string identity_type { get; set; }
            public int identity_type_id { get; set; }
            public string ipaddress { get; set; }
            public string last_name { get; set; }
            public object occupation { get; set; }
            public string other_name { get; set; }
            public object photo { get; set; }
            public object pin { get; set; }
            public object place_of_birth { get; set; }
            public object place_of_death { get; set; }
            public object place_of_live { get; set; }
            public object regoffice { get; set; }
            public string serial_number { get; set; }
            public object signature { get; set; }
            public bool success { get; set; }
            public string surname { get; set; }
            public string trx_id { get; set; }
            public string response { get; set; }
            public static void Serialize(string file, MetroPerson c)
            {
                System.Xml.Serialization.XmlSerializer xs
                   = new System.Xml.Serialization.XmlSerializer(c.GetType());
                StreamWriter writer = File.CreateText(file);
                xs.Serialize(writer, c);
                writer.Flush();
                writer.Close();
            }
            public static MetroPerson Deserialize(string file)
            {
                System.Xml.Serialization.XmlSerializer xs
                   = new System.Xml.Serialization.XmlSerializer(
                      typeof(MetroPerson));
                StreamReader reader = File.OpenText(file);
                MetroPerson c = (MetroPerson)xs.Deserialize(reader);
                reader.Close();
                return c;
            }

        }
        #endregion
    }
}
