using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Security.Cryptography;
using System.Data.SqlClient;
using System.Data;

namespace ReadExcel
{
    class GlobalVariable
    {
    
        public static string fetchedconnectionstring = "";
      
      

        public static string dbServerName = "";
        public static string dbDatabaseName = "";
        public static string dbUserName = "";
        public static string dbPassword = ""; // decript if twas encrypted
  

        public static string BuildXmlString(string xmlRootName, string[] values)
        {
            StringBuilder xmlString = new StringBuilder();

            xmlString.AppendFormat("<{0}>", xmlRootName);
            for (int i = 0; i < values.Length; i++)
            {
                xmlString.AppendFormat("<value>{0}</value>", values[i]);
            }
            xmlString.AppendFormat("</{0}>", xmlRootName);

            return xmlString.ToString();
        }
       
        public static string Encrypt(string toEncrypt, bool useHashing, string nkey)
        {
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);

            System.Configuration.AppSettingsReader settingsReader =
                                                new AppSettingsReader();
            // Get the key from config file

            string key = nkey + "sUiLeNrOc";
            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
              

                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
         
            tdes.Key = keyArray;
      
            tdes.Mode = CipherMode.ECB;
    

            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
          
            byte[] resultArray =
              cTransform.TransformFinalBlock(toEncryptArray, 0,
              toEncryptArray.Length);
           
            tdes.Clear();
          
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }
       

        public static int executeString(string sql)
        {
            int x = 1;
            //string err = "";
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
     
    }
}
