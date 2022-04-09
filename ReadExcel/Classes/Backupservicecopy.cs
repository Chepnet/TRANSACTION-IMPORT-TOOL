using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace ReadExcel.Classes
{
    class Backupservicecopy
    {
        //"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database1.mdf;Integrated Security=True;User Instance=True"); 
        private  string _connectionString = "Data Source=USER-PC\\SQLEXPRESS;User ID = sa; Password=pass123";
        private  string _backupFolderFullPath = "E:\\micropoint stuff\\PROJECTS\\BACKUPFOLDER";
        private  string[] _systemDatabaseNames = { "master", "tempdb", "model", "msdb" };
         
        public void BackupService(string connectionString, string backupFolderFullPath)
        {
            _connectionString = connectionString;
            _backupFolderFullPath = backupFolderFullPath;
        }

       
        void OnDependencyChange(object sender, SqlNotificationEventArgs e)
        {
            // Handle the event (for example, invalidate this cache entry).
        }



        public void BackupAllUserDatabases()
        {
            foreach (string databaseName in GetAllUserDatabases())
            {
                BackupDatabase(databaseName);
            }
        }

        public void BackupDatabase(string databaseName)
        {
            string filePath = BuildBackupPathWithFilename(databaseName);

            using (var connection = new SqlConnection(_connectionString))
            {
                var query = String.Format("BACKUP DATABASE [{0}] TO DISK='{1}'", databaseName, filePath);

                using (var command = new SqlCommand(query, connection))
                {

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }

        private void UpdateDataTable(DataTable table,  OleDbDataAdapter myDataAdapter)
        {
            DataTable xDataTable = table.GetChanges();

            // Check the DataTable for errors.
            if (xDataTable.HasErrors)
            {
                // Insert code to resolve errors.
            }

            // After fixing errors, update the database with the DataAdapter 
            myDataAdapter.Update(xDataTable);
        }

        private IEnumerable<string> GetAllUserDatabases()
        {
            var databases = new List<String>();

            DataTable databasesTable;

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                databasesTable = connection.GetSchema("Databases");//Obtaining database table names

                connection.Close();
            }

            foreach (DataRow row in databasesTable.Rows)
            {
                string databaseName = row["database_name"].ToString();

                if (_systemDatabaseNames.Contains(databaseName))
                    continue;
                //databasesTable.Select(DateTime.Now.ToString("yyyy-MM-dd"));
                //if(row.RowState!=DataRowState.Unchanged )
                //{
                //    databases.Add(databaseName);

                //}

                //databasesTable.GetChanges();
                if(row.RowState !=DataRowState.Unchanged )
                {

                }
                databasesTable.RowChanged += DatabasesTable_RowChanged1;
                
                databases.Add(databaseName);

            }

            return databases;
        }

        private void DatabasesTable_RowChanged1(object sender, DataRowChangeEventArgs e)
        {
            Console.WriteLine("Row_Changed Event: name={0}; action={1}",
        e.Row["name"], e.Action);
        }

        private void DatabasesTable_RowChanged(object sender, DataRowChangeEventArgs e)
        {
            
        }

        private string BuildBackupPathWithFilename(string databaseName)
        {
            string filename = string.Format("{0}-{1}.bak", databaseName, DateTime.Now.ToString("yyyy-MM-dd"));

            return Path.Combine(_backupFolderFullPath, filename);
        }
    }
}
