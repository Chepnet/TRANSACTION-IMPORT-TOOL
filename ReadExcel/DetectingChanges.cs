using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReadExcel
{
    public partial class DetectingChanges : Form
    {
        public DetectingChanges()
        {
            InitializeComponent();
        }
        string _connectionString = "Data Source=USER-PC\\SQLEXPRESS;Initial Catalog=PFUPGRADE;User ID=sa;Password=pass123";
        string queeingName = "";
        private void button1_Click(object sender, EventArgs e)
        { string _connectionString = "Data Source=USER-PC\\SQLEXPRESS;Initial Catalog=PFUPGRADE;User ID=sa;Password=pass123";
            SqlConnection SqlConnection = new SqlConnection( _connectionString );
            SqlDependency.Stop(_connectionString);
            SqlDependency.Start(_connectionString);
            SqlConnection.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = SqlConnection;
            command.CommandType = CommandType.Text;
            //command.CommandText = " SELECT [PatientCode] ,[EmailAddress] , SentTime  FROM [dbo].[EmailNotificationHistory]  where  [SentTime] is null";
            command.CommandText = " SELECT * FROM tblKins  ";
            // Create a dependency and associate it with the SqlCommand.
            //command.Notification = null;
            SqlDependency dependency = new SqlDependency(command);
            // Maintain the refence in a class member.  
           
            // Subscribe to the SqlDependency event.  , Its using sql server broker service. for this broker service must be enabled for this database.
            dependency.OnChange += new OnChangeEventHandler(OnDependencyChange);

           
            // Get the messages
            command.ExecuteReader();
        }
       

        private void button2_Click(object sender, EventArgs e)
        {
            SomeMethod();
        }
            void Initialization()
{
                // Create a dependency connection.
                SqlDependency.Start(_connectionString , queeingName);
            }

            void SomeMethod()
{
            SqlConnection connection = new SqlConnection(_connectionString);
            SqlDependency.Stop(_connectionString);
            SqlDependency.Start(_connectionString);
            // Assume connection is an open SqlConnection.
            connection.Open();
                // Create a new SqlCommand object which directly references (no synonyms) the data you want to check for changes.
                using (SqlCommand command = new SqlCommand("SELECT * from tblkins", connection))
                {
                    // Create a dependency and associate it with the SqlCommand.
                    SqlDependency dependency = new SqlDependency(command);
                    // Maintain the refence in a class member.

                    // Subscribe to the SqlDependency event.
                    dependency.OnChange += new OnChangeEventHandler(OnDependencyChange);

                    // Execute the command.
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Process the DataReader.
                    }
                }
            }

            // Handler method
            void OnDependencyChange(object sender, SqlNotificationEventArgs e)
            {
    // Handle the event (for example, invalidate this cache entry).
}

void Termination()
{
                // Release the dependency.
                SqlDependency.Stop(_connectionString, queeingName);
            }
        }
    
}
