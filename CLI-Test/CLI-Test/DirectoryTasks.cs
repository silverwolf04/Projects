using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLI_Test
{
    public struct Employee
    {
        public string Name { get; set; }
        public string PhoneNumber;
        public string EmailAddress;
        public string Department;
        public string Title;
        public string FaxNumber;

        public void ClearAll()
        {
            Name = "";
            PhoneNumber = "";
            EmailAddress = "";
            Department = "";
            Title = "";
            FaxNumber = "";
        }
    }

public class DirectoryTasks
    {
        public void GetData(out DataTable dataTable)
        {
            dataTable = new DataTable();

            try
            {
                //string queryString = "select empname, department, title, phonenumber, emailaddress, faxnumber from employee";
                string queryString = Properties.Settings.Default.QueryString;
                string connectionString = Properties.Settings.Default.ConnectionString;
                //if (connectionString.Length == 0)
                 //   statusCodes = StatusCodes.ConnectionStringMissing;

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    Console.WriteLine("State: {0}", connection.State);
                    Console.WriteLine("Query: {0}", queryString);

                    using(SqlCommand sqlCommand = new SqlCommand(queryString, connection))
                    {
                        SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
                        sqlDataAdapter.Fill(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
