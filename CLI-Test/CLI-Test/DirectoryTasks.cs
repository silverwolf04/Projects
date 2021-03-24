using System;
using System.Data;
using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;

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
        public string Building;
        public string Office;

        public void ClearAll()
        {
            Name = "";
            PhoneNumber = "";
            EmailAddress = "";
            Department = "";
            Title = "";
            FaxNumber = "";
            Building = "";
            Office = "";
        }
    }

    public class DirectoryTasks
    {
        public void GetData(out DataTable dataTable)
        {
            string queryString = Properties.Settings.Default.QueryString;
            string connectionString = Properties.Settings.Default.ConnectionString;

            if (Properties.Settings.Default.DataProvider == "Oracle")
            {
                Console.WriteLine("Using Oracle DataProvider");
                GetDataOracle(out dataTable, connectionString, queryString);
            }
            else // MSSQL / SQL / SQLServer
            {
                Console.WriteLine("Using MSSQL DataProvider");
                GetDataSQL(out dataTable, connectionString, queryString);
            }
        }
        public void GetDataOracle(out DataTable dataTable, string connectionString, string queryString)
        {
            dataTable = new DataTable();

            try
            {
                using (OracleConnection conn = new OracleConnection(connectionString))
                {
                    /*
                    conn.Open();
                    OracleCommand cmd = new OracleCommand();
                    cmd.Connection = conn;
                    cmd.CommandText = queryString;
                    cmd.CommandType = CommandType.Text;
                    */
                    OracleCommand cmd = new OracleCommand
                    {
                        Connection = conn,
                        CommandText = queryString,
                        CommandType = CommandType.Text
                    };
                    OracleDataReader dr = cmd.ExecuteReader();
                    dataTable.Load(dr);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        public void GetDataSQL(out DataTable dataTable, string connectionString, string queryString)
        {
            dataTable = new DataTable();

            try
            {
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
