using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;
using System.Linq;

namespace PdfCreator
{
    class DataConnectors
    {
        public DataProviders DataProvider { get; set; }
        public string ConnectionString { get; set; } = string.Empty;
        public string QueryString { get; set; } = string.Empty;
        public enum DataProviders
        {
            MSSQL,
            Oracle,
            Excel
        }
        public DataConnectors() => DataProvider = DataProviders.MSSQL;
        public DataConnectors(string provider)
        {
            ParseProvider(provider);
        }
        public void ParseProvider(string provider)
        {
            if (Enum.IsDefined(typeof(DataProviders), provider))
                DataProvider = (DataProviders)Enum.Parse(typeof(DataProviders), provider, true);
            else
                // Always use MSSQL if invalid provider in config
                Console.WriteLine("Invalid provider parsed");
        }
        public DataTable GetData()
        {
            DataTable dataTable = new DataTable();

            switch (DataProvider)
            {
                case DataConnectors.DataProviders.MSSQL:
                    Console.WriteLine("Using MSSQL DataProvider");
                    dataTable = GetDataSQL();
                    break;
                case DataConnectors.DataProviders.Oracle:
                    Console.WriteLine("Using Oracle DataProvider");
                    dataTable = GetDataOracle();
                    break;
                case DataConnectors.DataProviders.Excel:
                    // Requires Access Engine; 32bit for Any CPU or 64bit for x64 compile
                    // https://www.microsoft.com/en-us/download/details.aspx?id=13255
                    Console.WriteLine("Using Excel DataProvider");
                    dataTable = GetDataExcel();
                    break;
                default:
                    break;
            }
            return dataTable;
        }
        private DataTable GetDataSQL()
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection connection = new SqlConnection(ConnectionString))
                {
                    Console.WriteLine("State: {0}", connection.State);
                    Console.WriteLine("Query: {0}", QueryString);

                    using (SqlCommand sqlCommand = new SqlCommand(QueryString, connection))
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

            return dataTable;
        }
        private DataTable GetDataOracle()
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (OracleConnection conn = new OracleConnection(ConnectionString))
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
                        CommandText = QueryString,
                        CommandType = CommandType.Text
                    };
                    OracleDataReader dr = cmd.ExecuteReader();
                    dataTable.Load(dr);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return dataTable;
        }
        private DataTable GetDataExcel()
        {
            Console.WriteLine("Command Text:{0}", QueryString);
            DataTable dt = new DataTable();
            string Import_FileName = ConnectionString;
            string fileExtension = Path.GetExtension(Import_FileName);

            using (OleDbConnection conn = new OleDbConnection())
            {
                if (!File.Exists(Import_FileName))
                    Console.WriteLine("FilePath does not exist:{0}", Import_FileName);
                switch (fileExtension)
                {
                    case ".xls":
                        conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                        break;
                    case ".xlsx":
                        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                        break;
                    default:
                        Console.WriteLine("Invalid file extension:{0}", fileExtension);
                        break;
                }

                using (OleDbCommand comm = new OleDbCommand())
                {
                    //comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.CommandText = QueryString;
                    comm.Connection = conn;

                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                    }
                }
            }
            return dt;
        }
    }
}
