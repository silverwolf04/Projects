using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;

namespace PdfCreator
{
    static class DataConnectorsExt
    {
        public static string GetValue(this DataRow dataRow, string column)
        {
            string fieldVal = null;
            if (dataRow.Table.Columns.Contains(column))
                fieldVal = dataRow[column].ToString();

            return fieldVal;
        }
    }
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
            Console.WriteLine("Query: {0}", QueryString);
            Console.WriteLine("Using {0} DataProvider", DataProvider.ToString());

            switch (DataProvider)
            {
                case DataConnectors.DataProviders.MSSQL:
                    dataTable = GetDataSQL();
                    break;
                case DataConnectors.DataProviders.Oracle:
                    dataTable = GetDataOracle();
                    break;
                case DataConnectors.DataProviders.Excel:
                    // Requires Access Engine; 32bit for Any CPU or 64bit for x64 compile
                    // https://www.microsoft.com/en-us/download/details.aspx?id=13255
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
            DataTable dataTable = new DataTable();
            string importFilename = ConnectionString;
            string fileExtension = Path.GetExtension(importFilename);

            using (OleDbConnection conn = new OleDbConnection())
            {
                if (!File.Exists(importFilename))
                    Console.WriteLine("FilePath does not exist:{0}", importFilename);
                switch (fileExtension)
                {
                    case ".xls":
                        conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + importFilename + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                        break;
                    case ".xlsx":
                        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + importFilename + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
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
                        da.Fill(dataTable);
                    }
                }
            }
            return dataTable;
        }
    }
}
