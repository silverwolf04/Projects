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
            Unknown,
            MSSQL,
            Oracle,
            Excel
        }
        public DataConnectors() => DataProvider = DataProviders.Unknown;
        public DataConnectors(string provider)
        {
            ParseProvider(provider);
        }
        public void ParseProvider(string provider)
        {
            if (Enum.IsDefined(typeof(DataProviders), provider))
                DataProvider = (DataProviders)Enum.Parse(typeof(DataProviders), provider, true);
        }
        public DataTable GetData()
        {
            DataTable dataTable = new DataTable();
            Console.WriteLine("***********************");
            Console.WriteLine("DataProvider:{0}", DataProvider.ToString());
            Console.WriteLine("Query: {0}", QueryString);

            try
            {
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
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
            Console.WriteLine("Number of Rows:{0}", dataTable.Rows.Count);
            return dataTable;
        }
        private DataTable GetDataSQL()
        {
            DataTable dataTable = new DataTable();
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                using (SqlCommand sqlCommand = new SqlCommand(QueryString, connection))
                {
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
                    sqlDataAdapter.Fill(dataTable);
                }
            }
            return dataTable;
        }
        private DataTable GetDataOracle()
        {
            DataTable dataTable = new DataTable();
            using (OracleConnection conn = new OracleConnection(ConnectionString))
            {
                
                conn.Open();
                OracleCommand cmd = new OracleCommand
                {
                    Connection = conn,
                    CommandText = QueryString,
                    CommandType = CommandType.Text
                };
                OracleDataReader dr = cmd.ExecuteReader();
                dataTable.Load(dr);
            }
            return dataTable;
        }
        private DataTable GetDataExcel()
        {
            DataTable dataTable = new DataTable();
            string importFilename = ConnectionString;
            string fileExtension = Path.GetExtension(importFilename);
            if (!File.Exists(importFilename))
                Console.WriteLine("FilePath does not exist:{0}", importFilename);

            using (OleDbConnection conn = new OleDbConnection())
            {
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
