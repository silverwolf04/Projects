using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace PdfCreator
{
    class DataConnectors
    {
        public void GetData()
        {
            DataTable dataTable = ReadExcelFile("Test", "text.xlsx");

            DataColumnCollection dataColumnCollection = dataTable.Columns;
            int loopInt = 1;

            foreach(DataRow dataRow in dataTable.Rows)
            {

                foreach(DataColumn col in dataColumnCollection)
                {
                    Console.WriteLine("Row:{0} Column:{1} Value:{2}", loopInt, col, dataRow[col]);
                }

                loopInt++;
            }
        }
        private DataTable ReadExcelFile(string sheetName, string path)
        {

            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                if (fileExtension == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }
    }
}
