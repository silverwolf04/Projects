using System;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using System.Data;
using Hyland.Unity;


namespace Utilities
{
	public class OracleDatabaseAdapter
	{
		public enum StatusCodes
		{
			ConnectionOK,
			ConnectionFail,
			ConnectionStringMissing,
			Unknown
		}

		// datatable to test with
		public DataTable SampleTable()
		{
			DataTable myDataTable = new DataTable();
			myDataTable.Columns.Add("Test");
			myDataTable.Rows.Add("Sample Field");
			return myDataTable;
		}

		public string GetConnStringFromOnBase()
        {
			return "";
        }

		private static void AddParametersToCmd(Dictionary<string, string> queryParams, ref OracleCommand oracleCommand)
        {
			if (queryParams != null)
			{
				foreach (var entry in queryParams)
				{
					oracleCommand.Parameters.Add(entry.Key, entry.Value);
				}
			}
		}

		public StatusCodes OracleQuery(string sqlQuery, Dictionary<string, string> queryParams, out DataTable myDataTable)
        {
			StatusCodes statusCodes = StatusCodes.Unknown;
			myDataTable = new DataTable();
			OracleDataReader myDataReader = null;
			OracleCommand myCmd = new OracleCommand
			{
				CommandText = sqlQuery
			};

			try
            {
				string connectionString = CLI_Test.Properties.Settings.Default.ConnectionString;
				if (connectionString.Length == 0)
					statusCodes = StatusCodes.ConnectionStringMissing;

                AddParametersToCmd(queryParams, ref myCmd);
				myCmd.Connection = new OracleConnection(connectionString);
				myCmd.Connection.Open();
				statusCodes = StatusCodes.ConnectionOK;
				// CommandBehavior.CloseConnection will request the OracleCommand to close the connection
				// ... this occurs after supplying the recordset to the OracleDataReader object
				myDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection);

				if (myDataReader.HasRows)
					myDataTable.Load(myDataReader);
			}
			catch(Exception ex)
            {
				Console.WriteLine(ex.ToString());
            }

			if (myDataReader != null) { myDataReader.Dispose(); }
			myCmd.Dispose();
			
			return statusCodes;
        }

		// Assembles the Oracle connection, queries based on the string text & parameters, and returns the recordset to the caller row by row.
		public StatusCodes OracleQuery(Application app, string sqlQuery, Dictionary<string, string> queryParams, out DataTable myDataTable, string connectionStringName = "BannerSecureConnectionString")
		{
			StatusCodes statusCodes = StatusCodes.Unknown;
            OracleCommand myCmd = new OracleCommand
            {
                CommandText = sqlQuery
            };
            //DataTable myDataTable = new DataTable();
            OracleDataReader myDataReader = null;
			// scrub the datatable
			myDataTable = new DataTable();

			try
			{
				AddParametersToCmd(queryParams, ref myCmd);
				// create dictionary object to return to caller
				//Dictionary<string, string> records = new Dictionary<string, string>();
				string connectionString = GetConfigurationItem<string>(connectionStringName, app);
				myCmd.Connection = new OracleConnection(connectionString);
				myCmd.Connection.Open();
				statusCodes = StatusCodes.ConnectionOK;

				// CommandBehavior.CloseConnection will request the OracleCommand to close the connection
				// ... this occurs after supplying the recordset to the OracleDataReader object
				myDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection);

				if (myDataReader.HasRows)
					myDataTable.Load(myDataReader);
			}
			catch (Exception ex)
			{
				app.Diagnostics.Write(ex);
				statusCodes = StatusCodes.ConnectionFail;
			}
			finally
			{
				if (myDataReader != null) { myDataReader.Dispose(); }
				myCmd.Dispose();
			}
			return statusCodes;
		}

		private T GetConfigurationItem<T>(string key, Application app)
		{
#pragma warning disable IDE0059 // Unnecessary assignment of a value
            string value = String.Empty;
#pragma warning restore IDE0059 // Unnecessary assignment of a value

            if (!app.Configuration.TryGetValue(key, out value))
			{
				throw new Exception(String.Format("Required Unity Configuration Item '{0}' was not found", key));
			}

			return (T)Convert.ChangeType(value, typeof(T));
		}
	}
}
