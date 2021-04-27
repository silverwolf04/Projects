using System;

namespace PdfCreator
{
    class DirectoryTasks : DataConnectors
    {
        /// <summary>
        /// base constructor with no params passed
        /// </summary>
        public DirectoryTasks()
        {

        }

        /// <summary>
        /// Pass the pdftype to constructor; will set the DataProvider and ConnectionString for the type provided
        /// </summary>
        /// <param name="provider"></param>
        public DirectoryTasks(PdfRender.PdfTypes pdfTypes)
        {
            switch(pdfTypes)
            {
                case PdfRender.PdfTypes.Department:
                    // Parse the DataProvider for DepartmentPDF
                    ParseProvider(Properties.DepartmentPdf.Default.DataProvider);
                    // Invalid DataProvider will be 'Unknown' if no valid one is parsed
                    // Use the Global DatapProvider and ConnectionString instead
                    if (DataProvider == DataProviders.Unknown)
                    {
                        Console.WriteLine("Using Global DataProvider and ConnectionString");
                        ParseProvider(Properties.Settings.Default.DataProvider);
                        ConnectionString = Properties.Settings.Default.ConnectionString;
                    }

                    if (string.IsNullOrEmpty(ConnectionString))
                        ConnectionString = Properties.DepartmentPdf.Default.ConnectionString;
                    break;
                case PdfRender.PdfTypes.FacultyStaff:
                    // Parse the DataProvider for FacultyStaffPDF
                    ParseProvider(Properties.FacStaffPdf.Default.DataProvider);
                    // Invalid DataProvider will be 'Unknown' if no valid one is parsed
                    // Use the Global DatapProvider and ConnectionString instead
                    if(DataProvider == DataProviders.Unknown)
                    {
                        Console.WriteLine("Using Global DataProvider and ConnectionString");
                        ParseProvider(Properties.Settings.Default.DataProvider);
                        ConnectionString = Properties.Settings.Default.ConnectionString;
                    }

                    if(string.IsNullOrEmpty(ConnectionString))
                        ConnectionString = Properties.FacStaffPdf.Default.ConnectionString;

                    // a single query string exists for Faculty Staff
                    if(string.IsNullOrEmpty(QueryString))
                        QueryString = Properties.FacStaffPdf.Default.QueryString;
                    break;
            }
        }
    }
}
