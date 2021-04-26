
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
                    ParseProvider(Properties.DepartmentPdf.Default.DataProvider);
                    if(DataProvider == DataProviders.Unknown)
                    {
                        ParseProvider(Properties.Settings.Default.DataProvider);
                        ConnectionString = Properties.Settings.Default.ConnectionString;
                    }
                    if (string.IsNullOrEmpty(ConnectionString))
                        ConnectionString = Properties.DepartmentPdf.Default.ConnectionString;
                    break;
                case PdfRender.PdfTypes.FacultyStaff:
                    ParseProvider(Properties.FacStaffPdf.Default.DataProvider);
                    if(DataProvider == DataProviders.Unknown)
                    {
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
