namespace PdfCreator
{
    class DirectoryTasks : DataConnectors
    {
        public DirectoryTasks()
        {
            // base constructor with no params passed
        }

        /// <summary>
        /// Dervied classes do not inherit constructors from base class; this allows the constructor to be passed to the base class
        /// </summary>
        /// <param name="provider"></param>
        public DirectoryTasks(PdfRender.PdfTypes pdfTypes)
        {
            switch(pdfTypes)
            {
                case PdfRender.PdfTypes.Department:
                    if(DataProvider == DataProviders.MSSQL)
                        ParseProvider(Properties.DepartmentPdf.Default.DataProvider);
                    if (string.IsNullOrEmpty(ConnectionString))
                        ConnectionString = Properties.DepartmentPdf.Default.ConnectionString;
                    break;
                case PdfRender.PdfTypes.FacultyStaff:
                    if(DataProvider == DataProviders.MSSQL)
                        ParseProvider(Properties.FacStaffPdf.Default.DataProvider);
                    if(string.IsNullOrEmpty(ConnectionString))
                        ConnectionString = Properties.FacStaffPdf.Default.ConnectionString;
                    if(string.IsNullOrEmpty(QueryString))
                        QueryString = Properties.FacStaffPdf.Default.QueryString;
                    break;
            }
        }
    }
}
