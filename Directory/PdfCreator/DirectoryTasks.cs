namespace PdfCreator
{
    class DirectoryTasks : DataConnectors
    {
        public DirectoryTasks()
        {

        }

        /// <summary>
        /// Dervied classes do not inherit constructors from base class; this allows the constructor to be passed to the base class
        /// </summary>
        /// <param name="provider"></param>
        public DirectoryTasks(string provider) : base(provider)
        {

        }
    }
}
