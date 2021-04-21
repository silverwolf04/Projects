
namespace PdfCreator
{
    class Entry
    {
        public Entry()
        {

        }
        public string Name { get; set; }
        public string PhoneNumber { get; set; }
        public string EmailAddress { get; set; }
        public string Department { get; set; }
        public string Title { get; set; }
        public string FaxNumber { get; set; }
        public string Building { get; set; }
        public string Office { get; set; }
        public string Url { get; set; }
        public string Notes { get; set; }
        public string Category { get; set; }
        public string CellNumber { get; set; }
        public string Address { get; set; }
        public void ClearAll()
        {
            Name = string.Empty;
            PhoneNumber = string.Empty;
            EmailAddress = string.Empty;
            Department = string.Empty;
            Title = string.Empty;
            FaxNumber = string.Empty;
            Building = string.Empty;
            Office = string.Empty;
            Notes = string.Empty;
            CellNumber = string.Empty;
        }
    }
}
