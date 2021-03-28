using System;

namespace PdfCreator
{
    class Employee
    {
        public Employee() => Console.WriteLine("New Employee()");

        public string _name;
        public string _phoneNumber;
        public string _emailAddress;
        public string _department;
        public string _title;
        public string _faxNumber;
        public string _building;
        public string _office;
        public string _url;

        public string Name
        {
            get => _name;
            set => _name = value;
        }
        public string PhoneNumber
        {
            get => _phoneNumber;
            set => _phoneNumber = value;
        }
        public string EmailAddress
        {
            get => _emailAddress;
            set => _emailAddress = value;
        }
        public string Department
        {
            get => _department;
            set => _department = value;
        }
        public string Title
        {
            get => _title;
            set => _title = value;
        }
        public string FaxNumber
        {
            get => _faxNumber;
            set => _faxNumber = value;
        }
        public string Building
        {
            get => _building;
            set => _building = value;
        }
        public string Office
        {
            get => _office;
            set => _office = value;
        }
        public string Url
        {
            get => _url;
            set => _url = value;
        }

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
        }
    }
}
