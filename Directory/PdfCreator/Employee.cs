using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfCreator
{
    class Employee
    {
        public Employee()
        {

        }

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
            get { return _name; }
            set { _name = value; }
        }
        public string PhoneNumber
        {
            get { return _phoneNumber; }
            set { _phoneNumber = value; }
        }
        public string EmailAddress
        {
            get { return _emailAddress; }
            set { _emailAddress = value; }
        }
        public string Department
        {
            get { return _department; }
            set { _department = value; }
        }
        public string Title
        {
            get { return _title; }
            set { _title = value; }
        }
        public string FaxNumber
        {
            get { return _faxNumber; }
            set { _faxNumber = value; }
        }
        public string Building
        {
            get { return _building; }
            set { _building = value; }
        }
        public string Office
        {
            get { return _office; }
            set { _office = value; }
        }
        public string Url
        {
            get { return _url; }
            set { _url = value; }
        }

        public void ClearAll()
        {
            Name = "";
            PhoneNumber = "";
            EmailAddress = "";
            Department = "";
            Title = "";
            FaxNumber = "";
            Building = "";
            Office = "";
        }
    }
}
