using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLI_Test
{
    public struct Employee
    {
        public string Name { get; set; }
        public string PhoneNumber;
        public string EmailAddress;
        public string Department;
        public string Title;
        public string FaxNumber;

        public void ClearAll()
        {
            Name = "";
            PhoneNumber = "";
            EmailAddress = "";
            Department = "";
            Title = "";
            FaxNumber = "";
        }
    }

public class DirectoryTasks
    {
    }
}
