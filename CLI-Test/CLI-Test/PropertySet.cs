using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLI_Test
{
    public struct ListOfValues
    {
        public string Name { get; set; }
        public string Number { get; set; }
        public string Email { get; set; }

        public ListOfValues(string name, string number, string email)
        {
            if (string.IsNullOrEmpty(name))
            {
                Name = null;
            }
            else
            {
                Name = name;
            }
            Number = number;
            Email = email;
        }
    }

    public class PropertySetUsingStruct
    {
        public void ReadValues(ListOfValues listOfValues)
        {
            Console.WriteLine(listOfValues.Name);
            Console.WriteLine(listOfValues.Number);
            Console.WriteLine(listOfValues.Email);
        }
    }

    public class PropertySetUsingClass
    {
        public string name;
        public string number;
        public string email;

        public void ReadValues()
        {
            Console.WriteLine(email);
            Console.WriteLine(name);
            Console.WriteLine(number);
        }
    }

}
