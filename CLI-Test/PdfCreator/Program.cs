using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Hello World");
                if(args.Length > 0)
                {
                    if (args[1] == "facstaff")
                    {
                        FacStaff facStaff = new FacStaff();
                        facStaff.GenerateFile();
                    }
                }

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }

        }
    }
}
