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
                PdfRender pdfRender = new PdfRender(args);
                string arg;

                if (args.Length > 0)
                {
                    arg = args[0].ToString();
                }
                else
                {
                    arg = "Hello World!";
                    Console.WriteLine("Permitted arguments to pass; anything else is used for test rendering");
                    foreach (PdfRender.PdfTypes pdfTypes in Enum.GetValues(typeof(PdfRender.PdfTypes)))
                    {
                        Console.WriteLine(pdfTypes.ToString());
                    }
                    Console.WriteLine("******************************");
                }

                pdfRender.RequestPdf(arg);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
