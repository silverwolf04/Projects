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
            PdfRender pdfRender = new PdfRender(args);
            string arg;

            if(args.Length > 0)
            {
                arg = args[0].ToString();
            }
            else
            {
                arg = "Hello World!";
            }

            Console.WriteLine(pdfRender.PdfType.ToString());
            pdfRender.RequestPdf(arg);
        }
    }
}
