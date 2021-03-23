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
            PdfRender pdfRender = new PdfRender();
            if (string.Equals(args[0], "facultystaffpdf", StringComparison.OrdinalIgnoreCase))
            {
                pdfRender.FacultyStaffPDF();
            }
            else
            {
                pdfRender.GenerateTest(args[0]);
            }
        }
    }
}
