using System;

namespace PdfCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                bool emptyArray = (args == null || args.Length == 0);
                string arg;

                if (emptyArray)
                {
                    arg = "Hello World!";
                    LoopParams();
                }
                else
                    arg = args[0].ToString().ToLower();
 
                if(arg.StartsWith("-"))
                {
                    switch (arg)
                    {
                        case "-help":
                        case "-h":
                            LoopParams();
                            return;
                        default:
                            arg = arg.Remove(0, 1);
                            break;
                    }
                }

                PdfRender pdfRender = new PdfRender(arg);
                // pass 2nd argument as input string
                if (pdfRender.PdfType == PdfRender.PdfTypes.Test)
                    if (args.Length >= 2)
                        arg = args[1].ToString();
                pdfRender.RequestPdf(arg);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        static void LoopParams()
        {
            foreach (PdfRender.PdfTypes pdfTypes in Enum.GetValues(typeof(PdfRender.PdfTypes)))
            {
                Console.WriteLine("-" + pdfTypes.ToString());
            }
            Console.WriteLine("******************************");
        }
    }
}
