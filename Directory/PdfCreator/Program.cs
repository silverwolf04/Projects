using System;

namespace PdfCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string testString = "Hello World!";
                string arg = args.Length >= 1 ? args[0].ToString() : testString;

                if (arg.StartsWith("-"))
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
                Console.WriteLine("Argument: " + arg);
                PdfRender pdfRender = new PdfRender(arg);
                if(pdfRender.PdfType == PdfRender.PdfTypes.Test)
                {
                    LoopParams();
                    if (args.Length == 2)
                        arg = args[1].ToString();
                    else
                        arg = testString;
                }
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
