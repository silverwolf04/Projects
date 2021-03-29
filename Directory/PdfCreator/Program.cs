﻿using System;

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
                    arg = args[0].ToString();
 
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
