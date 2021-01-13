using System;
using System.IO;

namespace CLI_Test
{
    class Program
    {
        private enum ClassInArgument
        {
            Eprint,
            UrlGenerate,
            PageDownload,
            ToCharTest,
            PropertySetUsingStruct,
            PropertySetUsingClass
        }

        static void Main(string[] args)
        {
            try
            {
                // which class to run against
                string programClass = args[0];
                ClassInArgument classesInArgument = new ClassInArgument();
                if(!Enum.TryParse(programClass, true, out classesInArgument))
                    Console.WriteLine("Invalid class argument passed:" + programClass);

                switch (classesInArgument)
                {
                    //eprint
                    case ClassInArgument.Eprint:
                        // the inputfile to read from
                        string inputFile = args[1];

                        //string fundOrgInput = "org:10000-15000";
                        string fundOrgInput = args[2];

                        //string outputfile = @"C:\users\dcover\desktop\output.txt";
                        string outputfile = args[3];

                        EprintCustomReport eprintCustomReport = new EprintCustomReport();
                        EprintCustomReport.EprintStatusCodes eprintStatusCode = new EprintCustomReport.EprintStatusCodes();
                        FileStream fileStream = File.Open(inputFile, FileMode.Open);
                        eprintStatusCode = eprintCustomReport.BeginReport(fileStream, fundOrgInput, outputfile, "Verbose");

                        if (fileStream != null)
                        {
                            fileStream.Dispose();
                        }

                        Console.WriteLine("Eprint Status Code:{0}", eprintStatusCode.ToString());
                        break;
                    // urlgenerate
                    case ClassInArgument.UrlGenerate:
                        Console.WriteLine("Program Class=" + programClass);
                        UrlGenerator urlGenerator = new UrlGenerator();
                        urlGenerator.BuildUrl(args[1], args[2], args[3], args[4]);
                        break;
                    // pagedownload
                    case ClassInArgument.PageDownload:
                        PageDownload.URLfetch(args[1]);
                        break;
                    // tochartest
                    case ClassInArgument.ToCharTest:
                        ToCharTest.StringToChar(args);
                        break;
                    // propertysetusingstruct
                    case ClassInArgument.PropertySetUsingStruct:
                        if (args.Length < 4)
                        {
                            throw new Exception(string.Format("Invalid argument count; requires 4 total ({0} were passed)", args.Length));
                        }
                        PropertySetUsingStruct propertySet = new PropertySetUsingStruct();
                        ListOfValues listOfValues = new ListOfValues
                        {
                            Email = args[1],
                            Name = args[2],
                            Number = args[3]
                        };
                        propertySet.ReadValues(listOfValues);
                        break;
                    // propertysetusingclass
                    case ClassInArgument.PropertySetUsingClass:
                        PropertySetUsingClass propertySetUsingClass = new PropertySetUsingClass
                        {
                            email = args[1],
                            name = args[2],
                            number = args[3]
                        };
                        propertySetUsingClass.ReadValues();
                        break;
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
