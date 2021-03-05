using System;
using System.IO;
using System.Data;

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
            PropertySetUsingClass,
            OracleDatabaseAdapter,
            Encryptonite,
            PdfRender
        }

        public static void ArgCheck(string[] args, int argCount)
        {
            // disregard the count of the class argument
            int argLength = args.Length - 1;
            if(argLength != argCount)
            {
                throw new Exception(string.Format("Argument length is shorter than required ({0})", argCount + 1));
            }
        }

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    Array knownArgs = Enum.GetValues(typeof(ClassInArgument));
                    foreach(object knownArg in knownArgs)
                    {
                        Console.WriteLine(knownArg.ToString());
                    }
                    throw new Exception("No arguments passed");
                }

                // which class to run against
                string programClass = args[0];
                ClassInArgument classesInArgument = new ClassInArgument();
                if(!Enum.TryParse(programClass, true, out classesInArgument))
                    Console.WriteLine("Invalid class argument passed:" + programClass);

                switch (classesInArgument)
                {
                    //eprint
                    case ClassInArgument.Eprint:
                        Program.ArgCheck(args, 5);

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
                    case ClassInArgument.OracleDatabaseAdapter:
                        OracleDatabaseAdapter oracleDatabaseAdapter = new OracleDatabaseAdapter();
                        oracleDatabaseAdapter.OracleQuery(args[1], null, out DataTable dataTable);

                        if (dataTable.Rows.Count > 0)
                        {
                            int i = 0;
                            foreach (DataRow row in dataTable.Rows)
                            {
                                i++;

                                foreach (DataColumn col in dataTable.Columns)
                                {
                                    Console.WriteLine(String.Format("RowNum:{0}" + Environment.NewLine + "Column:{1}" + Environment.NewLine + "Value:{2}",
                                                                                                        i.ToString(), col.ColumnName, row[col].ToString()), "Output");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No rows were returned");
                        }
                        dataTable.Dispose();
                        break;
                    case ClassInArgument.Encryptonite:
                        Program.ArgCheck(args, 3);
                        if(string.Equals(args[1], "Encrypt", StringComparison.OrdinalIgnoreCase))
                        {
                            Encryptonite encryptonite = new Encryptonite(args[2]);
                            string encryptedStr = encryptonite.Encrypt(args[3], true);
                            Console.WriteLine(args[2]);
                            Console.WriteLine(args[3]);
                            Console.WriteLine(encryptedStr);
                        }
                        else if (string.Equals(args[1], "Decrypt", StringComparison.OrdinalIgnoreCase))
                        {
                            Encryptonite encryptonite = new Encryptonite(args[2]);
                            string encryptedStr = encryptonite.Decrypt(args[3], true);
                            Console.WriteLine(args[2]);
                            Console.WriteLine(args[3]);
                            Console.WriteLine(encryptedStr);
                        }
                        break;
                    case ClassInArgument.PdfRender:
                        Program.ArgCheck(args, 1);
                        PdfRender pdfRender = new PdfRender();
                        if(string.Equals(args[1], "facultystaffpdf", StringComparison.OrdinalIgnoreCase))
                        {
                            pdfRender.FacultyStaffPDF();
                        }
                        else
                        {
                            pdfRender.GenerateTest(args[1]);
                        }
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
