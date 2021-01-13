using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace CLI_Test
{
    public class EprintCustomReport
    {
        private enum SearchType
        {
            single,
            range,
            list
        }

        private enum FundOrg
        {
            FUND,
            ORG,
            Unknown
        }

        private enum DebugLevel
        {
            Verbose,
            Info,
            Warning,
            Error
        }

        public enum EprintStatusCodes
        {
            FileCreated,
            MissingFundOrg,
            MissingSearchValues,
            CriteriaNotFound
        }

        // console output level
        private static DebugLevel enumDebugLevel = DebugLevel.Error;
        // lines read from input
        readonly StringBuilder sbReadLines = new StringBuilder();
        // lines captured for custom report
        StringBuilder sbReportLines = new StringBuilder();
        SearchType enumSearchType = SearchType.single;
        FundOrg enumFundOrg = new FundOrg();
        // values provided within input string
        List<int> intSearchList = new List<int>();

        public EprintStatusCodes BeginReport(Stream inputStream, string fundOrgInput, string outputFile, string debugLevel)
        {
            enumDebugLevel = (DebugLevel)Enum.Parse(typeof(DebugLevel), debugLevel);
            // check that the input contains either FUND or ORG
            if (!FundOrgChoice(fundOrgInput, out enumFundOrg))
                return EprintStatusCodes.MissingFundOrg;

            // strip the 'FUND:' or 'ORG:' reference from the input string
            fundOrgInput = fundOrgInput.ToUpper().Replace(enumFundOrg.ToString() + ":", "");

            // build the search list from the string input
            if (!CreateSearchList(fundOrgInput, out intSearchList))
                return EprintStatusCodes.MissingSearchValues;

            // read each line of the system's version of the report
            using (var sr = new StreamReader(inputStream))
            {
                while (sr.Peek() >= 0)
                {
                    string line = sr.ReadLine();
                    //int accountIndex = -1;
                    int readUntil = 9;
                    int foundValue = -1;

                    // check if the line begins with a newpage
                    // the new page indicator has been converted to a hexString due to Studio not able to export the raw character
                    // convert the string-based line into bytes and compare it with the byte converted from the raw character that indicates a new page
                    byte[] lineBytes = Encoding.Default.GetBytes(line);
                    string hexString = BitConverter.ToString(lineBytes);

                    if (hexString.StartsWith("0C"))
                    {
                        //string filterString = "FUND:" + fundOrgSearch;
                        string filterFunctionString = enumFundOrg.ToString() + ":";
                        //string lineSearch = sbReadLines.ToString().Replace(" ", "");
                        string lineSearch = sbReadLines.ToString();
                        // get the index value of where '\r\nFUND:' or '\r\nORG:' is
                        int idx = lineSearch.IndexOf(Environment.NewLine + filterFunctionString);
                        // special case ORG
                        if (idx == -1)
                        {
                            switch (enumFundOrg)
                            {
                                // ORG may have a space between it and colon (:)
                                case FundOrg.ORG:
                                    filterFunctionString = enumFundOrg.ToString() + " :";
                                    idx = lineSearch.IndexOf(Environment.NewLine + filterFunctionString);
                                    break;
                                // GRANT can be an alias for FUND
                                case FundOrg.FUND:
                                    filterFunctionString = "GRANT:";
                                    //idx = lineSearch.IndexOf(filterFunctionString);
                                    break;
                            }
                        }

                        // typical filter was not found
                        if (idx == -1)
                        {
                            // special additional work for FUND
                            if (enumFundOrg == FundOrg.FUND)
                            {
                                int accountIndex = -1;

                                foreach (string accountseek in sbReadLines.ToString().Split(Environment.NewLine.ToCharArray()))
                                {
                                    if (accountIndex == -1)
                                    {
                                        // some reports reference ACCOUNT/ as an alias for FUND
                                        if (accountseek.Contains("ACCOUNT/"))
                                        {
                                            accountIndex = accountseek.IndexOf("ACCOUNT/");
                                        } // check to see if ACCOUNT exists and is not something like 'ACCOUNT TITLE' or 'ACCOUNT ACTIVITY'
                                        else if (accountseek.Contains("ACCOUNT   "))
                                        {
                                            accountIndex = accountseek.IndexOf("ACCOUNT   ");
                                            Console.WriteLine("accountIndex for 'ACCOUNT   '=" + accountIndex);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("AccountSeek={0} and AccountIndex={1}", accountseek, accountIndex);
                                        //if (accountseek.Length < (accountIndex + 9))
                                        //{
                                         //   readUntil = accountseek.Length - accountIndex;
                                        //}
                                        Console.WriteLine("readUntil=" + readUntil);
                                        //if (readUntil < 0)
                                        //    continue;
                                        if (accountseek.Length - accountIndex <= readUntil)
                                            continue;

                                        string accountParsedVal = accountseek.Substring(accountIndex, readUntil).Replace(" ", "");
                                        Console.WriteLine("accountParsedVal='{0}'", accountParsedVal.Length);

                                        try
                                        {
                                            int accountValue = Convert.ToInt32(
                                                Regex.Replace(
                                                    accountParsedVal,
                                                    "[^0-9]",
                                                    ""
                                                    )
                                            );

                                            if (intSearchList.Contains(accountValue))
                                            {
                                                foundValue = accountValue;
                                                //Console.WriteLine("Parsed Account Value Found in Search List:" + accountValue.ToString());
                                                break;
                                            }
                                            else if (enumSearchType == SearchType.range) // are the input values a range?
                                            {
                                                if (accountValue >= intSearchList[0] && accountValue <= intSearchList[1])
                                                {
                                                    //Console.WriteLine("ACCOUNT Value is in range");
                                                    foundValue = accountValue;
                                                    break;
                                                }
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            ConsoleOutput(string.Format("Value Parsed:{0}", accountParsedVal), DebugLevel.Warning);
                                            ConsoleOutput(string.Format("No number parsed:{0}", e.ToString()), DebugLevel.Warning);
                                        }
                                    }
                                }
                            }
                        }
                        else // filter was found
                        {
                            if ((idx + filterFunctionString.Length) < (idx + filterFunctionString.Length + readUntil))
                            {
                                // pass the starting position after
                                // get the value found in between empty spaces
                                string spacedVal = "";
                                if (filterFunctionString.Contains("GRANT"))
                                {
                                    spacedVal = ValueBetweenSpaces(lineSearch, idx + filterFunctionString.Length);
                                }
                                else
                                {
                                    spacedVal = ValueBetweenSpaces(lineSearch, idx + Environment.NewLine.Length + filterFunctionString.Length);
                                }

                                // attempt to parse the value that exists between empty spaces
                                if (!int.TryParse(spacedVal, out foundValue))
                                {
                                    // attempt to parse the value after FUND: or ORG: constants in report
                                    string parsedVal = lineSearch.Substring(idx + filterFunctionString.Length, readUntil);
                                    foundValue = Convert.ToInt32(
                                        Regex.Replace(
                                            parsedVal,
                                            "[^0-9]",
                                            ""
                                            )
                                        );
                                }
                            }
                        }

                        // add lines to the stringbuilder for the custom report based on the value type (single, list, range)
                        AppendLinesByType(enumSearchType, foundValue, ref sbReportLines);
                        // clear the loop stringbuilder
                        sbReadLines.Clear();
                    }
                    sbReadLines.Append(line + Environment.NewLine);
                }
            }

            // create the file containing all of the custom report lines
            return WriteToFile(outputFile, sbReportLines);
        }

        private void ConsoleOutput(string consoleOutput, DebugLevel debugLevel)
        {
            int currentLevel = (int)enumDebugLevel;
            int requiredLevel = (int)debugLevel;

            if(requiredLevel >= currentLevel)
            {
                Console.WriteLine(consoleOutput);
            }
        }

        private string ValueBetweenSpaces(string lineSearch, int idxStart)
        {
            // the value found in between empty spaces
            string spacedVal = "";
            //for (int i = idx + filterFunctionString.Length; i >= idx + filterFunctionString.Length; i++)
            for (int i = idxStart; i >= idxStart; i++)
            {
                char readChar = Convert.ToChar(lineSearch.Substring(i, 1));
                // the read character is white space
                if (char.IsWhiteSpace(readChar))
                {
                    // the value is not null or empty
                    if (!string.IsNullOrEmpty(spacedVal))
                    {
                        // we got the value we were looking for between spaces
                        break;
                    }
                }
                else
                {
                    spacedVal += readChar;
                }
            }

            return spacedVal;
        }

        private EprintStatusCodes WriteToFile(string outputFile, StringBuilder sbReportLines)
        {
            // write to file
            if (sbReportLines.Length > 0)
            {
                File.AppendAllText(outputFile, sbReportLines.ToString());
                return EprintStatusCodes.FileCreated;
            }
            else
            {
                ConsoleOutput("The criteria was not found in the selected report.", DebugLevel.Warning);
                return EprintStatusCodes.CriteriaNotFound;
            }
        }

        private void AppendLinesByType(SearchType enumSearchType, int foundValue, ref StringBuilder sbReportLines)
        {
            switch (enumSearchType)
            {
                case SearchType.single:
                    if (intSearchList.Contains(foundValue))
                    {
                        ConsoleOutput("Value is in input", DebugLevel.Verbose);
                        _ = sbReportLines.Append(sbReadLines + Environment.NewLine);
                    }

                    break;
                case SearchType.list:
                    if (intSearchList.Contains(foundValue))
                    {
                        ConsoleOutput("Value is in input list", DebugLevel.Verbose);
                        _ = sbReportLines.Append(sbReadLines + Environment.NewLine);
                    }

                    break;
                case SearchType.range:
                    if (foundValue >= intSearchList[0] && foundValue <= intSearchList[1])
                    {
                        ConsoleOutput("Value is in range", DebugLevel.Verbose);
                        _ = sbReportLines.Append(sbReadLines + Environment.NewLine);
                    }
                    break;
            }
        }

        private Boolean CreateSearchList(string fundOrgInput, out List<int> intSearchList)
        {
            intSearchList = new List<int>();

            if (fundOrgInput.Contains("-"))
            {
                enumSearchType = SearchType.range;
                string[] strArr = fundOrgInput.Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
                int[] arr = Array.ConvertAll(strArr, s => int.Parse(s));

                if (arr[0] > arr[1])
                {
                    ConsoleOutput("Beginning value is less than the ending value.", DebugLevel.Error);
                    return false;
                }

                ConsoleOutput("Value is a range.", DebugLevel.Info);
                intSearchList.AddRange(arr);
            }
            else if (fundOrgInput.Contains(","))
            {
                enumSearchType = SearchType.list;

                string[] strArr = fundOrgInput.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                int[] arr = Array.ConvertAll(strArr, s => int.Parse(s));
                ConsoleOutput("Values are a set list.", DebugLevel.Info);
                intSearchList.AddRange(arr);
            }
            else
            {
                ConsoleOutput("Value is handled as a single value.", DebugLevel.Info);
                intSearchList.Add(int.Parse(fundOrgInput));
            }

            return true;
        }

        private Boolean FundOrgChoice(string fundOrgInput, out FundOrg enumFundOrg)
        {
            if (fundOrgInput.ToUpper().Contains("FUND:"))
            {
                enumFundOrg = FundOrg.FUND;
                ConsoleOutput("FUND will be searched.", DebugLevel.Info);
                return true;
            }
            else if (fundOrgInput.ToUpper().Contains("ORG:"))
            {
                enumFundOrg = FundOrg.ORG;
                ConsoleOutput("ORG will be searched.", DebugLevel.Info);
                return true;
            }

            ConsoleOutput("Criteria has not been specified.", DebugLevel.Info);
            enumFundOrg = FundOrg.Unknown;
            return false;
        }
    }
}
