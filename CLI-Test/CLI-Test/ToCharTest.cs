using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace CLI_Test
{
    class ToCharTest
    {
        public static void StringToChar(string[] sa)
        {
            foreach(string s in sa.Skip(1))
            {
                Console.WriteLine("The string input is '{0}'", s);
                if (char.TryParse(Regex.Unescape(s), out char c))
                {
                    Console.WriteLine("The char value is '{0}'", c);
                }
                else
                {
                    Console.WriteLine("Unable to parse string to char '{0}'", s);
                }

                Console.WriteLine("**********************");
            }
        }
    }
}
