using System;
using System.Text;
using System.Security.Cryptography;

namespace CLI_Test
{
    class UrlGenerator
    {
        public void BuildUrl(string formUrl, string token, string fieldName, string cwid)
        {
            StringBuilder param = new StringBuilder();
            param.Append("&");
            param.Append(Uri.EscapeDataString("ufpre" + fieldName));
            param.Append("=");
            param.Append(Uri.EscapeDataString(cwid));

            string uriString = param.ToString();

            Console.WriteLine("String before hashing:" + uriString);

            //Calculate the hash
            byte[] tokenBytes = Convert.FromBase64String(token);

            byte[] calculatedHash = null;
            using (HMACSHA256 hmac = new HMACSHA256(tokenBytes))
            {
                calculatedHash = hmac.ComputeHash(Encoding.UTF8.GetBytes(uriString));
            }
            string hash = Convert.ToBase64String(calculatedHash);

            Console.WriteLine("String after hashing:" + hash);

            //Create the final URL

            StringBuilder finalUrl = new StringBuilder();
            finalUrl.Append(formUrl);
            finalUrl.Append(uriString);
            finalUrl.Append("&");
            finalUrl.Append(Uri.EscapeDataString("ufprehash"));
            finalUrl.Append("=");
            finalUrl.Append(Uri.EscapeDataString(hash));

            Console.WriteLine(finalUrl);
        }
    }
}
