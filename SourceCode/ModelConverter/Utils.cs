using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelConverter
{
    class Utils
    {
        public static string UppercaseWords(string text, char sperateCode)
        {
            string[] tmpString = text.Split(sperateCode);
            string returnString = "";
            foreach (var s in tmpString)
            {
                if (s.Length > 0)
                    returnString += Char.ToUpper(s[0]) + s.Substring(1).ToLower();
            }
            return returnString;
        }

        public static string GetDefaultStringValue(string text)
        {
            if (text.Length == 0)
                return "";
            string returnString = text.Replace("(N'", "").Replace("')", "").Replace("((", "").Replace("))", "");
            return returnString;
        }

        public static string GetDefaultNumberValue(string text)
        {
            if (text.Length == 0)
                return "0";
            string returnString = text.Replace("(N'", "").Replace("')", "").Replace("((", "").Replace("))", "");
            return returnString;
        }
    }
}
