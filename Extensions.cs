using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ImportPOC2
{
    public static class Extensions
    {
        public static List<string> ConvertToList(this string value)
        {
            var lstValues = new List<string>();
            if (!string.IsNullOrWhiteSpace(value))
            {
                var result = Regex.Matches(value, @"\w+([-:;=+*\s]+\w+)*|\[.*?\]").Cast<Match>().Select(m => m.Value).ToArray();
                char[] charsToTrim = {'[', ']'};
                foreach (string str in result)
                {
                    var trimStr = str;
                    trimStr = trimStr.Trim(charsToTrim);
                    if (!string.IsNullOrWhiteSpace(trimStr))
                    {
                        lstValues.Add(trimStr);
                    }
                }
            }
            return lstValues;
        }

        public static FieldInfo SplitValue(this string value, char separator)
        {
            FieldInfo retVal = new FieldInfo(); 
            if (!string.IsNullOrWhiteSpace(value))
            {
                var tokens = value.Split(separator);

                if (tokens.Length == 2)
                {
                    retVal.CodeValue = tokens[0];
                    retVal.AdditionalInfo = tokens[1];
                }
                else
                {
                    retVal.CodeValue = tokens[0];
                    retVal.AdditionalInfo = string.Empty;
                }
            }

            return retVal;
        }
    }

    public class FieldInfo
    {
        public string CodeValue { get; set; }
        public string AdditionalInfo { get; set; }
    }
}


