using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportPOC2.Processors
{
    public static class BasicStringFieldProcessor
    {
        /// <summary>
        /// Returns updated value of string field, based upon following rules:
        /// 1) if text is empty, no update occurs
        /// 2) if text is literial "NULL", field is emptied of its value
        /// 3) otherwise, new value is returned.
        /// </summary>
        /// <param name="newValue">updated value of string field</param>
        /// <param name="origValue">original value of string field</param>
        /// <returns>string</returns>
        public static string UpdateField(string newValue, string origValue)
        {
            string retVal = origValue;

            if (!string.IsNullOrWhiteSpace(newValue))
            {
                retVal = (newValue == "NULL" ? string.Empty : newValue);
            }
            return retVal;
        }
    }
}
