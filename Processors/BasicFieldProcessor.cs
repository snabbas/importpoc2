using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportPOC2.Processors
{
    public static class BasicFieldProcessor
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
            var retVal = origValue;

            if (!string.IsNullOrWhiteSpace(newValue))
            {
                retVal = (newValue == "NULL" ? string.Empty : newValue);
            }
            return retVal;
        }

        public static bool UpdateField(string newValue, bool origValue)
        {
            var retVal = origValue;

            if (!string.IsNullOrWhiteSpace(newValue))
            {
                retVal = (newValue.ToLower() == "y");
            }
            return retVal;
        }

        public static DateTime? UpdateField(string newValue, DateTime? origValue)
        {
            var retVal = origValue;

            if (!string.IsNullOrWhiteSpace(newValue))
            {
               if(newValue == "NULL")
                 retVal = null;
               else
                 retVal = Convert.ToDateTime(newValue);
            }
            return retVal;
        }

        internal static int? UpdateField(string newValue, int? origValue)
        {
            var retVal = origValue;

            if (!string.IsNullOrWhiteSpace(newValue))
            {
                int? tmp = Convert.ToInt32(newValue);
                retVal = (newValue == "NULL" ? null : tmp);
            }

            return retVal;
        }
    }
}
