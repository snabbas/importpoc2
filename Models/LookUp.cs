using System.Collections.Generic;
using Radar.Models.Criteria;

namespace ImportPOC2
{
    public class CodeValueLookUp
    {
        public string Code { get; set; }
        public string Value { get; set; }
    }

    public class GenericLookUp
    {
        public string CodeValue { get; set; }
        public long? ID { get; set; }
    }

    public class KeyValueLookUp
    {
        public long Key { get; set; }
        public string Value { get; set; }
    }

    public class CurrencyLookUp
    {
        public int Number { get; set; }
        public string Code { get; set; }
    }

    public class SafetyWarningLookUp
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public class CostTypeLookUp
    {
        public string DisplayName { get; set; }
        public string Code { get; set; }
    }  

    public class ThemeLookUp
    {
        public ThemeLookUp()
        {
            SetCodeValues = new List<SetCodeValue>(); 
        }

        public string Code { get; set; }
        public ICollection<SetCodeValue> SetCodeValues { get; set; }
    }
}
