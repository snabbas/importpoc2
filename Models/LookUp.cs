using System.Collections.Generic;
using Radar.Models.Criteria;

namespace ImportPOC2
{
    public class CodeValueLookUp
    {
        public string Code { get; set; }
        public string Value { get; set; }
    }

    public class CodeDescriptionLookUp
    {
        public string Code { get; set; }
        public string Description { get; set; }
    }


    public class GenericLookUp
    {
        public string CodeValue { get; set; }
        public long? ID { get; set; }
        //this property will be used in case of sizes which has 10 different size types
        public string CriteriaCode { get; set; } 
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

    public class ImprintCriteriaLookUp
    {
        public ImprintCriteriaLookUp()
        {
            CodeValueGroups = new List<CodeValueGroup>();
        }

        public string Code { get; set; }
        public ICollection<CodeValueGroup> CodeValueGroups { get; set; }        
    }
  
    public class GenericIdLookup
    {
        public string CriteriaCode { get; set; }
        public int CriteriaAttributeId { get; set; }
        public long CustomSetCodeValueId { get; set; }
    }
}
