using System.Collections.Generic;
using Radar.Models.Pricing;

namespace ImportPOC2.Models
{
    public class PriceGridMap
    {
        public bool IsBasePrice { get; set; }
        public string GridName { get; set; }
        public long TargetGridId { get; set; }
        public List<string> CriteriaList { get; set; }
        public List<Price> Prices { get; set; }
        public List<PricingItem> PricingItems { get; set; }
    }
}
