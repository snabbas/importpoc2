using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportPOC2.Models
{
    public class PriceGridMap
    {
        public bool IsBasePrice { get; set; }
        public string GridName { get; set; }
        public long TargetGridId { get; set; }
        public List<string> CriteriaList { get; set; }
    }
}
