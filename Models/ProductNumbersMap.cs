using System.Collections.Generic;
using Radar.Models.Product;

namespace ImportPOC2.Models
{
    public class ProductNumbersMap
    {
        public ProductNumbersMap()
        {
            ProdNumberConfig = new List<ProductNumberConfiguration>();
        }
        
        public string ProdNo { get; set; }
        public List<ProductNumberConfiguration> ProdNumberConfig { get; set; }        
    }
}
