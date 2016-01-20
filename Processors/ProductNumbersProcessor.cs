using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Radar.Models.Criteria;
using Radar.Models.Product;
using ImportPOC2.Models;

namespace ImportPOC2.Processors
{
    public class ProductNumbersProcessor
    {
        private CriteriaProcessor _criteriaProcessor;
        private PriceProcessor _priceProcessor;
        private Product _currentProduct;
        private List<ProductNumber> existingProductNumbers;
        private List<ProductNumbersMap> productNumbersMap;

        public ProductNumbersProcessor(CriteriaProcessor critriaProcessor, PriceProcessor priceProcessor, Product currentProduct)
        {
            _criteriaProcessor = critriaProcessor;
            _priceProcessor = priceProcessor;
            _currentProduct = currentProduct;
            existingProductNumbers = new List<ProductNumber>();

            if (_currentProduct.ProductNumbers != null)
                existingProductNumbers = _currentProduct.ProductNumbers.ToList();

            productNumbersMap = new List<ProductNumbersMap>();
        }

        public void ProcessProductNumberRow(ProductRow sheetRow)
        {
            List<CriteriaSetValue> csValueList = null;

            if (!(string.IsNullOrWhiteSpace(sheetRow.Product_Number_Criteria_1)) || (!string.IsNullOrWhiteSpace(sheetRow.Product_Number_Criteria_2))
                && (!string.IsNullOrWhiteSpace(sheetRow.Product_Number_Other)))
            {
                if (!string.IsNullOrWhiteSpace(sheetRow.Product_Number_Criteria_1))                
                    csValueList = _criteriaProcessor.GetValuesBySheetSpecification(sheetRow.Product_Number_Criteria_1);                

                if (!string.IsNullOrWhiteSpace(sheetRow.Product_Number_Criteria_2))                
                    csValueList.AddRange(_criteriaProcessor.GetValuesBySheetSpecification(sheetRow.Product_Number_Criteria_2));

                var newProdNumbersMap = new ProductNumbersMap
                {
                    ProdNo = sheetRow.Product_Number_Other,
                    ProdNumberConfig = csValueList.Select(v => new ProductNumberConfiguration { CriteriaSetValueId = v.ID }).ToList()
                };

                productNumbersMap.Add(newProdNumbersMap);

                if (csValueList.Any())
                {
                    var csValueIds = csValueList.Select(v => v.ID).ToList();
                    var foundProdNum = findProductNumber(csValueIds);
                    if (foundProdNum == null)
                    {
                        //create new
                        createNewProductNumber(sheetRow.Product_Number_Other, csValueIds);
                    }
                    else
                    {
                        //update existing
                        foundProdNum.Value = sheetRow.Product_Number_Other;
                    }
                }
            }
        }

        private void createNewProductNumber(string prodNumberValue, List<long> csValueIds)
        {
            //create new product number
            var prodNumber = new ProductNumber();
            prodNumber.ID = Utils.IdGenerator.getNextid();
            prodNumber.ProductId = _currentProduct.ID;
            prodNumber.Value = prodNumberValue;

            prodNumber.ProductNumberConfigurations = new List<ProductNumberConfiguration>();

            //create product number configuration
            foreach (var id in csValueIds)
            {
                var prodNumberConfig = new ProductNumberConfiguration();
                prodNumberConfig.ID = Utils.IdGenerator.getNextid();
                prodNumberConfig.ProductNumberId = prodNumber.ID;
                prodNumberConfig.CriteriaSetValueId = id;

                prodNumber.ProductNumberConfigurations.Add(prodNumberConfig);
            }

            _currentProduct.ProductNumbers.Add(prodNumber);
        }

        private ProductNumber findProductNumber(List<long> csValueIds)
        {            
            var retVal = _currentProduct.ProductNumbers.FirstOrDefault(p => p.PriceGridId == null && p.ProductNumberConfigurations.Count() == csValueIds.Count() && 
                p.ProductNumberConfigurations.All(c => csValueIds.Contains(c.CriteriaSetValueId)));            

            return retVal;
        }

        public void FinalizeProductNumbers()
        {
            //delete product numbers that are missing from the initial state of product numbers collection
            existingProductNumbers.ForEach(num =>
            {
                var csValueIds = num.ProductNumberConfigurations.Select(n => n.CriteriaSetValueId).ToList();
                var foundProdNum = productNumbersMap.FirstOrDefault(n => n.ProdNumberConfig.Count() == csValueIds.Count() && 
                    n.ProdNumberConfig.All(c => csValueIds.Contains(c.CriteriaSetValueId)));

                if (foundProdNum == null)
                    _currentProduct.ProductNumbers.Remove(num);
            });                      
        }
    }
}
