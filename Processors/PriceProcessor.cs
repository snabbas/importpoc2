using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ImportPOC2.Models;
using Radar.Models.Pricing;
using Radar.Models.Product;

namespace ImportPOC2.Processors
{
    public class PriceProcessor
    {
        private CriteriaProcessor _criteriaProcessor;
        public List<PriceGridMap> PriceGridMaps;
        private int _baseGridCount;
        private int _upchargeGridCount;

        public PriceProcessor(CriteriaProcessor criteriaProcessor)
        {
            _criteriaProcessor = criteriaProcessor;
            PriceGridMaps = new List<PriceGridMap>();
        }

        public void ProcessPriceRow(ProductRow sheetRow, Product productModel)
        {
            //price grids are uniquely identified by criteria 
            //therefore build criteria definition processor
            //use results to determine what grid to consider for updates
            //keep track of grids from sheet vs. grids from product model for removal when all sheet grids have been processed
            if (!string.IsNullOrWhiteSpace(sheetRow.Base_Price_Name))
            {
                //add a base price grid to our map collection
                var newMap = new PriceGridMap
                {
                    IsBasePrice =  true,
                    CriteriaList = buildCriteriaList(sheetRow.Base_Price_Criteria_1, sheetRow.Base_Price_Criteria_2),
                    GridName =  sheetRow.Base_Price_Name
                };
                PriceGridMaps.Add(newMap);

                PriceGrid curGrid = null;
                if (string.IsNullOrWhiteSpace(sheetRow.Base_Price_Criteria_1))
                {
                    //we need to look for grid by name since there is no criteria
                    curGrid = productModel.PriceGrids.FirstOrDefault(g => g.IsBasePrice && g.Description == sheetRow.Base_Price_Name);
                }
                else
                {
                    var csvalueids = _criteriaProcessor.GetValuesBySheetSpecification(sheetRow.Base_Price_Criteria_1);
                    if (!string.IsNullOrWhiteSpace(sheetRow.Base_Price_Criteria_2))
                    {
                        csvalueids.AddRange(_criteriaProcessor.GetValuesBySheetSpecification(sheetRow.Base_Price_Criteria_2));
                    }

                    if (csvalueids.Any())
                    {
                        //var matchingGrids = productModel.PriceGrids.Where(g => g.PricingItems.Any(i => csvalueids.Contains(i.CriteriaSetValueId)));
                        int x = 0;
                    }

                }

                if (curGrid != null)
                {
                    //found it by name... anything special to do here? 
                    
                }
                else
                {
                    //assume that it is new here. 
                    curGrid = new PriceGrid
                    {
                        //Currency = "USD", //TODO: need default currency for product not a constant here
                        Description = sheetRow.Base_Price_Name,
                        IsBasePrice = true,
                        ID = Utils.IdGenerator.getNextid(), 
                        Prices = new Collection<Price>(),
                        PricingItems = new Collection<PricingItem>(),
                    };


                    productModel.PriceGrids.Add(curGrid);
                }
                curGrid.PriceIncludes = sheetRow.Price_Includes;
                curGrid.DisplaySequence = _baseGridCount++;
                curGrid.IsQUR = string.Equals(sheetRow.QUR_Flag, "Y", StringComparison.InvariantCultureIgnoreCase);

                fillBasePricesFromSheet(curGrid.Prices, sheetRow);

            }
            if (!string.IsNullOrWhiteSpace(sheetRow.Upcharge_Name))
            {
                var newMap = new PriceGridMap
                {
                    IsBasePrice =  false,
                    CriteriaList = buildCriteriaList(sheetRow.Upcharge_Criteria_1, sheetRow.Upcharge_Criteria_2),
                    GridName = sheetRow.Upcharge_Name
                };
                PriceGridMaps.Add(newMap);
            }

        }


        private void fillBasePricesFromSheet(ICollection<Price> collection, ProductRow sheetRow)
        {
            //TODO: using sheetrow, build a price object, then test against collection for add/change
            // does it make sense to use collection length and # of sheet cols to determine new? 
            var newPrice = createPrice(sheetRow.Q1, sheetRow.P1, sheetRow.D1);
            newPrice = createPrice(sheetRow.Q2, sheetRow.P2, sheetRow.D2);
            newPrice = createPrice(sheetRow.Q3, sheetRow.P3, sheetRow.D3);
            newPrice = createPrice(sheetRow.Q4, sheetRow.P4, sheetRow.D4);
            newPrice = createPrice(sheetRow.Q5, sheetRow.P5, sheetRow.D5);
            newPrice = createPrice(sheetRow.Q6, sheetRow.P6, sheetRow.D6);
            newPrice = createPrice(sheetRow.Q7, sheetRow.P7, sheetRow.D7);
            newPrice = createPrice(sheetRow.Q8, sheetRow.P8, sheetRow.D8);
            newPrice = createPrice(sheetRow.Q9, sheetRow.P9, sheetRow.D9);
            newPrice = createPrice(sheetRow.Q10, sheetRow.P10, sheetRow.D10);
        }

        private Price createPrice(string qty, string price, string discountCode)
        {
            Price newPrice = null;
            if (!string.IsNullOrWhiteSpace(qty) && !string.IsNullOrWhiteSpace(price) )//&& !string.IsNullOrWhiteSpace(discountCode))
            {
                int parsedQty = 0;
                Int32.TryParse(qty, out parsedQty);

                decimal parsedPrice = 0;
                decimal.TryParse(price, out parsedPrice);

                newPrice = new Price
                {
                    Quantity = parsedQty,
                    ListPrice = parsedPrice,
                    DiscountRate = getDiscountByCode(discountCode)
                };
            }

            return newPrice;
        }

        //lookup discount object by code
        private DiscountRate getDiscountByCode(string discountCode)
        {
            var retVal = Lookups.DiscountRates.FirstOrDefault(r => r.IndustryDiscountCode == "Z");

            if (!string.IsNullOrWhiteSpace(discountCode))
            {
                var existing = Lookups.DiscountRates.FirstOrDefault(r => r.IndustryDiscountCode == discountCode);

                if (existing != null)
                {
                    retVal = existing;
                }
            }

            return retVal;
        }

        private void fillUpchargePricesFromSheet(ICollection<Price> collection, ProductRow sheetRow)
        {
            //test  

        }

        private List<string> buildCriteriaList(string p1, string p2)
        {
            var retVal = new List<string>();

            if (!string.IsNullOrWhiteSpace(p1))
                retVal.Add(p1);

            if (!string.IsNullOrWhiteSpace(p2))
                retVal.Add(p2);

            return retVal;
        }

        /// <summary>
        /// invoke once the product has completed all processing (i.e., all rows have been read)
        /// this will perform final pricing actions, such as removing price grids from the product that 
        /// were not specified in the sheet.
        /// </summary>
        public void FinalizeProductPricing()
        {
            //compare price grid map to current product price grid list
            int x = 0;
            _baseGridCount = 0;
            _upchargeGridCount = 0;
        }
    }
}
