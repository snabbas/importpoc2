using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportPOC2.Models;
using ImportPOC2.Utils;
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

            //start by building a map using sheet and product grids so we know what exists, what doesn't, etc.
            /* base prices */
            if (!string.IsNullOrWhiteSpace(sheetRow.Base_Price_Name))
            {
                //add a base price grid to our map collection
                var newMap = new PriceGridMap
                {
                    IsBasePrice =  true,
                    CriteriaList = buildCriteriaList(sheetRow.Base_Price_Criteria_1, sheetRow.Base_Price_Criteria_2),
                    GridName =  sheetRow.Base_Price_Name
                };

                //first, let's see if the grid already exists
                PriceGrid curGrid = null;
                if (string.IsNullOrWhiteSpace(sheetRow.Base_Price_Criteria_1))
                {
                    //we need to look for grid by name since there is no criteria
                    curGrid = productModel.PriceGrids.FirstOrDefault(g => g.IsBasePrice && g.Description == sheetRow.Base_Price_Name);
                }
                else
                {
                    //since we have criteria, we need to use it to find a grid
                    //pull the criteria values, and look for a grid that has the same exact value list. 
                    //otherwise we will treat it as a new grid. 
                    var csValueList = _criteriaProcessor.GetValuesBySheetSpecification(sheetRow.Base_Price_Criteria_1);
                    if (!string.IsNullOrWhiteSpace(sheetRow.Base_Price_Criteria_2))
                    {
                        csValueList.AddRange(_criteriaProcessor.GetValuesBySheetSpecification(sheetRow.Base_Price_Criteria_2));
                    }

                    //with the list of CSValue IDs, find a grid that matches them all
                    if (csValueList.Any())
                    {
                        var justIds = csValueList.Select(v => v.ID);
                        
                        var matchingGrids = productModel.PriceGrids.Where(g => g.IsBasePrice && g.PricingItems.All(i => justIds.Contains(i.CriteriaSetValueId))).ToList();

                        if (matchingGrids.Any())
                        {
                            curGrid = matchingGrids.First();
                        }
                    }
                }
                //if it exists, put the ID into our mapping table; otherwise use a unique ID to represent a new Grid. 
                newMap.TargetGridId = curGrid != null ? curGrid.ID : IdGenerator.getNextid();

                //hold onto the data in the model format so we can quickly build grids for the product as needed.
                fillBasePricesFromSheet(newMap.Prices, sheetRow);
                fillPricingItemsFromSheet(newMap.PricingItems, sheetRow);

                PriceGridMaps.Add(newMap);
            }

            /* upcharges */
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

        private void fillPricingItemsFromSheet(ICollection<PricingItem> collection, ProductRow sheetRow)
        {
            var csValueList = _criteriaProcessor.GetValuesBySheetSpecification(sheetRow.Base_Price_Criteria_1);

            foreach (var value in csValueList)
            {
                var item = new PricingItem
                {
                    CriteriaSetValueId = value.ID,
                    ID = IdGenerator.getNextid()
                    //TODO: what other item values do we need here?
                };
                collection.Add(item);
            }
        }

        private void fillBasePricesFromSheet(ICollection<Price> collection, ProductRow sheetRow)
        {
            var newPrice = createPrice(sheetRow.Q1, sheetRow.P1, sheetRow.D1);
            if (newPrice!=null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q2, sheetRow.P2, sheetRow.D2);
            if (newPrice != null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q3, sheetRow.P3, sheetRow.D3);
            if (newPrice != null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q4, sheetRow.P4, sheetRow.D4);
            if (newPrice != null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q5, sheetRow.P5, sheetRow.D5);
            if (newPrice != null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q6, sheetRow.P6, sheetRow.D6);
            if (newPrice != null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q7, sheetRow.P7, sheetRow.D7);
            if (newPrice != null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q8, sheetRow.P8, sheetRow.D8);
            if (newPrice != null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q9, sheetRow.P9, sheetRow.D9);
            if (newPrice != null)
                collection.Add(newPrice);
            newPrice = createPrice(sheetRow.Q10, sheetRow.P10, sheetRow.D10);
            if (newPrice != null)
                collection.Add(newPrice);
        }

        private Price createPrice(string qty, string price, string discountCode)
        {
            Price newPrice = null;
            if (!string.IsNullOrWhiteSpace(qty) && !string.IsNullOrWhiteSpace(price)) //&& !string.IsNullOrWhiteSpace(discountCode))
            {
                var parsedQty = 0;
                decimal parsedPrice = 0;

                if (Int32.TryParse(qty, out parsedQty) && decimal.TryParse(price, out parsedPrice))
                {
                    newPrice = new Price
                    {
                        Quantity = parsedQty,
                        ListPrice = parsedPrice,
                        DiscountRate = getDiscountByCode(discountCode)
                    };
                }
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
        public void FinalizeProductPricing(Product currentProduct)
        {
            //compare price grid map to current product price grid list

            var basePriceMap = PriceGridMaps.Where(m => m.IsBasePrice).ToList();
            var upchargeMap = PriceGridMaps.Where(m => !m.IsBasePrice).ToList();

            var productPrices = currentProduct.PriceGrids.Where(g => g.IsBasePrice).ToList();
            var productUpcharges = currentProduct.PriceGrids.Where(g => !g.IsBasePrice).ToList();

            //update base prices as needed

            //update upcharges as needed. 

            //if a grid (upcharges or base prices) is on the product but not in the map, it needs to be deleted
            var priceGridIdsToDelete = productPrices.Select(g => g.ID).Except(basePriceMap.Select(m => m.TargetGridId)).ToList();
            priceGridIdsToDelete.AddRange(productUpcharges.Select(g => g.ID).Except(upchargeMap.Select(m => m.TargetGridId)).ToList());
            
            priceGridIdsToDelete.ForEach(d =>
            {
                var existing = productPrices.FirstOrDefault(p => p.ID == d);
                currentProduct.PriceGrids.Remove(existing);
            });
            
        }
    }
}
