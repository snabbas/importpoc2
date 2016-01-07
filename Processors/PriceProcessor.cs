using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportPOC2.Models;
using Radar.Models.Pricing;

namespace ImportPOC2.Processors
{
    public class PriceProcessor
    {
        private CriteriaProcessor _criteriaProcessor;
        public List<PriceGridMap> PriceGridMaps;
        private int _baseGridCount;
        private int _upchargeGridCount;
        private int globalId; //TODO: this needs to be stored in global object instead

        public PriceProcessor()
        {
            _criteriaProcessor = new CriteriaProcessor();
            PriceGridMaps = new List<PriceGridMap>();
        }

        public void ProcessPriceRow(ProductRow sheetRow, Radar.Models.Product.Product productModel)
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

                //todo: really this should be by criteria, but note that criteria could be empty for single base price grid
                var curGrid = productModel.PriceGrids.FirstOrDefault(g => g.Description == sheetRow.Base_Price_Name);
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
                        ID = globalId--, 
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

        }

        private void fillUpchargePricesFromSheet(ICollection<Price> collection, ProductRow sheetRow)
        {

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
        public void Finalize()
        {
            //compare price grid map to current product price grid list
            int x = 0;
            _baseGridCount = 0;
            _upchargeGridCount = 0;
        }
    }
}
