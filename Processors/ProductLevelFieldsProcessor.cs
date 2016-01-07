using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using ImportPOC2.Models;
using Radar.Models.Product;
using Radar.Models;
using ImportPOC2.Utils;
using Radar.Core.Models.Batch;

namespace ImportPOC2.Processors
{
    public class ProductLevelFieldsProcessor
    {
        private readonly ProductRow _currentProductRow = null;
        private readonly bool _firstRowForProduct = false;
        private Product _currentProduct = null;
        private bool _publishCurrentProduct = false;
        private static int _globalUniqueId = 0;
        private Batch _currentBatch = null;

        public ProductLevelFieldsProcessor(ProductRow currentProductRow, bool firstRowForProduct, Product currentProduct, bool publishCurrentProduct, Batch currentBatch)
        {
            _currentProductRow = currentProductRow;
            _firstRowForProduct = firstRowForProduct;
            _currentProduct = currentProduct;
            _publishCurrentProduct = publishCurrentProduct;
            _currentBatch = currentBatch;
        }

        public void ProcessProductLevelFields()
        {            
            processProductName(_currentProductRow.Product_Name);
            processProductNumber(_currentProductRow.Product_Number);
            processProductSku(_currentProductRow.Product_SKU);
            processInventoryLink(_currentProductRow.Product_Inventory_Link);
            processInventoryStatus(_currentProductRow.Product_Inventory_Status);
            processInventoryQty(_currentProductRow.Product_Inventory_Quantity);
            processDescription(_currentProductRow.Description);
            processSummary(_currentProductRow.Summary);
            processImage(_currentProductRow.Prod_Image);
            processCategory(_currentProductRow.Category);
            processKeywords(_currentProductRow.Keywords);
            processAdditionalShippingInfo(_currentProductRow.Shipping_Info);
            processAdditionalProductInfo(_currentProductRow.Additional_Info);
            processDistributorOnlyViewFlag(_currentProductRow.Distributor_View_Only);
            processDistributorOnlyComment(_currentProductRow.Distibutor_Only);
            processProductDisclaimer(_currentProductRow.Disclaimer);
            processCurrency(_currentProductRow.Currency);
            processLessThanMinimum(_currentProductRow.Less_Than_Min);
            processPriceType(_currentProductRow.Price_Type);
            processProductDataSheet(_currentProductRow.Product_Data_Sheet);
            ////breakout price - not processed
            processConfirmationDate(_currentProductRow.Confirmed_Thru_Date);
            processDontMakeActive(_currentProductRow.Dont_Make_Active);
            //breakout by attribute -- not processed
            //seo flag -- not processed
        }
       
        private void processProductName(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Name = BasicFieldProcessor.UpdateField(text, _currentProduct.Name);
        }

        private void processPriceType(string text)
        {
            if (_firstRowForProduct)
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    text = "List";
                }

                var priceTypeFound = Lookups.CostTypesLookup.FirstOrDefault(t => t.Code == text);

                if (priceTypeFound != null)
                {
                    _currentProduct.CostTypeCode = BasicFieldProcessor.UpdateField(text, _currentProduct.CostTypeCode);
                }
                else
                {
                    //log batch error 
                    //Validation.AddValidationError(_currentBatch, "", text, _currentProduct.ID, _currentProduct.ExternalProductId);                    
                    //_hasErrors = true;
                }
            }
        }

        private void processCurrency(string text)
        {
            if (_firstRowForProduct)
            {
                if (_currentProduct.PriceGrids.Count > 0)
                {
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        text = "USD";
                    }

                    var currencyFound = Lookups.CurrencyLookup.FirstOrDefault(t => t.Code == text);

                    if (currencyFound != null)
                    {
                        foreach (var priceGrid in _currentProduct.PriceGrids.Where(priceGrid => priceGrid.Currency.Code != currencyFound.Code))
                        {
                            priceGrid.Currency = new Radar.Models.Pricing.Currency
                            {
                                Code = currencyFound.Code,
                                Number = currencyFound.Number
                            };
                        }
                    }
                    else
                    {
                        //log batch error 
                        //Validation.AddValidationError(_currentBatch, "", text, _currentProduct.ID, _currentProduct.ExternalProductId);         
                        //_hasErrors = true;
                    }
                }
            }
        }

        private void processInventoryStatus(string text)
        {
            if (_firstRowForProduct)
            {
                var inventoryStatusFound = Lookups.InventoryStatusesLookup.FirstOrDefault(t => t.Value == text);
                if (inventoryStatusFound != null)
                {
                    _currentProduct.ProductLevelInventoryStatusCode = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductLevelInventoryStatusCode);
                }
                else
                {
                    //Validation.AddValidationError(_currentBatch, "", text, _currentProduct.ID, _currentProduct.ExternalProductId);         
                    //_hasErrors = true;
                }
            }
        }

        private void processProductSku(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.ProductLevelSku = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductLevelSku);
        }

        private void processInventoryQty(string text)
        {
            if (_firstRowForProduct)
            {
                _currentProduct.ProductLevelInventoryQuantity = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductLevelInventoryQuantity);
            }
        }

        private void processDistributorOnlyViewFlag(string text)
        {
            //throw new NotImplementedException();
        }

        private void processDistributorOnlyComment(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.DistributorComments = BasicFieldProcessor.UpdateField(text, _currentProduct.DistributorComments);
        }

        private void processProductDisclaimer(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Disclaimer = BasicFieldProcessor.UpdateField(text, _currentProduct.Disclaimer);
        }

        private void processConfirmationDate(string text)
        {
            if (_firstRowForProduct)
            {
                _currentProduct.PriceConfirmationDate = BasicFieldProcessor.UpdateField(text, _currentProduct.PriceConfirmationDate);
            }
        }

        private void processAdditionalProductInfo(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AdditionalInfo = BasicFieldProcessor.UpdateField(text, _currentProduct.AdditionalInfo);
        }

        private void processAdditionalShippingInfo(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AddtionalShippingInfo = BasicFieldProcessor.UpdateField(text, _currentProduct.AddtionalShippingInfo);
        }

        private void processKeywords(string text)
        {
            //comma delimited list of keywords - only "visible" keywords, never ad or seo keywords
            if (_firstRowForProduct)
            {
                var keywords = text.ConvertToList();
                if (_currentProduct.ProductKeywords == null)
                {
                    _currentProduct.ProductKeywords = new Collection<ProductKeyword>();
                }
                keywords.ForEach(keyword =>
                {
                    var keyObj = _currentProduct.ProductKeywords.FirstOrDefault(p => p.Value == keyword);
                    if (keyObj == null)
                    {
                        //need to add it
                        var newKeyword = new ProductKeyword { Value = keyword, TypeCode = "HIDD", ID = _globalUniqueId-- };
                        _currentProduct.ProductKeywords.Add(newKeyword);
                    }
                });
                //now select any product keywords that are not in the sheet's list, and remove them
                var toRemove = _currentProduct.ProductKeywords.Where(p => !keywords.Contains(p.Value)).ToList();
                toRemove.ForEach(r => _currentProduct.ProductKeywords.Remove(r));
            }
        }

        private void processCategory(string text)
        {
            if (_firstRowForProduct)
            {
                var categories = text.ConvertToList();

                //just in case it's totally empty/null
                if (_currentProduct.SelectedProductCategories == null)
                    _currentProduct.SelectedProductCategories = new Collection<ProductCategory>();

                categories.ForEach(curCat =>
                {
                    //need to lookup categories
                    var category = Lookups.CategoryList.FirstOrDefault(c => c.Name == curCat);
                    if (category != null)
                    {

                        var existing = _currentProduct.SelectedProductCategories.FirstOrDefault(c => c.Code == category.Code);
                        if (existing == null)
                        {
                            var newCat = new ProductCategory { Code = category.Code, AdCategoryFlg = false };
                            _currentProduct.SelectedProductCategories.Add(newCat);
                        }
                        else
                        {
                            //s/b nothing to do? 
                        }
                    }
                });

                //remove any categories from product that aren't on the sheet; get list of codes from the sheet 
                var sheetCategoryList = categories.Join(Lookups.CategoryList, cat => cat, lookup => lookup.Name, (cat, lookup) => lookup.Code);
                var toRemove = _currentProduct.SelectedProductCategories.Where(c => !sheetCategoryList.Contains(c.Code)).ToList();
                toRemove.ForEach(r => _currentProduct.SelectedProductCategories.Remove(r));
            }
        }

        private void processImage(string text)
        {
            if (_firstRowForProduct)
            {
                //text here should be a list of comma sepearated URLs, in order of display
                var urls = text.ConvertToList();

                var curUrlCount = 1;
                urls.ForEach(currentUrl =>
                {
                    if (!string.IsNullOrWhiteSpace(currentUrl))
                    {
                        //note: no validation here, Radar already does all of that.
                        var curMedia = _currentProduct.ProductMediaItems.FirstOrDefault(m => m.Media.Url == currentUrl);
                        if (curMedia == null)
                        {
                            //add it
                            var m = new Media { Url = currentUrl };
                            var pm = new ProductMediaItem { Media = m, MediaRank = curUrlCount++ };
                            _currentProduct.ProductMediaItems.Add(pm);
                        }
                        else
                        {
                            curMedia.MediaRank = curUrlCount++;
                        }
                    }
                });
            }
        }

        private void processDontMakeActive(string text)
        {
            if (_firstRowForProduct)
            {
                if (!string.IsNullOrWhiteSpace(text) && text.ToLower() == "y")
                    _publishCurrentProduct = false;
            }
        }

        private void processInventoryLink(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.ProductLevelInventoryLink = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductLevelInventoryLink);
        }

        private void processSummary(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Summary = BasicFieldProcessor.UpdateField(text, _currentProduct.Summary);
        }

        private void processDescription(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Description = BasicFieldProcessor.UpdateField(text, _currentProduct.Description);
        }

        private void processProductNumber(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AsiProdNo = BasicFieldProcessor.UpdateField(text, _currentProduct.AsiProdNo);
        }

        private void processLessThanMinimum(string text)
        {
            if (_firstRowForProduct)
            {
                _currentProduct.IsOrderLessThanMinimumAllowed = BasicFieldProcessor.UpdateField(text, _currentProduct.IsOrderLessThanMinimumAllowed);
            }
        }

        private void processProductDataSheet(string text)
        {
            if (_firstRowForProduct)

                if (_currentProduct.ProductDataSheet == null)
                    _currentProduct.ProductDataSheet = new ProductDataSheet();

            _currentProduct.ProductDataSheet.Url = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductDataSheet.Url);
        }
    }
}
