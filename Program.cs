using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using ASI.Contracts.Stats;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportPOC2.Processors;
using Newtonsoft.Json;
using Radar.Core.Models.Batch;
using Radar.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using Radar.Models;
using Radar.Models.Criteria;
using Radar.Models.Product;
using ProductKeyword = Radar.Models.Product.ProductKeyword;
using ProductMediaItem = Radar.Models.Product.ProductMediaItem;
[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace ImportPOC2
{
    class Program
    {
        private static SharedStringTablePart _stringTable;
        private static List<string> _sheetColumnsList;
        private static IQueryable<Radar.Core.Models.Import.Template> _mapping;
        private static string _curXid;
        private static int _companyId;
        private static Batch _curBatch;
        private static UowPRODTask _prodTask;
        private static readonly HttpClient RadarHttpClient = new HttpClient {BaseAddress = new Uri("http://local-espupdates.asicentral.com/api/api/")};
        private static Product _currentProduct;
        private static bool _firstRowForProduct = true;
        private static int globalUniqueId = 0;

        private static List<Category> _catlist = null;
        private static List<Category> CategoryList
        {
            get
            {
                if (_catlist == null)
                {

                    var results = RadarHttpClient.GetAsync("lookup/product_categories").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _catlist = JsonConvert.DeserializeObject<List<Category>>(content);
                    }

                }
                return _catlist;
            }
            set { _catlist = value; }
        }

        private static List<ProductColorGroup> _colorGroupList = null;

        private static List<ProductColorGroup> colorGroupList
        {
            get
            {
                if (_colorGroupList == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/colors").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _colorGroupList = JsonConvert.DeserializeObject<List<ProductColorGroup>>(content);
                    }
                }
                return _colorGroupList;
            }
            set { _colorGroupList = value; }
        }

        private static log4net.ILog _log;
        private static bool _hasErrors = false;


        static void Main(string[] args)
        {
            _log = log4net.LogManager.GetLogger(typeof(Program));

            //onetime stuff
            RadarHttpClient.DefaultRequestHeaders.Accept.Clear();
            RadarHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            //get directory from config
            //TODO: change this to read from config

            //NOTES: 
            //file name is batch ID - use to retreive batch to assign details. 
            // also use to determine company ID of sheet. 

            var curDir = Directory.GetCurrentDirectory();

            _log.InfoFormat("running in {0}", curDir);

            var xlsxFiles = Directory.GetFiles(curDir, "*.xlsx");

            if (!xlsxFiles.Any())
            {
                _log.InfoFormat("No xlsx files found in {0}", curDir);
            }
            else
            {
                //go get column positions from DB. 
                _prodTask = new UowPRODTask();
                _mapping = _prodTask.Template.GetAllWithInclude(t => t.TemplateMapping)
                    .Where(t => t.AuditStatusCode == Radar.Core.Common.Constants.StatusCode.Audit.ACTIVE)
                    .Select(t => t);

                foreach (var xlsxFile in xlsxFiles)
                {
                    processFile(xlsxFile);
                }
            }

            // import only supports XLSX files... right? 

            Console.Write("Press <enter> key to exit");
            Console.ReadLine();
        }

        private static void processFile(string xlsxFile)
        {
            _log.InfoFormat("processing {0}", Path.GetFileName(xlsxFile));
            //processor: 

            //initizations
            //get batch: 
            var batchId = getBatchId(xlsxFile);
            _curBatch = getBatchById(batchId);

            if (_curBatch == null)
            {
                _log.ErrorFormat("unable to find batch {0}", batchId);
            }
            else
            {
                //open file
                using (var document = SpreadsheetDocument.Open(xlsxFile, false))
                {
                    //set up some refs
                    var workbookPart = document.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    _stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    var headerValidated = false;
                    _curXid = string.Empty;

                    foreach (var row in sheetData.Elements<Row>())
                    {
                        if (!headerValidated)
                        {
                            // validate header
                            headerValidated = validateHeader(row);
                            //todo: we need to bail if header is not validated
                            if (!headerValidated)
                            {
                                logit("Unknown sheet format - cannot process");
                                break;
                            }
                        }
                        else
                        {
                            processDataRow(row);
                            _firstRowForProduct = false; 
                        }
                    }

                    //processing: 
                    // map column positions to field names for this sheet
                    // 
                    //loop:
                    //get next row 
                    //determine row is same product or changed
                    // if changed get product based on XID/COmpanyId
                    // if new, create empty Radar product (viewmodel)
                    // loop through columns
                    //use colum position in map to call into function for each field for processing
                    // if column is NULL, ?? remove/update? 
                    // if column is blank, ?? remove/update? 
                    // if column has data, update viewmodel appropriately
                }
            }
        }

        private static Batch getBatchById(long batchId)
        {   
            var batch = _prodTask.Batch.GetAllWithInclude(
                t => t.BatchDataSources,
                t => t.BatchErrorLogs,
                t => t.BatchProducts)
                .FirstOrDefault(b => b.BatchId == batchId);

            if (batch != null && batch.CompanyId.HasValue)
            {
                _companyId = (int)batch.CompanyId.Value;
            }

            return batch;

        }

        private static long getBatchId(string xlsxFile)
        {
            long retVal = 0;
            //using file name, extract batch ID 
            var filename = Path.GetFileNameWithoutExtension(xlsxFile);
            if (filename != null)
            {
                retVal = long.Parse(filename);
            }

            return retVal;
        }

        //at this point we "know" the columns match a specified format;
        //here we "walk" through each column and handle appropriately
        private static void processDataRow(Row row)
        {
            foreach (var column in row.Elements<Cell>())
            {
                var text = getCellText(column);
                var curColIndex = getColIndex(column);
                if (curColIndex == 0 && string.IsNullOrWhiteSpace(text))
                {
                    //row is invalid without XID, bail out
                    //logit("empty row encountered, skipping");
                    break;
                }
                if (curColIndex == 0 && _curXid != text)
                {
                    //XID has changed, it's a new product
                    finishProduct();
                    _curXid = text;
                    startProduct();
                }
                else
                {
                    try
                    {
                        //it's a column on the current product, process it. 
                        processColumn(curColIndex, text);
                    }
                    catch (Exception exc)
                    {
                        _hasErrors = true; 
                        _log.Error("Unhandled exception occurred:",exc);
                    }
                }

                //colIndex++;
                //no longer incrementing this as getColIndex function determines where we are in the sheet
                //some columns are skipped if they don't have data. 
            }
        }

        /// <summary>
        /// converts an excel column reference value (e.g., "A", "AZ", "Q") into column index.
        /// </summary>
        /// <param name="column"></param>
        /// <returns>column index as integer</returns>
        private static int getColIndex(Cell column)
        {
            var retVal = 0;

            var colName = Regex.Replace(column.CellReference.Value, "[0-9]*", "");

            foreach (var c in colName)
            {
                retVal *= 26;
                retVal += (c - 'A' + 1);
            }
            retVal--; //need zero-based index

            return retVal;
        }

        private static void startProduct()
        {
            if (!string.IsNullOrWhiteSpace(_curXid))
            {
                //using current XID, check if product exists, otherwise create new empty model 
                _currentProduct = getProductByXid() ?? new Radar.Models.Product.Product { CompanyId = _companyId };
                _firstRowForProduct = true;
                _hasErrors = false;
            }
        }

        //TODO: currently only able to do this via product import controller endpoint
        //therefore, refactor radar somehow to expose this via core/data/? 
        private static Radar.Models.Product.Product getProductByXid()
        {
            Radar.Models.Product.Product retVal = null;

            //TODO: read environment from CoNFIG
            var endpointUrl = string.Format("productimport?externalProductId={0}&companyId={1}", _curXid, _companyId);

            try
            {
                var results = RadarHttpClient.GetAsync(endpointUrl).Result;

                if (results.IsSuccessStatusCode)
                {
                    var content = results.Content.ReadAsStringAsync().Result;
                    retVal = JsonConvert.DeserializeObject<Product>(content);
                }
                else
                {
                    logit(string.Format("Unable to retreive product xid:{0} for companyid: {1} reason:{2}", _curXid, _companyId, results.StatusCode));
                }
            }
            catch (Exception exc)
            {
                //something bad happened.
                logit(string.Format("Error querying product {0}/{1}:\r\n{2}", _curXid, _companyId, exc.Message));
            }

            return retVal;
        }

        private static void processColumn(int colIndex, string text)
        {
            //map the current column 
            var colName = _sheetColumnsList.ElementAt(colIndex);
            switch (colName)
            {
                case "XID":
                    //shouldn't be anything to do here
                    break;
                case "Product_Name":
                    processProductName(text);
                    break;
                case "Product_Number":
                    processProductNumber(text);
                    break;

                case "Description":
                    processDescription(text);
                    break;

                case "Summary":
                    processSummary(text);
                    break;

                case "Prod_Image":
                    processImage(text);
                    break;

                case "Category":
                    processCategory(text);
                    break;

                case "Keywords":
                    processKeywords(text);
                    break;

                case "Inventory_Link":
                    processInventoryLink(text);
                    break;

                case "Product_Color":
                    processColor(text);
                    break;

                case "Material":
                    processMaterial(text);
                    break;

                case "Size_Group":
                    processSizeGroup(text);
                    break;

                case "Size_Values":
                    processSizeValues(text);
                    break;

                case "Additional_Color":
                    break;
                case "Additional_Info":
                    break;
                case "Additional_Location":
                    break;
                case "Artwork":
                    break;
                case "Breakout_by_other_attribute":
                    break;
                case "Breakout_by_price":
                    break;
                case "Can_order_only_one":
                    break;
                case "Catalog_Information":
                    break;
                case "Comp_Cert":
                    break;
                case "Confirmed_Thru_Date":
                    break;
                case "Disclaimer":
                    break;
                case "Distibutor_Only":
                    break;
                case "Distributor_View_Only":
                    break;
                case "Dont_Make_Active":
                    break;
                case "Imprint_Color":
                    break;
                case "Imprint_Location":
                    break;
                case "Imprint_Method":
                    break;
                case "Imprint_Size":
                    break;
                case "Inventory_Quantity":
                    break;
                case "Inventory_Status":
                    break;
                case "Less_Than_Min":
                    break;
                case "Linename":
                    break;
                case "Option_Additional_Info":
                    break;
                case "Option_Name":
                    break;
                case "Option_Type":
                    break;
                case "Option_Values":
                    break;
                case "Origin":
                    break;
                case "Packaging":
                    break;
                case "Personalization":
                    break;
                case "Product_Data_Sheet":
                    break;
                case "Product_Inventory_Link":
                    break;
                case "Product_Inventory_Quantity":
                    break;
                case "Product_Inventory_Status":
                    break;
                case "Product_Number_Criteria_1":
                    break;
                case "Product_Number_Criteria_2":
                    break;
                case "Product_Number_Other":
                    break;
                case "Product_Number_Price":
                    break;
                case "Product_Sample":
                    break;
                case "Product_SKU":
                    break;
                case "Production_Time":
                    break;
                case "Req_for_order":
                    break;
                case "Rush_Service":
                    break;
                case "Rush_Time":
                    break;
                case "Safety_Warnings":
                    break;
                case "Same_Day_Service":
                    break;
                case "SEO_FLG":
                    break;
                case "Shape":
                    break;
                case "Ship_Plain_Box":
                    break;
                case "Shipper_Bills_By":
                    break;
                case "Shipping_Dimensions":
                    break;
                case "Shipping_Info":
                    break;
                case "Shipping_Items":
                    break;
                case "Shipping_Weight":
                    break;
                case "SKU":
                    break;
                case "SKU_Based_On":
                    break;
                case "SKU_Criteria_1":
                    break;
                case "SKU_Criteria_2":
                    break;
                case "SKU_Criteria_3":
                    break;
                case "SKU_Criteria_4":
                    break;
                case "Sold_Unimprinted":
                    break;
                case "Spec_Sample":
                    break;
                case "Theme":
                    break;
                case "Tradename":
                    break;
                /* pricing fields */
                case "Price_Includes":
                    break;
                case "Price_Type":
                    break;
                case "Currency":
                    break;
                case "Base_Price_Criteria_1":
                    break;
                case "Base_Price_Criteria_2":
                    break;
                case "Base_Price_Name":
                    break;
                case "QUR_Flag":
                    break;
                case "D1":
                case "D10":
                case "D2":
                case "D3":
                case "D4":
                case "D5":
                case "D6":
                case "D7":
                case "D8":
                case "D9":
                    handleDiscountCode(text, colName);
                    break;
                case "P1":
                case "P10":
                case "P2":
                case "P3":
                case "P4":
                case "P5":
                case "P6":
                case "P7":
                case "P8":
                case "P9":
                    handleBasePrice(text, colName);
                    break;
                case "Q1":
                case "Q10":
                case "Q2":
                case "Q3":
                case "Q4":
                case "Q5":
                case "Q6":
                case "Q7":
                case "Q8":
                case "Q9":
                    handleBaseQty(text, colName);
                    break;
                /* upcharge fields */
                case "Upcharge_Criteria_1":
                    break;
                case "Upcharge_Criteria_2":
                    break;
                case "Upcharge_Details":
                    break;
                case "Upcharge_Level":
                    break;
                case "Upcharge_Name":
                    break;
                case "Upcharge_Type":
                    break;
                case "UD1":
                case "UD10":
                case "UD2":
                case "UD3":
                case "UD4":
                case "UD5":
                case "UD6":
                case "UD7":
                case "UD8":
                case "UD9":
                    handleUpchargeDiscount(text, colName);
                    break;
                case "UP1":
                case "UP10":
                case "UP2":
                case "UP3":
                case "UP4":
                case "UP5":
                case "UP6":
                case "UP7":
                case "UP8":
                case "UP9":
                    handleUpchargePrice(text, colName);
                    break;
                case "UQ1":
                case "UQ10":
                case "UQ2":
                case "UQ3":
                case "UQ4":
                case "UQ5":
                case "UQ6":
                case "UQ7":
                case "UQ8":
                case "UQ9":
                    handleUpchargeQty(text, colName);
                    break;
                case "U_QUR_Flag":
                    break;
            }
        }

        private static void processMaterial(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processSizeGroup(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processSizeValues(string text)
        {
            //throw new NotImplementedException();
        }

        private static void handleUpchargeQty(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleUpchargePrice(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleUpchargeDiscount(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleBaseQty(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleBasePrice(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleDiscountCode(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void processColor(string text)
        {
            //colors are comma delimited 
            if (_firstRowForProduct)
            {
                var colorList = extractCsvList(text);
                colorList.ForEach(c =>
                {
                    //each color is in format of colorName=alias
                    //TODO: COMBO COLORS
                    string colorName;
                    string aliasName;
                    var colorWithAlias = c.Split('=');
                    if (colorWithAlias.Length > 1)
                    {
                        colorName = colorWithAlias[0];
                        aliasName = colorWithAlias[1];
                    }
                    else
                    {
                        colorName = c;
                        aliasName = c;
                    }
                    // if colorname isn't recognized, then it gets "UNCLASSIFIED/other" grouping
                    var productColors = getCriteriaSetValuesByCode("PRCL");
                    var colorObj = colorGroupList.SelectMany(g => g.CodeValueGroups).FirstOrDefault(g => String.Equals(g.Description, colorName, StringComparison.CurrentCultureIgnoreCase));
                    var existing = productColors.FirstOrDefault(p => p.Value == aliasName);

                    long setCodeId = 0;
                    if (colorObj == null)
                    {
                        //they picked a color that doesn't exist, so we choose the "other" set code to assign it on the new value
                        colorObj = colorGroupList.SelectMany(g => g.CodeValueGroups).FirstOrDefault(g => string.Equals(g.Description, "Unclassified/Other", StringComparison.CurrentCultureIgnoreCase));
                        if (colorObj != null)
                        {
                            setCodeId = colorObj.SetCodeValues.First().Id;
                        }
                    }
                    else
                    {
                        setCodeId = colorObj.SetCodeValues.First().Id;
                    }

                    if (existing == null)
                    {
                        //needs to be added
                        createNewValue("PRCL", aliasName, setCodeId);
                    }
                    else
                    {
                        //update existing if its different colorname?
                        //updateValue(existing, setCodeId);
                    }
                });

                //remove colors from product here. 
            }
        }

        private static void createNewValue(string criteriaCode, string aliasName, long setCodeValueId, string optionName = "")
        {
            var cSet = getCriteriaSetByCode(criteriaCode, optionName);

            if (!cSet.Any())
            {
                
            }
        }

        private static IEnumerable<CriteriaSetValue> getCriteriaSetValuesByCode(string criteriaCode, string optionName = "")
        {
            var result = new List<CriteriaSetValue>();

            var cSet = getCriteriaSetByCode(criteriaCode, optionName);
            result = cSet.SelectMany(c => c.CriteriaSetValues).ToList();

            return result;
        }

        private static IEnumerable<ProductCriteriaSet> getCriteriaSetByCode(string criteriaCode, string optionName = "")
        {
            var retVal = new List<ProductCriteriaSet>();
            var prodConfig = _currentProduct.ProductConfigurations.FirstOrDefault(c => c.IsDefault);

            if (prodConfig != null)
            {
                retVal = prodConfig.ProductCriteriaSets.Where(c => c.CriteriaCode == criteriaCode).ToList();
            }

            return retVal;
        }

        private static List<string> extractCsvList(string text)
        {
            //returns list of strings from input string, split on commas, each value is trimmed, and only non-empty values are returned.
            return text.Split(',').Select(str => str.Trim()).Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
        }

        private static void processKeywords(string text)
        {
            //comma delimited list of keywords - only "visible" keywords, never ad or seo keywords
            if (_firstRowForProduct)
            {
                var keywords = extractCsvList(text);
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
                        var newKeyword = new ProductKeyword {Value = keyword, TypeCode = "HIDD", ID = globalUniqueId--};
                        _currentProduct.ProductKeywords.Add(newKeyword);
                    }
                });
                //now select any product keywords that are not in the sheet's list, and remove them
                var toRemove = _currentProduct.ProductKeywords.Where(p => !keywords.Contains(p.Value)).ToList();
                toRemove.ForEach(r => _currentProduct.ProductKeywords.Remove(r));
            }
        }

        private static void processCategory(string text)
        {
            if (_firstRowForProduct)
            {
                var categories = extractCsvList(text);

                //just in case it's totally empty/null
                if (_currentProduct.SelectedProductCategories == null)
                    _currentProduct.SelectedProductCategories = new Collection<ProductCategory>();

                categories.ForEach(curCat =>
                {
                    //need to lookup categories
                    var category = CategoryList.FirstOrDefault(c => c.Name == curCat);
                    if (category != null)
                    {

                        var existing = _currentProduct.SelectedProductCategories.FirstOrDefault(c => c.Code == category.Code);
                        if (existing == null)
                        {
                            var newCat = new ProductCategory {Code = category.Code, AdCategoryFlg = false};
                            _currentProduct.SelectedProductCategories.Add(newCat);
                        }
                        else
                        {
                            //s/b nothing to do? 
                        }
                    }
                });

                //remove any categories from product that aren't on the sheet; get list of codes from the sheet 
                var sheetCategoryList = categories.Join(CategoryList, cat => cat, lookup => lookup.Name, (cat, lookup) => lookup.Code);
                var toRemove = _currentProduct.SelectedProductCategories.Where(c => !sheetCategoryList.Contains(c.Code)).ToList();
                toRemove.ForEach(r => _currentProduct.SelectedProductCategories.Remove(r));
            }
        }

        private static void processImage(string text)
        {
            if (_firstRowForProduct)
            {
                //text here should be a list of comma sepearated URLs, in order of display
                var urls = extractCsvList(text);
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
                            var m = new Media {Url = currentUrl};
                            var pm = new ProductMediaItem {Media = m, MediaRank = curUrlCount++};
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

        private static void processInventoryLink(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.ProductLevelInventoryLink = BasicStringFieldProcessor.UpdateField(text);
        }

        private static void processSummary(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Summary = BasicStringFieldProcessor.UpdateField(text);
        }

        private static void processDescription(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Description = BasicStringFieldProcessor.UpdateField(text);
        }

        private static void processProductNumber(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AsiProdNo = BasicStringFieldProcessor.UpdateField(text);
        }

        private static void processProductName(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Name = BasicStringFieldProcessor.UpdateField(text);
        }

        ///// <summary>
        ///// Returns updated value of string field, based upon following rules:
        ///// 1) if text is empty, no update occurs
        ///// 2) if text is literial "NULL", field is emptied of its value
        ///// 3) otherwise, new value is returned.
        ///// </summary>
        ///// <param name="newValue"></param>
        ///// <returns>string</returns>
        //private static string updateField(string newValue)
        //{
        //    string retVal = newValue;

        //    if (!string.IsNullOrWhiteSpace(newValue))
        //    {
        //        retVal = (newValue == "NULL" ? string.Empty : newValue);
        //    }
        //    return retVal;
        //}

        private static void finishProduct()
        {
            //if we've started a radar model, 
            // we "send" the product to Radar for processing. 
            if (_currentProduct != null && !_hasErrors)
            {
                var x = _currentProduct;
            }
        }


        private static bool validateHeader(Row row)
        {
            var retVal = true;

            //get list of columns from this header row
            _sheetColumnsList = getColumnsFromSheet(row);


            //TODO: how to have these settings "sent in" here

            var columnMetaDataPriceV1 = getColumnsByFormatVersion("ASIS", "V1");
            var columnMetaDataFullV1 = getColumnsByFormatVersion("ASIF", "V1");
            var columnMetaDataPriceV2 = getColumnsByFormatVersion("ASIS", "V2");
            var columnMetaDataFullV2 = getColumnsByFormatVersion("ASIF", "V2");

            retVal = compareColumns(columnMetaDataFullV2);

            if (retVal)
            {
                _log.InfoFormat("Sheet is Full V2 format.");
            }
            else
            {
                retVal = compareColumns(columnMetaDataPriceV2);

                if (retVal)
                {
                    _log.InfoFormat("Sheet is Price V2 format.");
                }
                else
                {
                    retVal = compareColumns(columnMetaDataFullV1);

                    if (retVal)
                    {
                        _log.InfoFormat("Sheet is Full V1 format.");
                    }
                    else
                    {
                        retVal = compareColumns(columnMetaDataPriceV1);
                        if (retVal)
                        {
                            _log.InfoFormat("Sheet is Price V1 format.");
                        }
                        else
                        {
                            _log.Warn("Sheet is UNKNOWN format - cannot process");
                        }
                    }
                }
            }
            //override for testing
            //retVal = true; 

            return retVal;
        }

        private static bool compareColumns(List<string> columnMetaData)
        {
            //test by combining lists where the names match, removing quotes (for csv we needed quote removal, openxml should remove them)
            //stole this from current barista logic, but note that ZIP uses "first" list length to know when to stop comparing. 
            // this means additional columns in second list are ignored.

            var test = columnMetaData
               .Zip(_sheetColumnsList, (a, b) => a.Replace("\"\"", "\"").Trim('\"') == b ? 1 : 0)
               .Select((a, i) => new { Index = i, Value = a }).ToArray();

            //TODO: log which value(s) didn't match to error log? maybe outside of here instead. 
            
            return test.All(a => a.Value == 1);
        }

        private static List<string> getColumnsFromSheet(Row row)
        {
            return row.Elements<Cell>().Select(getCellText).ToList();
        }

        private static List<string> getColumnsByFormatVersion(string p1, string p2)
        {
            // we need list of column names, in order by sequence
            //price import v1 is ASIS 0.0.1
            //full v1 is ASIF 0.0.5
            //both v2 2.0.0, ASIS/ASIF price/ful
            var templateCode = p1;
            var version = (p2 == "V2" ? "2.0.0" : p1=="ASIS" ? "0.0.1" : "0.0.5");

            var retVal = _mapping.Where(t => t.TemplateCode == templateCode && t.Version == version)
                .SelectMany(t => t.TemplateMapping)
                .OrderBy(m => m.SeqNo)
                .Select(m => m.SourceField).ToList();
            
            // return this list
            return retVal;
        }

        /// <summary>
        /// Cell text can be stored directly or in a shared "string table" at the sheet level
        /// this method decodes the magic bits that returns the actual data in the specified cell object
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        private static string getCellText(Cell c)
        {
            string text = null;
            if (c.CellValue != null)
            {
                text = c.CellValue.InnerText;
                if (c.DataType != null)
                {
                    switch (c.DataType.Value)
                    {
                        case CellValues.SharedString:

                            if (_stringTable != null)
                            {
                                text = _stringTable.SharedStringTable.ElementAt(int.Parse(text)).InnerText;
                            }
                            break;
                        //case CellValues.String:
                        //    text = c.CellValue.InnerText;
                        case CellValues.Boolean:
                            switch (text)
                            {
                                case "0":
                                    text = "False";
                                    break;
                                default:
                                    text = "True";
                                    break;
                            }

                            break;
                    }
                }
            }
            return text;
        }


        private static void logit(string message)
        {

            _log.Debug(message);
        }
    }
}
