﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportPOC2.Processors;
using Newtonsoft.Json;
using Radar.Core.Models.Batch;
using Radar.Data;
using Radar.Models;
using Radar.Models.Criteria;
using Radar.Models.Product;
using Constants = Radar.Core.Common.Constants;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using ProductKeyword = Radar.Models.Product.ProductKeyword;
using ProductMediaItem = Radar.Models.Product.ProductMediaItem;
using Radar.Models.Company;
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

        private static List<KeyValueLookUp> _shapesLookup = null;
        private static List<KeyValueLookUp> shapesLookup
        {
            get
            {
                if (_shapesLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/shapes").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _shapesLookup = JsonConvert.DeserializeObject<List<KeyValueLookUp>>(content);
                    }
                }
                return _shapesLookup;
            }
            set { _shapesLookup = value; }
        }

        private static List<SetCodeValue> _themesLookup = null;
        private static List<SetCodeValue> themesLookup
        {
            get
            {
                if (_themesLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/themes").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        List<ThemeLookUp> themeGroups = JsonConvert.DeserializeObject<List<ThemeLookUp>>(content);
                        _themesLookup = new List<SetCodeValue>();

                        themeGroups.ForEach(t =>
                        {
                            _themesLookup.AddRange(t.SetCodeValues);
                        });
                    }
                }
                return _themesLookup;
            }
            set { _themesLookup = value; }
        }

        private static List<GenericLookUp> _originsLookup = null;
        private static List<GenericLookUp> originsLookup
        {
            get
            {
                if (_originsLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/origins").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _originsLookup = JsonConvert.DeserializeObject<List<GenericLookUp>>(content);
                    }
                }
                return _originsLookup;
            }
            set { _originsLookup = value; }
        }

        private static List<KeyValueLookUp> _packagingLookup = null;
        private static List<KeyValueLookUp> packagingLookup
        {
            get
            {
                if (_packagingLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/packaging").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _packagingLookup = JsonConvert.DeserializeObject<List<KeyValueLookUp>>(content);
                    }
                }
                return _packagingLookup;
            }
            set { _packagingLookup = value; }
        }

        private static List<ImprintCriteriaLookUp> _imprintCriteriaLookup = null;
        private static List<ImprintCriteriaLookUp> imprintCriteriaLookup
        {
            get
            {
                if (_imprintCriteriaLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/criteria?code=IMPR").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _imprintCriteriaLookup = JsonConvert.DeserializeObject<List<ImprintCriteriaLookUp>>(content);
                    }
                }
                return _imprintCriteriaLookup;
            }
            set { _imprintCriteriaLookup = value; }
        }

        private static List<LineName> _linenamesLookup = null;
        private static List<LineName> linenamesLookup
        {
            get
            {
                if (_linenamesLookup == null)
                {
                    long companyId = _currentProduct.CompanyId;
                    var results = RadarHttpClient.GetAsync("lookup/linenames?company_id=" + companyId).Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _linenamesLookup = JsonConvert.DeserializeObject<List<LineName>>(content);
                    }
                }
                return _linenamesLookup;
            }
            set { linenamesLookup = value; }
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
                    .Where(t => t.AuditStatusCode == Constants.StatusCode.Audit.ACTIVE)
                    .Select(t => t);

                foreach (var xlsxFile in xlsxFiles)
                {

                    //NOTE: xlsxFile is expected to be in format {batchId}.xlsx 
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
                long temp;
                if (long.TryParse(filename, out temp))
                {
                    retVal = temp;
                }
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
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            processColumn(curColIndex, text);
                        }
                    }
                    catch (Exception exc)
                    {
                        _hasErrors = true;
                        _log.Error("Unhandled exception occurred:", exc);
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
                _currentProduct = getProductByXid() ?? new Product { CompanyId = _companyId };
                _firstRowForProduct = true;
                _hasErrors = false;
            }
        }

        //TODO: currently only able to do this via product import controller endpoint
        //therefore, refactor radar somehow to expose this via core/data/? 
        private static Product getProductByXid()
        {
            Product retVal = null;

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
                    if (results.StatusCode == HttpStatusCode.NotFound)
                    {
                        _log.InfoFormat("Product XID {0} not found under companyID {1}; creating as new.", _curXid, _companyId);
                    }
                    else
                    {
                        _log.WarnFormat("Unable to retreive product xid:{0} for companyid: {1} reason:{2}", _curXid, _companyId, results.StatusCode);
                    }
                }
            }
            catch (Exception exc)
            {
                //something bad happened.
                _log.Error(string.Format("Error querying product XID {0} for companyId{1}", _curXid, _companyId), exc);
            }

            return retVal;
        }

        private static void processColumn(int colIndex, string text)
        {
            //map the current column 
            var colName = _sheetColumnsList.ElementAt(colIndex);
            switch (colName)
            {
                    /* product-level fields */
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

                case "Shipping_Info": //AKA additional shipping information 
                    processAdditionalShippingInfo(text);
                    break;
                case "Additional_Info"://AKA additional product information 
                    processAdditionalProductInfo(text);
                    break;
                case "Confirmed_Thru_Date":
                    processConfirmationDate(text);
                    break;
                case "Disclaimer":
                    processProductDisclaimer(text);
                    break;
                case "Distibutor_Only"://this is the comment 
                    processDistributorOnlyComment(text);
                    break;
                case "Distributor_View_Only"://this field sets IncludeAppOfferList
                    processDistributorOnlyViewFlag(text);
                    break;
                case "Product_Data_Sheet":
                    processProductDataSheet(text);
                    break;
                case "Catalog_Information": //format: catalogname:year:page
                    processCatalogInfo(text);
                    break;

                case "Dont_Make_Active":
                    //TODO: set flag to determine if publish should be attempted when POSTing
                    break;

                    /* product level SKU info */
                case "Product_Inventory_Link":
                    processInventoryLink(text);
                    break;
                case "Product_Inventory_Quantity":
                    processInventoryQty(text);
                    break;
                case "Product_Inventory_Status":
                    processInventoryStatus(text);
                    break;
                case "Product_SKU":
                    processProductSKU(text);
                    break;
                    /* product level SKU info end */

                    /* not processed by import even though the column is in the sheet*/
                case "Breakout_by_other_attribute":
                    break;
                case "Breakout_by_price":
                    break;
                case "SEO_FLG":
                    /* ignored on Imports */
                    break;
                /* not used */

                /* non-criteria set collections */
                case "Prod_Image":
                    processImage(text);
                    break;

                case "Category":
                    processCategory(text);
                    break;

                case "Keywords":
                    processKeywords(text);
                    break;

                case "Linename":
                    processLineNames(text);
                    break;

                case "Safety_Warnings":
                    processSafetyWarnings(text);
                    break;

                case "Comp_Cert":
                    processComplianceCertifications(text);
                    break;

                /* shipping info */
                    //these are individual columns as well, no need to validate bills by as Radar does that 
                case "Shipping_Dimensions":
                    break;
                case "Shipping_Items":
                    break;
                case "Shipping_Weight":
                    break;
                case "Shipper_Bills_By":
                    break;
                /* shipping info */

                    /* criteria sets */
                case "Product_Color":
                    processProductColors(text);
                    break;

                case "Material":
                    processMaterials(text);
                    break;

                case "Size_Group":
                    processSizeGroups(text);
                    break;

                case "Size_Values":
                    processSizeValues(text);
                    break;

                case "Additional_Color":
                    processAdditionalColors(text);
                    break;
                case "Additional_Location":
                    processAdditionalLocations(text);
                    break;
                case "Artwork":
                    processImprintArtwork(text);
                    break;
                case "Imprint_Color":
                    processImprintColors(text);
                    break;
                case "Imprint_Location":
                    processImprintLocations(text);
                    break;
                case "Imprint_Method":
                    processImprintMethods(text);
                    break;
                case "Imprint_Size":
                    processImprintSizes(text);
                    break;
                case "Less_Than_Min":
                    processLessThanMinimum(text);
                    break;
                case "Origin":
                    processOrigins(text);
                    break;
                case "Packaging":
                    processPackagingOptions(text);
                    break;
                case "Ship_Plain_Box":
                    break;
                case "Personalization":
                    break;
                case "Product_Sample":
                    break;
                case "Production_Time":
                    break;
                case "Rush_Service":
                    break;
                case "Rush_Time":
                    break;
                case "Same_Day_Service":
                    break;
                case "Shape":
                    processShapes(text);
                    break;
                case "Sold_Unimprinted":
                    break;
                case "Spec_Sample":
                    break;
                case "Theme":
                    processThemes(text);                    
                    break;
                case "Tradename":
                    processTradenames(text); 
                    break;

                    /* product numbers */
                    /* TODO: put individual fields into "prod num" object, then process at row change into product model */
                case "Product_Number_Criteria_1":
                    break;
                case "Product_Number_Criteria_2":
                    break;
                case "Product_Number_Other": //product number text 
                    break;

                case "Product_Number_Price": //Product number at price grid level, really part of pricing
                    break;

                /* product numbers */ 

                /* options */
                    /* TODO: collect info into "option" object then process at row change */
                case "Option_Type"://PROP, SHOP, IMOP
                    break;
                case "Option_Name":
                    break;
                case "Req_for_order": //flag
                    break;
                case "Can_order_only_one": //flag
                    break;
                case "Option_Additional_Info"://comments/additional info
                    break;
                case "Option_Values"://CSV list of values
                    break;
                /* end options */
                
                /* variation level SKU */
                    /* TODO: for SKU, "collect" info into SKU object, then process at row change into product model */
                case "SKU":
                    break;
                case "SKU_Based_On"://not used any longer
                    break;
                case "SKU_Criteria_1": //format: "criteriaCode: value"
                    break;
                case "SKU_Criteria_2"://format: "criteriaCode: value"
                    break;
                case "SKU_Criteria_3"://NOT supported yet
                    break;
                case "SKU_Criteria_4"://NOT supported yet
                    break;
                case "Inventory_Quantity":
                    break;
                case "Inventory_Status":
                    break;
                case "Inventory_Link":
                    break;
                /* end SKU */

                /* pricing fields */
                    /* TODO: for pricing, "collect" info into pricing object, then process at row change into product model */
                case "Price_Type": //new field to distinguish List/Net pricing
                    break;
                case "Currency": //product level and required
                    break;
                case "Base_Price_Name"://required for each grid
                    break;
                case "Base_Price_Criteria_1":
                    break;
                case "Base_Price_Criteria_2":
                    break;
                case "Price_Includes":
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
                    handleBaseDiscountCodes(text, colName);
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
                    handleBasePrices(text, colName);
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
                case "Upcharge_Name": //required
                    break;
                case "Upcharge_Type": //run charge, color charge, etc. validate against lookup, required
                    break;
                case "Upcharge_Level": //order or qty level, default "other"
                    break;
                case "Upcharge_Criteria_1":
                    break;
                case "Upcharge_Criteria_2":
                    break;
                case "Upcharge_Details": //aka price includes
                    break;
                case "U_QUR_Flag":
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
                    handleUpchargeDiscounts(text, colName);
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
                    handleUpchargePrices(text, colName);
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
            }
        }

        private static void processShapes(string text)
        {            
            //comma delimited list of shapes
            if (_firstRowForProduct)
            {
                string criteriaCode = Constants.CriteriaCodes.Shape;
                var shapes = extractCsvList(text);
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                ICollection<CriteriaSetValue> existingCsvalues = new List<CriteriaSetValue>();

                if (criteriaSet == null)
                {
                    criteriaSet = AddCriteriaSet(criteriaCode);
                }           
                else
                {
                    existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                }

                shapes.ForEach(s =>
                {
                    var shapeFound = shapesLookup.FirstOrDefault(l => l.Value == s);
                    if (shapeFound != null)
                    {
                        var exists = existingCsvalues.Any(v => v.BaseLookupValue.ToLower() == s.ToLower());
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            createNewValue(criteriaCode, s, shapeFound.Key);
                        }
                    }
                    else
                    {
                        //log batch error
                        AddValidationError(criteriaCode, s);
                        _hasErrors = true;
                    }
                });

                //delete values that are missing from the list in the file
                var valuesToDelete = existingCsvalues.Select(e => e.BaseLookupValue).Except(shapes).Select(s => s).ToList();                

                valuesToDelete.ForEach(e =>
                {
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.BaseLookupValue == e);
                    criteriaSet.CriteriaSetValues.Remove(toDelete);
                });
            }
        }

        private static void processThemes(string text)
        {
            //comma delimited list of themes
            if (_firstRowForProduct)
            {
                string criteriaCode = Constants.CriteriaCodes.Theme;
                var themes = extractCsvList(text);
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                ICollection<CriteriaSetValue> existingCsvalues = new List<CriteriaSetValue>();                             

                if (criteriaSet == null)
                {
                    criteriaSet = AddCriteriaSet(criteriaCode);
                }
                else
                {
                    existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                }

                themes.ForEach(theme =>
                {                    
                    var themeFound = themesLookup.FirstOrDefault(t => t.CodeValue.ToLower() == theme.ToLower());

                    if (themeFound != null)
                    {
                        var exists = existingCsvalues.Any(v => v.BaseLookupValue.ToLower() == theme.ToLower());
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            createNewValue(criteriaCode, theme, themeFound.ID);
                        }
                    }
                    else
                    {
                        //log batch error
                        AddValidationError(criteriaCode, theme);
                        _hasErrors = true;
                    }
                });

                //delete values that are missing from the list in the file
                var valuesToDelete = existingCsvalues.Select(e => e.BaseLookupValue).Except(themes).Select(t => t).ToList();

                valuesToDelete.ForEach(e =>
                {
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.BaseLookupValue == e);
                    criteriaSet.CriteriaSetValues.Remove(toDelete);
                });
            }
        }

        private static void processTradenames(string text)
        {            
            //comma delimited list of trade names
            if (_firstRowForProduct)
            {
                string criteriaCode = Constants.CriteriaCodes.TradeName;
                var tradenames = extractCsvList(text);
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                ICollection<CriteriaSetValue> existingCsvalues = new List<CriteriaSetValue>();

                if (criteriaSet == null)
                {
                    criteriaSet = AddCriteriaSet(criteriaCode);
                }
                else
                {
                    existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                }

                tradenames.ForEach(tradename =>
                {
                    var results = DataFetchers.Lookup.GetMatchingTradenames(tradename);                    
                    var  tradenameFound = results.FirstOrDefault();

                    if (tradenameFound != null)
                    {                                       
                        var exists = existingCsvalues.Any(v => v.BaseLookupValue.ToLower() == tradename.ToLower());
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            createNewValue(criteriaCode, tradename, tradenameFound.Key);
                        }
                    }
                    else
                    {
                        //log batch error
                        AddValidationError(criteriaCode, tradename);
                        _hasErrors = true;
                    }
                });

                //delete values that are missing from the list in the file
                var valuesToDelete = existingCsvalues.Select(e => e.BaseLookupValue).Except(tradenames).Select(t => t).ToList();

                valuesToDelete.ForEach(e =>
                {
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.BaseLookupValue == e);
                    criteriaSet.CriteriaSetValues.Remove(toDelete);
                });
            }
        }

        private static void processOrigins(string text)
        {
            //comma delimited list of origins
            if (_firstRowForProduct)
            {
                string criteriaCode = "ORGN";
                var origins = extractCsvList(text);
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                ICollection<CriteriaSetValue> existingCsvalues = new List<CriteriaSetValue>();

                if (criteriaSet == null)
                {
                    criteriaSet = AddCriteriaSet(criteriaCode);
                }
                else
                {
                    existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                }

                origins.ForEach(org =>
                {
                    var originFound = originsLookup.FirstOrDefault(l => l.CodeValue.ToLower() == org.ToLower());
                    if (originFound != null)
                    {
                        var exists = existingCsvalues.Any(v => v.BaseLookupValue.ToLower() == org.ToLower());
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            createNewValue(criteriaCode, org, originFound.ID.Value);
                        }
                    }
                    else
                    {
                        //log batch error
                        AddValidationError(criteriaCode, org);
                        _hasErrors = true;
                    }
                });

                //delete values that are missing from the list in the file
                var valuesToDelete = existingCsvalues.Select(e => e.BaseLookupValue).Except(origins).Select(s => s).ToList();

                valuesToDelete.ForEach(e =>
                {
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.BaseLookupValue == e);
                    criteriaSet.CriteriaSetValues.Remove(toDelete);
                });
            }
        }

        private static void processPackagingOptions(string text)
        {
            //comma delimited list of packaging options
            if (_firstRowForProduct)
            {
                string criteriaCode = Constants.CriteriaCodes.Packaging;
                var packagingOptions = extractCsvList(text);
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                ICollection<CriteriaSetValue> existingCsvalues = new List<CriteriaSetValue>();
                long customPackagingScvId = 0;

                //get id for custom packaging
                var customPackaging = packagingLookup.FirstOrDefault(p => p.Value == "Custom");
                if (customPackaging != null)
                {
                    customPackagingScvId = customPackaging.Key;
                }

                if (criteriaSet == null)
                {
                    criteriaSet = AddCriteriaSet(criteriaCode);
                }
                else
                {
                    existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                }

                packagingOptions.ForEach(pkg =>
                {
                    var packagingOptionFound = packagingLookup.FirstOrDefault(l => l.Value.ToLower() == pkg.ToLower());
                    if (packagingOptionFound != null)
                    {
                        var exists = existingCsvalues.Any(v => v.Value.ToLower() == pkg.ToLower());
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            createNewValue(criteriaCode, pkg, packagingOptionFound.Key);
                        }
                    }
                    else
                    {
                        //this will be a custom packaging option                        
                        createNewValue(criteriaCode, pkg, customPackagingScvId, "CUST");
                    }
                });

                //delete values that are missing from the list in the file
                var valuesToDelete = existingCsvalues.Select(e => e.Value).Except(packagingOptions).Select(s => s).ToList();

                valuesToDelete.ForEach(e =>
                {
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.Value == e);
                    criteriaSet.CriteriaSetValues.Remove(toDelete);
                });
            }
        }

        private static void processLessThanMinimum(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processImprintSizes(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processImprintMethods(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processImprintLocations(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processImprintColors(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processImprintArtwork(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processAdditionalLocations(string text)
        {            
            //comma delimited list of additional locations
            if (_firstRowForProduct)
            {
                string criteriaCode = Constants.CriteriaCodes.AdditionaLocation;
                var additionalLocations = extractCsvList(text);
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                ICollection<CriteriaSetValue> existingCsvalues = new List<CriteriaSetValue>();
                long customAddlLocScvId = 0;

                //get id for custom additional location
                var addlLoc = imprintCriteriaLookup.FirstOrDefault(i => i.Code == Constants.CriteriaCodes.AdditionaLocation);

                if (addlLoc != null)
                {
                    var group = addlLoc.CodeValueGroups.FirstOrDefault(cvg => cvg.Description == "Other");

                    if (group != null)
                    {
                        customAddlLocScvId = group.SetCodeValues.FirstOrDefault().ID;
                    }
                }                

                if (criteriaSet == null)
                {
                    criteriaSet = AddCriteriaSet(criteriaCode);
                }
                else
                {
                    existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                }

                additionalLocations.ForEach(l =>
                {
                    //check if the value already exists
                    var exists = existingCsvalues.Any(v => v.Value.ToLower() == l.ToLower());
                    if (!exists)
                    {                        
                        //add new value if it doesn't exists                        
                        createNewValue(criteriaCode, l, customAddlLocScvId, "CUST");                        
                    }                    
                });

                //delete values that are missing from the list in the file
                var valuesToDelete = existingCsvalues.Select(e => e.Value).Except(additionalLocations).Select(s => s).ToList();

                valuesToDelete.ForEach(e =>
                {
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.Value == e);
                    criteriaSet.CriteriaSetValues.Remove(toDelete);
                });
            }
        }

        private static void processAdditionalColors(string text) 
        {
            //comma delimited list of additional colors
            if (_firstRowForProduct)
            {
                string criteriaCode = Constants.CriteriaCodes.AdditionalColor;
                var additionalColors = extractCsvList(text);
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                ICollection<CriteriaSetValue> existingCsvalues = new List<CriteriaSetValue>();
                long customAddColScvId = 0;

                //get id for custom additional color
                var addlCol = imprintCriteriaLookup.FirstOrDefault(i => i.Code == Constants.CriteriaCodes.AdditionalColor);

                if (addlCol != null)
                {
                    var group = addlCol.CodeValueGroups.FirstOrDefault(cvg => cvg.Description == "Other");

                    if (group != null)
                    {
                        customAddColScvId = group.SetCodeValues.FirstOrDefault().ID;
                    }
                }

                if (criteriaSet == null)
                {
                    criteriaSet = AddCriteriaSet(criteriaCode);
                }
                else
                {
                    existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                }

                additionalColors.ForEach(l =>
                {
                    //check if the value already exists
                    var exists = existingCsvalues.Any(v => v.Value.ToLower() == l.ToLower());
                    if (!exists)
                    {
                        //add new value if it doesn't exists                        
                        createNewValue(criteriaCode, l, customAddColScvId, "CUST");
                    }
                });

                //delete values that are missing from the list in the file
                var valuesToDelete = existingCsvalues.Select(e => e.Value).Except(additionalColors).Select(s => s).ToList();

                valuesToDelete.ForEach(e =>
                {
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.Value == e);
                    criteriaSet.CriteriaSetValues.Remove(toDelete);
                });
            }
        }

        private static void processSafetyWarnings(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processComplianceCertifications(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processLineNames(string text)
        {            
            //comma delimited list of line names
            if (_firstRowForProduct)
            {                
                var linenames = extractCsvList(text);
                var existingLinenames = _currentProduct.SelectedLineNames;

                linenames.ForEach(linename =>
                {
                    var linenameFound = linenamesLookup.FirstOrDefault(l => l.Name == linename);
                    if (linenameFound != null)
                    {
                        //check if the line name is already associated with the product
                        var exists = existingLinenames.Any(v => v.ID == linenameFound.ID);
                        if (!exists)
                        {
                            //associate line name with the product
                            _currentProduct.SelectedLineNames.Add(linenameFound);
                        }
                    }
                    else
                    {
                        //log batch error
                        AddValidationError("LNNM", linename);
                        _hasErrors = true;
                    }
                });

                //delete line names that are missing from the list in the file
                var lineneamesToDelete = existingLinenames.Select(e => e.Name).Except(linenames).Select(s => s).ToList();

                lineneamesToDelete.ForEach(l =>
                {
                    var toDelete = _currentProduct.SelectedLineNames.FirstOrDefault(v => v.Name == l);
                    _currentProduct.SelectedLineNames.Remove(toDelete);                    
                });
            }
        }

        private static void processInventoryStatus(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processProductSKU(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processInventoryQty(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processCatalogInfo(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processProductDataSheet(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processDistributorOnlyViewFlag(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processDistributorOnlyComment(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processProductDisclaimer(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processConfirmationDate(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processAdditionalProductInfo(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processAdditionalShippingInfo(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processMaterials(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processSizeGroups(string text)
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

        private static void handleUpchargePrices(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleUpchargeDiscounts(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleBaseQty(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleBasePrices(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void handleBaseDiscountCodes(string text, string colName)
        {
            //throw new NotImplementedException();
        }

        private static void processProductColors(string text)
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

        private static void createNewValue(string criteriaCode, string value, long setCodeValueId, string valueTypeCode = "LOOK", string valueDetail = "", string optionName = "")
        {
            var cSet = getCriteriaSetByCode(criteriaCode, optionName);

            if (cSet == null)
            {
                cSet = AddCriteriaSet(criteriaCode, optionName);  
            }

            //create new criteria set value
            var newCsv = new CriteriaSetValue
            {
                CriteriaCode = criteriaCode,
                CriteriaSetId = cSet.CriteriaSetId,
                Value = value,
                ID = --globalUniqueId,
                ValueTypeCode = valueTypeCode,
                CriteriaValueDetail = valueDetail
            };

            //create new criteria set code value
            var newCscv = new CriteriaSetCodeValue
            {
                CriteriaSetValueId = newCsv.ID,
                SetCodeValueId = setCodeValueId,
                ID = --globalUniqueId
            };

            newCsv.CriteriaSetCodeValues.Add(newCscv);
            cSet.CriteriaSetValues.Add(newCsv);
        }

        private static IEnumerable<CriteriaSetValue> getCriteriaSetValuesByCode(string criteriaCode, string optionName = "")
        {
            var result = new List<CriteriaSetValue>();

            var cSet = getCriteriaSetByCode(criteriaCode, optionName);
            result = cSet.CriteriaSetValues.ToList();

            return result;
        }

        private static ProductCriteriaSet getCriteriaSetByCode(string criteriaCode, string optionName = "")
        {
            var cSets = new List<ProductCriteriaSet>();
            ProductCriteriaSet retVal = new ProductCriteriaSet();
            var prodConfig = _currentProduct.ProductConfigurations.FirstOrDefault(c => c.IsDefault);

            if (prodConfig != null)
            {
                cSets = prodConfig.ProductCriteriaSets.Where(c => c.CriteriaCode == criteriaCode).ToList();
                if (!string.IsNullOrWhiteSpace(optionName))
                {
                    retVal = cSets.Where(c => c.CriteriaDetail == optionName).FirstOrDefault();
                }
                else
                {
                    retVal = cSets.FirstOrDefault();
                }
            }

            return retVal;
        }

        private static ProductCriteriaSet AddCriteriaSet(string criteriaCode, string optionName = "")
        {
            var newCs = new ProductCriteriaSet
            {
                CriteriaCode = criteriaCode,
                CriteriaSetId = 0,
                ProductId = _currentProduct.ID
            };

            if (!string.IsNullOrWhiteSpace(optionName))
            {
                newCs.CriteriaDetail = optionName;
            }

            _currentProduct.ProductConfigurations.FirstOrDefault(cfg => cfg.IsDefault).ProductCriteriaSets.Add(newCs);

            return newCs;
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
                //var categories = text.ConvertToList();//TODO: do not split on spaces.
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
                //var urls = text.ConvertToList();//TODO: do not split on spaces
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
                _currentProduct.ProductLevelInventoryLink = BasicStringFieldProcessor.UpdateField(text, _currentProduct.ProductLevelInventoryLink);
        }

        private static void processSummary(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Summary = BasicStringFieldProcessor.UpdateField(text, _currentProduct.Summary);
        }

        private static void processDescription(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Description = BasicStringFieldProcessor.UpdateField(text, _currentProduct.Description);
        }

        private static void processProductNumber(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AsiProdNo = BasicStringFieldProcessor.UpdateField(text, _currentProduct.AsiProdNo);
        }

        private static void processProductName(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Name = BasicStringFieldProcessor.UpdateField(text, _currentProduct.Name);
        }        

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
            //TODO: for now we just figure out which sheet we're using, 
            //but in reality the user selects it and tells us what they're using
            // which way do we go with this rewrite?

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

        /// <summary>
        /// using input string, compare against provided lookup list and return list of values that match with appropriate code, otherwise null if no match. 
        /// </summary>
        /// <param name="lstInputLookups">list of strings to lookup and validate against LookupList</param>
        /// <param name="lookupList">List of known values to match against</param>
        /// <returns>list of lookup values with matching code or NULL</returns>
        public static List<GenericLookUp> ValidateLookupValues(List<string> lstInputLookups, List<GenericLookUp> lookupList)
        {
            var lstUserLookupList = new List<GenericLookUp>();
            //var lstInputLookups = strInputLookups.ConvertToList();
            foreach (var lookupValue in lstInputLookups)
            {
                var existingLookup = lookupList.Find(l => l.CodeValue == lookupValue);
                lstUserLookupList.Add(existingLookup ?? new GenericLookUp { CodeValue = lookupValue, ID = null });
            }
            return lstUserLookupList;
        }

        private static bool compareColumns(IEnumerable<string> columnMetaData)
        {
            //test by combining lists where the names match, removing quotes (for csv we needed quote removal, openxml should remove them)
            //stole this from current barista logic, but note that ZIP uses "first" list length to know when to stop comparing. 
            // this means additional columns in second list (_sheetcolumnslist here) are ignored.

            var test = columnMetaData
               .Zip(_sheetColumnsList, (a, b) => a.Replace("\"\"", "\"").Trim('\"') == b ? 1 : 0)
               .Select((a, i) => new { Index = i, Value = a }).ToArray();

            //TODO: log which value(s) didn't match to error log, see barista code for this
            
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

        private static void AddValidationError(string criteriaCode, string info)
        {
            _curBatch.BatchErrorLogs.Add(new BatchErrorLog
            {
                FieldCode = criteriaCode,
                ErrorMessageCode = "ILUV",
                AdditionalInfo = info,
                ProductId = _currentProduct.ID,
                ExternalProductId = _curXid
            });
        }

        private static void logit(string message)
        {
            _log.Debug(message);
        }
    }
}
