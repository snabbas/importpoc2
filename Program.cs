using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportPOC2.Models;
using ImportPOC2.Processors;
using Newtonsoft.Json;
using Radar.Core.Models.Batch;
using Radar.Data;
using Radar.Models;
using Radar.Models.Criteria;
using Radar.Models.Product;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using Constants = Radar.Core.Common.Constants;
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
        private static HttpClient _radarHttpClient;
        private static Product _currentProduct;
        private static bool _firstRowForProduct = true;
        private static int _globalUniqueId = 0;
        private static bool _publishCurrentProduct = true;

        private static log4net.ILog _log;
        private static bool _hasErrors = false;
        private static ProductRow _curProdRow;

        static void Main(string[] args)
        {
            _log = log4net.LogManager.GetLogger(typeof(Program));

            //get service location from config 
            var baseUri = ConfigurationManager.AppSettings["radarApiLocation"] ?? string.Empty;
            _radarHttpClient = new HttpClient { BaseAddress = new Uri(baseUri) };

            //onetime stuff
            _radarHttpClient.DefaultRequestHeaders.Accept.Clear();
            _radarHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            Lookups.RadarHttpClient = _radarHttpClient;

            //NOTES: 
            //file name is batch ID - use to retreive batch to assign details. 
            // also use to determine company ID of sheet. 

            //get directory from config
            //TODO: change this to read from config
            //var curDir = Directory.GetCurrentDirectory();
            var curDir = ConfigurationManager.AppSettings["excelFileLocation"] ?? Directory.GetCurrentDirectory();

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
            //as batch has changed, so has company ID, so ensure company-specific lookups are appropriately updated
            Lookups.CurrentCompanyId = _companyId;

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
                            _publishCurrentProduct = true;
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
            _curProdRow = new ProductRow();

            foreach (var column in row.Elements<Cell>())
            {
                var text = getCellText(column);
                var colIndex = getColIndex(column);

                if (colIndex == 0 && string.IsNullOrWhiteSpace(text))
                {
                    //row is invalid without XID, bail out
                    //logit("empty row encountered, skipping");
                    break;
                }

                populateProdObject(colIndex, text);

                if (colIndex == 0 && _curXid != text)
                {
                    //XID has changed, it's a new product
                    finishProduct();
                    _curXid = text;
                    startProduct();
                }

            }
            try
            {
                processCurrentProdRow();
            }
            catch (Exception exc)
            {
                _hasErrors = true;
                _log.Error("Unhandled exception occurred:", exc);
            }
        }

        private static void processCurrentProdRow()
        {
            processProductLevelFields();
            processSimpleLookups();
            processColorsMaterials();
            processSizes();
            processOptions();
            processPricing();
            processProductNumbers();
            processSkuInventory();
        }

        private static void processSkuInventory()
        {
            //TODO: VNI-10
        }

        private static void processProductNumbers()
        {
            //TODO: VNI-9
        }

        private static void processPricing()
        {
            //TODO: VNI-8
        }

        private static void processOptions()
        {
            //TODO: VNI-7
        }

        private static void processSizes()
        {
            //TODO: VNI-6
        }

        private static void processColorsMaterials()
        {
            //TODO: VNI-5
            processProductColors(_curProdRow.Product_Color);
            processMaterials(_curProdRow.Material);
        }

        private static void processSimpleLookups()
        {
            processCatalogInfo(_curProdRow.Catalog_Information);
            processShapes(_curProdRow.Shape);
            processThemes(_curProdRow.Theme);
            processTradenames(_curProdRow.Tradename);
            processOrigins(_curProdRow.Origin);
            processImprintMethods(_curProdRow.Imprint_Method);
            processLineNames(_curProdRow.Linename);
            processImprintArtwork(_curProdRow.Artwork);
            processImprintColors(_curProdRow.Imprint_Color);
            processSoldUnimprinted(_curProdRow.Sold_Unimprinted);           
            //personalization
            processImprintSizes(_curProdRow.Imprint_Size);
            processImprintLocations(_curProdRow.Imprint_Location);
            processAdditionalColors(_curProdRow.Additional_Color);
            processAdditionalLocations(_curProdRow.Additional_Location);
            //product sample
            //spec sample
            //production time
            //rush service
            //rush time
            //same day
            processPackagingOptions(_curProdRow.Packaging);
            processShippingItems(_curProdRow.Shipping_Items);           
            //shipping dimensions
            //shipping weight
            //shipping bills by
            //ship plain box
            processComplianceCertifications(_curProdRow.Comp_Cert);
            processProductDataSheet(_curProdRow.Product_Data_Sheet);
            processSafetyWarnings(_curProdRow.Safety_Warnings);
        }

        private static void processProductLevelFields()
        {
            processProductName(_curProdRow.Product_Name);
            processProductNumber(_curProdRow.Product_Number);
            processProductSku(_curProdRow.Product_SKU);
            processInventoryLink(_curProdRow.Product_Inventory_Link);
            processInventoryStatus(_curProdRow.Product_Inventory_Status);
            processInventoryQty(_curProdRow.Product_Inventory_Quantity);
            processDescription(_curProdRow.Description);
            processSummary(_curProdRow.Summary);
            processImage(_curProdRow.Prod_Image);
            processCategory(_curProdRow.Category);
            processKeywords(_curProdRow.Keywords);
            processAdditionalShippingInfo(_curProdRow.Shipping_Info);
            processAdditionalProductInfo(_curProdRow.Additional_Info);
            processDistributorOnlyViewFlag(_curProdRow.Distributor_View_Only);
            processDistributorOnlyComment(_curProdRow.Distibutor_Only);
            processProductDisclaimer(_curProdRow.Disclaimer);
            processCurrency(_curProdRow.Currency);
            processLessThanMinimum(_curProdRow.Less_Than_Min);
            processPriceType(_curProdRow.Price_Type);
            //breakout price - not processed
            processConfirmationDate(_curProdRow.Confirmed_Thru_Date);
            processDontMakeActive(_curProdRow.Dont_Make_Active);
            //breakout by attribute -- not processed
            //seo flag -- not processed
        }

        //TODO: gotta be a better way to do this. 
        private static void populateProdObject(int colIndex, string text)
        {
            //map the current column 
            var colName = _sheetColumnsList.ElementAt(colIndex);

            switch (colName)
            {
                case "Additional_Color":
                    _curProdRow.Additional_Color = text;
                    break;
                case "Additional_Info":
                    _curProdRow.Additional_Info = text;
                    break;
                case "Additional_Location":
                    _curProdRow.Additional_Location = text;
                    break;
                case "Artwork":
                    _curProdRow.Artwork = text;
                    break;
                case "Base_Price_Criteria_1":
                    _curProdRow.Base_Price_Criteria_1 = text;
                    break;
                case "Base_Price_Criteria_2":
                    _curProdRow.Base_Price_Criteria_2 = text;
                    break;
                case "Base_Price_Name":
                    _curProdRow.Base_Price_Name = text;
                    break;
                case "Breakout_by_other_attribute":
                    _curProdRow.Breakout_by_other_attribute = text;
                    break;
                case "Breakout_by_price":
                    _curProdRow.Breakout_by_price = text;
                    break;
                case "Can_order_only_one":
                    _curProdRow.Can_order_only_one = text;
                    break;
                case "Catalog_Information":
                    _curProdRow.Catalog_Information = text;
                    break;
                case "Category":
                    _curProdRow.Category = text;
                    break;
                case "Comp_Cert":
                    _curProdRow.Comp_Cert = text;
                    break;
                case "Confirmed_Thru_Date":
                    _curProdRow.Confirmed_Thru_Date = text;
                    break;
                case "Currency":
                    _curProdRow.Currency = text;
                    break;
                case "D1":
                    _curProdRow.D1 = text;
                    break;
                case "D10":
                    _curProdRow.D10 = text;
                    break;
                case "D2":
                    _curProdRow.D2 = text;
                    break;
                case "D3":
                    _curProdRow.D3 = text;
                    break;
                case "D4":
                    _curProdRow.D4 = text;
                    break;
                case "D5":
                    _curProdRow.D5 = text;
                    break;
                case "D6":
                    _curProdRow.D6 = text;
                    break;
                case "D7":
                    _curProdRow.D7 = text;
                    break;
                case "D8":
                    _curProdRow.D8 = text;
                    break;
                case "D9":
                    _curProdRow.D9 = text;
                    break;
                case "Description":
                    _curProdRow.Description = text;
                    break;
                case "Disclaimer":
                    _curProdRow.Disclaimer = text;
                    break;
                case "Distibutor_Only":
                    _curProdRow.Distibutor_Only = text;
                    break;
                case "Distributor_View_Only":
                    _curProdRow.Distributor_View_Only = text;
                    break;
                case "Dont_Make_Active":
                    _curProdRow.Dont_Make_Active = text;
                    break;
                case "Imprint_Color":
                    _curProdRow.Imprint_Color = text;
                    break;
                case "Imprint_Location":
                    _curProdRow.Imprint_Location = text;
                    break;
                case "Imprint_Method":
                    _curProdRow.Imprint_Method = text;
                    break;
                case "Imprint_Size":
                    _curProdRow.Imprint_Size = text;
                    break;
                case "Inventory_Link":
                    _curProdRow.Inventory_Link = text;
                    break;
                case "Inventory_Quantity":
                    _curProdRow.Inventory_Quantity = text;
                    break;
                case "Inventory_Status":
                    _curProdRow.Inventory_Status = text;
                    break;
                case "Keywords":
                    _curProdRow.Keywords = text;
                    break;
                case "Less_Than_Min":
                    _curProdRow.Less_Than_Min = text;
                    break;
                case "Linename":
                    _curProdRow.Linename = text;
                    break;
                case "Material":
                    _curProdRow.Material = text;
                    break;
                case "Option_Additional_Info":
                    _curProdRow.Option_Additional_Info = text;
                    break;
                case "Option_Name":
                    _curProdRow.Option_Name = text;
                    break;
                case "Option_Type":
                    _curProdRow.Option_Type = text;
                    break;
                case "Option_Values":
                    _curProdRow.Option_Values = text;
                    break;
                case "Origin":
                    _curProdRow.Origin = text;
                    break;
                case "P1":
                    _curProdRow.P1 = text;
                    break;
                case "P10":
                    _curProdRow.P10 = text;
                    break;
                case "P2":
                    _curProdRow.P2 = text;
                    break;
                case "P3":
                    _curProdRow.P3 = text;
                    break;
                case "P4":
                    _curProdRow.P4 = text;
                    break;
                case "P5":
                    _curProdRow.P5 = text;
                    break;
                case "P6":
                    _curProdRow.P6 = text;
                    break;
                case "P7":
                    _curProdRow.P7 = text;
                    break;
                case "P8":
                    _curProdRow.P8 = text;
                    break;
                case "P9":
                    _curProdRow.P9 = text;
                    break;
                case "Packaging":
                    _curProdRow.Packaging = text;
                    break;
                case "Personalization":
                    _curProdRow.Personalization = text;
                    break;
                case "Price_Includes":
                    _curProdRow.Price_Includes = text;
                    break;
                case "Price_Type":
                    _curProdRow.Price_Type = text;
                    break;
                case "Prod_Image":
                    _curProdRow.Prod_Image = text;
                    break;
                case "Product_Color":
                    _curProdRow.Product_Color = text;
                    break;
                case "Product_Data_Sheet":
                    _curProdRow.Product_Data_Sheet = text;
                    break;
                case "Product_Inventory_Link":
                    _curProdRow.Product_Inventory_Link = text;
                    break;
                case "Product_Inventory_Quantity":
                    _curProdRow.Product_Inventory_Quantity = text;
                    break;
                case "Product_Inventory_Status":
                    _curProdRow.Product_Inventory_Status = text;
                    break;
                case "Product_Name":
                    _curProdRow.Product_Name = text;
                    break;
                case "Product_Number":
                    _curProdRow.Product_Number = text;
                    break;
                case "Product_Number_Criteria_1":
                    _curProdRow.Product_Number_Criteria_1 = text;
                    break;
                case "Product_Number_Criteria_2":
                    _curProdRow.Product_Number_Criteria_2 = text;
                    break;
                case "Product_Number_Other":
                    _curProdRow.Product_Number_Other = text;
                    break;
                case "Product_Number_Price":
                    _curProdRow.Product_Number_Price = text;
                    break;
                case "Product_Sample":
                    _curProdRow.Product_Sample = text;
                    break;
                case "Product_SKU":
                    _curProdRow.Product_SKU = text;
                    break;
                case "Production_Time":
                    _curProdRow.Production_Time = text;
                    break;
                case "Q1":
                    _curProdRow.Q1 = text;
                    break;
                case "Q10":
                    _curProdRow.Q10 = text;
                    break;
                case "Q2":
                    _curProdRow.Q2 = text;
                    break;
                case "Q3":
                    _curProdRow.Q3 = text;
                    break;
                case "Q4":
                    _curProdRow.Q4 = text;
                    break;
                case "Q5":
                    _curProdRow.Q5 = text;
                    break;
                case "Q6":
                    _curProdRow.Q6 = text;
                    break;
                case "Q7":
                    _curProdRow.Q7 = text;
                    break;
                case "Q8":
                    _curProdRow.Q8 = text;
                    break;
                case "Q9":
                    _curProdRow.Q9 = text;
                    break;
                case "QUR_Flag":
                    _curProdRow.QUR_Flag = text;
                    break;
                case "Req_for_order":
                    _curProdRow.Req_for_order = text;
                    break;
                case "Rush_Service":
                    _curProdRow.Rush_Service = text;
                    break;
                case "Rush_Time":
                    _curProdRow.Rush_Time = text;
                    break;
                case "Safety_Warnings":
                    _curProdRow.Safety_Warnings = text;
                    break;
                case "Same_Day_Service":
                    _curProdRow.Same_Day_Service = text;
                    break;
                case "SEO_FLG":
                    _curProdRow.SEO_FLG = text;
                    break;
                case "Shape":
                    _curProdRow.Shape = text;
                    break;
                case "Ship_Plain_Box":
                    _curProdRow.Ship_Plain_Box = text;
                    break;
                case "Shipper_Bills_By":
                    _curProdRow.Shipper_Bills_By = text;
                    break;
                case "Shipping_Dimensions":
                    _curProdRow.Shipping_Dimensions = text;
                    break;
                case "Shipping_Info":
                    _curProdRow.Shipping_Info = text;
                    break;
                case "Shipping_Items":
                    _curProdRow.Shipping_Items = text;
                    break;
                case "Shipping_Weight":
                    _curProdRow.Shipping_Weight = text;
                    break;
                case "Size_Group":
                    _curProdRow.Size_Group = text;
                    break;
                case "Size_Values":
                    _curProdRow.Size_Values = text;
                    break;
                case "SKU":
                    _curProdRow.SKU = text;
                    break;
                case "SKU_Based_On":
                    _curProdRow.SKU_Based_On = text;
                    break;
                case "SKU_Criteria_1":
                    _curProdRow.SKU_Criteria_1 = text;
                    break;
                case "SKU_Criteria_2":
                    _curProdRow.SKU_Criteria_2 = text;
                    break;
                case "SKU_Criteria_3":
                    _curProdRow.SKU_Criteria_3 = text;
                    break;
                case "SKU_Criteria_4":
                    _curProdRow.SKU_Criteria_4 = text;
                    break;
                case "Sold_Unimprinted":
                    _curProdRow.Sold_Unimprinted = text;
                    break;
                case "Spec_Sample":
                    _curProdRow.Spec_Sample = text;
                    break;
                case "Summary":
                    _curProdRow.Summary = text;
                    break;
                case "Theme":
                    _curProdRow.Theme = text;
                    break;
                case "Tradename":
                    _curProdRow.Tradename = text;
                    break;
                case "U_QUR_Flag":
                    _curProdRow.U_QUR_Flag = text;
                    break;
                case "UD1":
                    _curProdRow.UD1 = text;
                    break;
                case "UD10":
                    _curProdRow.UD10 = text;
                    break;
                case "UD2":
                    _curProdRow.UD2 = text;
                    break;
                case "UD3":
                    _curProdRow.UD3 = text;
                    break;
                case "UD4":
                    _curProdRow.UD4 = text;
                    break;
                case "UD5":
                    _curProdRow.UD5 = text;
                    break;
                case "UD6":
                    _curProdRow.UD6 = text;
                    break;
                case "UD7":
                    _curProdRow.UD7 = text;
                    break;
                case "UD8":
                    _curProdRow.UD8 = text;
                    break;
                case "UD9":
                    _curProdRow.UD9 = text;
                    break;
                case "UP1":
                    _curProdRow.UP1 = text;
                    break;
                case "UP10":
                    _curProdRow.UP10 = text;
                    break;
                case "UP2":
                    _curProdRow.UP2 = text;
                    break;
                case "UP3":
                    _curProdRow.UP3 = text;
                    break;
                case "UP4":
                    _curProdRow.UP4 = text;
                    break;
                case "UP5":
                    _curProdRow.UP5 = text;
                    break;
                case "UP6":
                    _curProdRow.UP6 = text;
                    break;
                case "UP7":
                    _curProdRow.UP7 = text;
                    break;
                case "UP8":
                    _curProdRow.UP8 = text;
                    break;
                case "UP9":
                    _curProdRow.UP9 = text;
                    break;
                case "Upcharge_Criteria_1":
                    _curProdRow.Upcharge_Criteria_1 = text;
                    break;
                case "Upcharge_Criteria_2":
                    _curProdRow.Upcharge_Criteria_2 = text;
                    break;
                case "Upcharge_Details":
                    _curProdRow.Upcharge_Details = text;
                    break;
                case "Upcharge_Level":
                    _curProdRow.Upcharge_Level = text;
                    break;
                case "Upcharge_Name":
                    _curProdRow.Upcharge_Name = text;
                    break;
                case "Upcharge_Type":
                    _curProdRow.Upcharge_Type = text;
                    break;
                case "UQ1":
                    _curProdRow.UQ1 = text;
                    break;
                case "UQ10":
                    _curProdRow.UQ10 = text;
                    break;
                case "UQ2":
                    _curProdRow.UQ2 = text;
                    break;
                case "UQ3":
                    _curProdRow.UQ3 = text;
                    break;
                case "UQ4":
                    _curProdRow.UQ4 = text;
                    break;
                case "UQ5":
                    _curProdRow.UQ5 = text;
                    break;
                case "UQ6":
                    _curProdRow.UQ6 = text;
                    break;
                case "UQ7":
                    _curProdRow.UQ7 = text;
                    break;
                case "UQ8":
                    _curProdRow.UQ8 = text;
                    break;
                case "UQ9":
                    _curProdRow.UQ9 = text;
                    break;
                case "XID":
                    _curProdRow.XID = text;
                    break;
            }
        }


        /// <summary>
        /// converts an excel column reference value (e.g., "A", "AZ", "Q") into column index.
        /// </summary>
        /// <param name="column"></param>
        /// <returns>column index as integer</returns>
        private static int getColIndex(CellType column)
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
                var results = _radarHttpClient.GetAsync(endpointUrl).Result;

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

        //changed to use generic processing method
        private static void processShapes(string text)
        {            
            genericProcess(text, Constants.CriteriaCodes.Shape, Lookups.ShapesLookup);
        }

        private static void genericProcess(string text, string criteriaCode, IEnumerable<GenericLookUp> lookup)
        {
            if (_firstRowForProduct)
            {
                var valueList = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);

                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

                valueList.ForEach(value =>
                {
                    var existing = lookup.FirstOrDefault(l => l.CodeValue.ToLower() == value.ToLower());
                    if (existing != null)
                    {
                        var exists = existingCsvalues.Any(csv => csv.BaseLookupValue.ToLower() == value.ToLower());
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            if (existing.ID != null)
                                createNewValue(criteriaCode, value, existing.ID.Value);
                        }
                    }
                    else
                    {
                        //log batch error
                        addValidationError(criteriaCode, value);
                        _hasErrors = true;
                    }
                });

                deleteCsValues(existingCsvalues, valueList, criteriaSet);
            }
        }

        //TODO: I want this to be a single method that uses the same lookup type as method below
        private static void genericProcess(string text, string criteriaCode, IEnumerable<SetCodeValue> lookup)
        {
            //comma delimited list of type
            if (_firstRowForProduct)
            {
                var valueList = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);

                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                
                valueList.ForEach(item =>
                {
                    var exists = lookup.FirstOrDefault(l => String.Equals(l.CodeValue, item, StringComparison.CurrentCultureIgnoreCase));

                    if (exists != null)
                    {
                        var existing = existingCsvalues.Any(csv => csv.BaseLookupValue.ToLower() == item.ToLower());
                        //add new value if it doesn't exists
                        if (!existing)
                        {
                            createNewValue(criteriaCode, item, exists.ID);
                        }
                    }
                    else
                    {
                        //log batch error
                        addValidationError(criteriaCode, item);
                        _hasErrors = true;
                    }
                });

                deleteCsValues(existingCsvalues, valueList, criteriaSet);
            }
        }
        //TODO: see above, this should be same method as above. 
        private static void genericProcess(string text, string criteriaCode, IEnumerable<KeyValueLookUp> lookup)
        {
            //comma delimited list of type
            if (_firstRowForProduct)
            {
                var valueList = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);

                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

                valueList.ForEach(item =>
                {
                    var exists = lookup.FirstOrDefault(l => String.Equals(l.Value, item, StringComparison.CurrentCultureIgnoreCase));

                    if (exists != null)
                    {
                        var existing = existingCsvalues.Any(csv => csv.BaseLookupValue.ToLower() == item.ToLower());
                        //add new value if it doesn't exists
                        if (!existing)
                        {
                            createNewValue(criteriaCode, item, exists.Key);
                        }
                    }
                    else
                    {
                        //log batch error
                        addValidationError(criteriaCode, item);
                        _hasErrors = true;
                    }
                });

                deleteCsValues(existingCsvalues, valueList, criteriaSet);
            }
        }

        private static void genericProcessImprintCriteria(string text, string criteriaCode)
        {
            if (_firstRowForProduct)
            {
                var valueList = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                long customSetCodeValueId = 0;

                //get id for custom additional location
                var criteriaLookUp = Lookups.ImprintCriteriaLookup.FirstOrDefault(i => i.Code == criteriaCode);

                if (criteriaLookUp != null)
                {
                    var group = criteriaLookUp.CodeValueGroups.FirstOrDefault(cvg => cvg.Description == "Other");
                    if (group != null)
                    {
                        var setCodeValue = group.SetCodeValues.FirstOrDefault();
                        if (setCodeValue != null)
                            customSetCodeValueId = setCodeValue.ID;
                    }
                }

                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

                valueList.ForEach(value =>
                {
                    //check if the value already exists
                    var exists = existingCsvalues.Any(csv => csv.Value.ToLower() == value.ToLower());
                    if (!exists)
                    {
                        //add new value if it doesn't exist
                        createNewValue(criteriaCode, value, customSetCodeValueId, "CUST");
                    }
                });

                deleteCsValues(existingCsvalues, valueList, criteriaSet);
            }
        }

        private static void processThemes(string text)
        {
            genericProcess(text, Constants.CriteriaCodes.Theme, Lookups.ThemesLookup);
        }

        private static void processTradenames(string text)
        {            
            //comma delimited list of trade names
            if (_firstRowForProduct)
            {
                var criteriaCode = Constants.CriteriaCodes.TradeName;
                var tradenames = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);

                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

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
                        addValidationError(criteriaCode, tradename);
                        _hasErrors = true;
                    }
                });

                deleteCsValues(existingCsvalues, tradenames, criteriaSet);
            }
        }

        private static void processPriceType(string text)
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
                    addValidationError("priceType", text);
                    _hasErrors = true;
                }
            }
        }        

        private static void processCurrency(string text)
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
                        addValidationError("Currency", text);
                        _hasErrors = true;
                    }
                }                
            }
        }

        private static void processOrigins(string text)
        {
            genericProcess(text, "ORGN", Lookups.OriginsLookup);
        }

        private static void processShippingItems(string text)
        {
            if (_firstRowForProduct)
            {
                var criteriaCode = "SHES";
                var shippingItems = text.Split(':');
                var criteriaSet = getCriteriaSetByCode(criteriaCode);

                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

                if (shippingItems.Length == 2)
                {
                    var items = shippingItems[0];
                    var unit = shippingItems[1];
                    var criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Unit");
                    var unitFound = criteriaAttribute.UnitsOfMeasure.FirstOrDefault(u => u.DisplayName == unit);
                    if (unitFound != null)
                    {
                        var exists = existingCsvalues.Select(v => v.Value).SingleOrDefault();
                        //add new value if it doesn't exists
                        if (exists != null)
                        {
                            //if (exists.UnitValue != items)
                                exists.UnitValue = items;
                            //if (exists.UnitOfMeasureCode != unit)
                                exists.UnitOfMeasureCode = unit;
                        }
                        else
                        {
                            var value = new
                            {
                                CriteriaAttributeId = criteriaAttribute.ID,
                                UnitValue = items,
                                UnitOfMeasureCode = unit
                            };

                            var group = criteriaAttribute.CriteriaItem.CodeValueGroups.FirstOrDefault();
                            var setCodeValueId = 0L;
                            if (group != null)
                            {
                                var setCodeValue = group.SetCodeValues.FirstOrDefault();
                                if (setCodeValue != null)
                                    setCodeValueId = setCodeValue.ID;
                            }

                            createNewValue(criteriaCode, value, setCodeValueId, "CUST");
                        }
                    }
                    else
                    {
                        //log batch error
                        addValidationError(criteriaCode, unit);
                        _hasErrors = true;
                    }
                }
            }
        }

        private static void processPackagingOptions(string text)
        {
            //comma delimited list of packaging options
            if (_firstRowForProduct)
            {
                var criteriaCode = Constants.CriteriaCodes.Packaging;
                var packagingOptions = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                
                long customPackagingScvId = 0;

                //get id for custom packaging
                var customPackaging = Lookups.PackagingLookup.FirstOrDefault(p => p.Value == "Custom");
                if (customPackaging != null)
                {
                    customPackagingScvId = customPackaging.Key;
                }

                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                

                packagingOptions.ForEach(pkg =>
                {
                    var packagingOptionFound = Lookups.PackagingLookup.FirstOrDefault(l => String.Equals(l.Value, pkg, StringComparison.CurrentCultureIgnoreCase));
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

                deleteCsValues(existingCsvalues, packagingOptions, criteriaSet);
            }
        }

        private static void processLessThanMinimum(string text)
        {
            if (_firstRowForProduct)
            {
                _currentProduct.IsOrderLessThanMinimumAllowed = BasicFieldProcessor.UpdateField(text, _currentProduct.IsOrderLessThanMinimumAllowed);
            }
        }

        private static void processImprintSizes(string text)
        {
            //comma delimited list of imprint sizes
            //if (_firstRowForProduct)
            //{
            //    string criteriaCode = Constants.CriteriaCodes.ImprintSizeLocation;
            //    var imprintSizes = extractCsvList(text);
            //    var criteriaSet = getCriteriaSetByCode(criteriaCode);
            //    ICollection<CriteriaSetValue> existingCsvalues = new List<CriteriaSetValue>();
            //    long customImprintSizeLocationScvId = 0;

            //    //get id for imprint size location
            //    var imsz = imprintSizeLocation.FirstOrDefault(i => i.Code == criteriaCode);
            //    if (imsz != null)
            //    {
            //        var group = imsz.CodeValueGroups.FirstOrDefault(cvg => cvg.Description == "Other");
            //        if (group != null)
            //        {
            //            customImprintSizeLocationScvId = group.SetCodeValues.FirstOrDefault().ID;
            //        }
            //    }

            //    if (criteriaSet == null)
            //    {
            //        criteriaSet = AddCriteriaSet(criteriaCode);
            //    }
            //    else
            //    {
            //        existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
            //    }                

            //    deleteCsValues(existingCsvalues, imprintSizes, criteriaSet);
            //}
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
            //comma delimited list of imprint colors
            if (_firstRowForProduct)
            {
                var criteriaCode = Constants.CriteriaCodes.ImprintColor;
                var imprintColors = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                
                long imprintColorScvId = 0;

                //get set code value id for imprint color
                var imcl = Lookups.ImprintColorLookup.FirstOrDefault(i => i.Code == criteriaCode);

                if (imcl != null)
                {
                    var group = imcl.CodeValueGroups.FirstOrDefault(cvg => cvg.Description == "Other");

                    if (group != null)
                    {
                        var setCodeValue = group.SetCodeValues.FirstOrDefault();
                        if (setCodeValue != null)
                            imprintColorScvId = setCodeValue.ID;
                    }
                }     

                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                
                imprintColors.ForEach(color =>
                {
                    //check if the value already exists
                    var exists = existingCsvalues.Any(v => v.Value.ToLower() == color.ToLower());
                    if (!exists)
                    {
                        //add new value if it doesn't exists                        
                        createNewValue(criteriaCode, color, imprintColorScvId, "CUST");
                    }
                });

                deleteCsValues(existingCsvalues, imprintColors, criteriaSet);
            }
        }

        private static void processSoldUnimprinted(string text)
        {
            if (_firstRowForProduct)
            {
                //should be Y/N
                var soldUnimprinted = text;
                var validValues = new [] {"Y", "N"};
                if (!string.IsNullOrWhiteSpace(soldUnimprinted) && !validValues.Contains(soldUnimprinted))
                {
                    addValidationError("GNER", "invalid value for Sold Unimprinted");
                    return;
                }

                long soldUnimprintedScvId = 0;

                //get set code value id for sold unimprinted
                var unimprinted = Lookups.ImprintMethodsLookup.FirstOrDefault(i => i.CodeValue == "Unimprinted");

                if (unimprinted != null)
                {
                    if (unimprinted.ID != null)
                        soldUnimprintedScvId = unimprinted.ID.Value;
                }

                var criteriaCode = Constants.CriteriaCodes.ImprintMethod;
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                CriteriaSetValue unimprintedCsvalue = null;

                if (criteriaSet != null)
                {
                    unimprintedCsvalue = getCsValueBySetCodeValueId(soldUnimprintedScvId, criteriaSet);
                }

                if (soldUnimprinted == "Y")
                {                    
                    //create new value for unimprinted if it doesn't exists
                    if (unimprintedCsvalue == null)
                    {                                                                        
                        createNewValue(criteriaCode, "Unimprinted", soldUnimprintedScvId);                        
                    }
                    _currentProduct.IsAvailableUnimprinted = true;
                }
                else
                {                    
                    if (criteriaSet != null && unimprintedCsvalue != null)
                    {
                        criteriaSet.CriteriaSetValues.Remove(unimprintedCsvalue);
                    }
                    _currentProduct.IsAvailableUnimprinted = false;
                }
            }
        }

        private static void processImprintArtwork(string text)
        {
            //throw new NotImplementedException();
        }


        private static void processAdditionalLocations(string text)
        {            
            //comma delimited list of additional locations
            genericProcessImprintCriteria(text, Constants.CriteriaCodes.AdditionaLocation);
        }

        private static void processAdditionalColors(string text) 
        {
            //comma delimited list of additional colors
            genericProcessImprintCriteria(text, Constants.CriteriaCodes.AdditionalColor);
        }

        private static void processSafetyWarnings(string text)
        {
            if (_firstRowForProduct)
            {
                var safetyWarnings = text.ConvertToList();

                if (_currentProduct.SelectedSafetyWarnings == null)
                    _currentProduct.SelectedSafetyWarnings = new Collection<SafetyWarning>();

                safetyWarnings.ForEach(curSafetyWarning =>
                {
                    //need to lookup safetyWarnings
                    var safetyWarning = Lookups.SafetywarningsLookup.FirstOrDefault(c => c.Value == curSafetyWarning);
                    if (safetyWarning != null)
                    {
                        var existing = _currentProduct.SelectedSafetyWarnings.FirstOrDefault(c => c.Description == safetyWarning.Value);
                        if (existing == null)
                        {
                            var newSafetyWarning = new SafetyWarning { Code = safetyWarning.Key, WarningText = safetyWarning.Value };
                            _currentProduct.SelectedSafetyWarnings.Add(newSafetyWarning);
                        }
                        else
                        {
                            //s/b nothing to do? 
                        }
                    }
                });

                //remove any safety Warnings from product that aren't on the sheet; get list of codes from the sheet 
                var sheetSafetyWarningsList = safetyWarnings.Join(Lookups.SafetywarningsLookup, saf => saf, lookup => lookup.Value, (saf, lookup) => lookup.Value);
                var toRemove = _currentProduct.SelectedSafetyWarnings.Where(c => !sheetSafetyWarningsList.Contains(c.Description)).ToList();
                toRemove.ForEach(r => _currentProduct.SelectedSafetyWarnings.Remove(r));
            }
        }

        private static void processComplianceCertifications(string text)
        {
            if (_firstRowForProduct)
            {
                var complianceCertifications = text.ConvertToList();

                if (_currentProduct.SelectedComplianceCerts == null)
                    _currentProduct.SelectedComplianceCerts = new Collection<ProductComplianceCert>();

                complianceCertifications.ForEach(curCert =>
                {
                    //need to lookup complianceCertifications
                    var complianceCert = Lookups.ComplianceLookup.FirstOrDefault(c => c.Value == curCert);
                    if (complianceCert != null)
                    {
                        var existing = _currentProduct.SelectedComplianceCerts.FirstOrDefault(c => c.Description == complianceCert.Value);
                        if (existing == null)
                        {
                            var newComplianceCert = new ProductComplianceCert { ComplianceCertId = Convert.ToInt32(complianceCert.Key), Description = complianceCert.Value };
                            _currentProduct.SelectedComplianceCerts.Add(newComplianceCert);
                        }
                        else
                        {
                            //s/b nothing to do? 
                        }
                    }
                });

                //remove any ComplianceCert from product that aren't on the sheet; get list of codes from the sheet 
                var sheetComplianceCertList = complianceCertifications.Join(Lookups.ComplianceLookup, cer => cer, lookup => lookup.Value, (cer, lookup) => lookup.Value);
                var toRemove = _currentProduct.SelectedComplianceCerts.Where(c => !sheetComplianceCertList.Contains(c.Description)).ToList();
                toRemove.ForEach(r => _currentProduct.SelectedComplianceCerts.Remove(r));
            }
        }

        private static void processLineNames(string text)
        {            
            //comma delimited list of line names
            if (_firstRowForProduct)
            {
                var linenames = text.ConvertToList();
                var existingLinenames = _currentProduct.SelectedLineNames;

                linenames.ForEach(linename =>
                {
                    var linenameFound = Lookups.LinenamesLookup.FirstOrDefault(l => l.Name == linename);
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
                        addValidationError("LNNM", linename);
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
            if (_firstRowForProduct)
            {
                var inventoryStatusFound = Lookups.InventoryStatusesLookup.FirstOrDefault(t => t.Value == text);
                if (inventoryStatusFound != null)
                {
                    _currentProduct.ProductLevelInventoryStatusCode = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductLevelInventoryStatusCode);
                }
                else
                {
                    addValidationError("InventoryStatus", text);
                    _hasErrors = true;
                }
            }
        }

        private static void processProductSku(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.ProductLevelSku = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductLevelSku);
        }

        private static void processInventoryQty(string text)
        {
            if (_firstRowForProduct)
            {
                _currentProduct.ProductLevelInventoryQuantity = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductLevelInventoryQuantity);
            }
        }

        private static void processCatalogInfo(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processProductDataSheet(string text)
        {
            if (_firstRowForProduct)

                if (_currentProduct.ProductDataSheet == null)
                    _currentProduct.ProductDataSheet = new ProductDataSheet();

                _currentProduct.ProductDataSheet.Url = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductDataSheet.Url);
        }

        private static void processDistributorOnlyViewFlag(string text)
        {
            //throw new NotImplementedException();
        }

        private static void processDistributorOnlyComment(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.DistributorComments = BasicFieldProcessor.UpdateField(text, _currentProduct.DistributorComments);
        }

        private static void processProductDisclaimer(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Disclaimer = BasicFieldProcessor.UpdateField(text, _currentProduct.Disclaimer);
        }

        private static void processConfirmationDate(string text)
        {
            if (_firstRowForProduct)
            {              
                _currentProduct.PriceConfirmationDate = BasicFieldProcessor.UpdateField(text, _currentProduct.PriceConfirmationDate);
            }
        }

        private static void processAdditionalProductInfo(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AdditionalInfo = BasicFieldProcessor.UpdateField(text, _currentProduct.AdditionalInfo);
        }

        private static void processAdditionalShippingInfo(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AddtionalShippingInfo = BasicFieldProcessor.UpdateField(text, _currentProduct.AddtionalShippingInfo);
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
                var colorList = text.ConvertToList();
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
                    var colorObj = Lookups.ColorGroupList.SelectMany(g => g.CodeValueGroups).FirstOrDefault(g => String.Equals(g.Description, colorName, StringComparison.CurrentCultureIgnoreCase));
                    var existing = productColors.FirstOrDefault(p => p.Value == aliasName);

                    long setCodeId = 0;
                    if (colorObj == null)
                    {
                        //they picked a color that doesn't exist, so we choose the "other" set code to assign it on the new value
                        colorObj = Lookups.ColorGroupList.SelectMany(g => g.CodeValueGroups).FirstOrDefault(g => string.Equals(g.Description, "Unclassified/Other", StringComparison.CurrentCultureIgnoreCase));
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

        private static void createNewValue(string criteriaCode, object value, long setCodeValueId, string valueTypeCode = "LOOK", string valueDetail = "", string optionName = "")
        {
            var cSet = getCriteriaSetByCode(criteriaCode, optionName) ;

            //create new criteria set value
            var newCsv = new CriteriaSetValue
            {
                CriteriaCode = criteriaCode,
                CriteriaSetId = cSet.CriteriaSetId,
                Value = value,
                ID = --_globalUniqueId,
                ValueTypeCode = valueTypeCode,
                CriteriaValueDetail = valueDetail
            };

            //create new criteria set code value
            var newCscv = new CriteriaSetCodeValue
            {
                CriteriaSetValueId = newCsv.ID,
                SetCodeValueId = setCodeValueId,
                ID = --_globalUniqueId
            };

            newCsv.CriteriaSetCodeValues.Add(newCscv);
            cSet.CriteriaSetValues.Add(newCsv);
        }

        private static IEnumerable<CriteriaSetValue> getCriteriaSetValuesByCode(string criteriaCode, string optionName = "")
        {
            var cSet = getCriteriaSetByCode(criteriaCode, optionName);
            var result = cSet.CriteriaSetValues.ToList();

            return result;
        }

        private static ProductCriteriaSet getCriteriaSetByCode(string criteriaCode, string optionName = "")
        {
            ProductCriteriaSet retVal = null;
            var prodConfig = _currentProduct.ProductConfigurations.FirstOrDefault(c => c.IsDefault);

            if (prodConfig != null)
            {
                var cSets = prodConfig.ProductCriteriaSets.Where(c => c.CriteriaCode == criteriaCode).ToList();
                retVal = !string.IsNullOrWhiteSpace(optionName) ? cSets.FirstOrDefault(c => c.CriteriaDetail == optionName) : cSets.FirstOrDefault();
            }

            retVal = retVal ?? addCriteriaSet(criteriaCode, optionName);

            return retVal;
        }

        private static ProductCriteriaSet addCriteriaSet(string criteriaCode, string optionName = "")
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

            var productConfiguration = _currentProduct.ProductConfigurations.FirstOrDefault(cfg => cfg.IsDefault);
            if (productConfiguration != null)
                productConfiguration.ProductCriteriaSets.Add(newCs);

            return newCs;
        }

        private static CriteriaSetValue getCsValueBySetCodeValueId(long scvId, ProductCriteriaSet criteriaSet)
        {           
            CriteriaSetValue retVal = null;

            if (criteriaSet != null)
            {                               
                foreach (var v in criteriaSet.CriteriaSetValues)
                {
                    var scv = v.CriteriaSetCodeValues.FirstOrDefault(s => s.SetCodeValueId == scvId);
                    if (scv != null)
                    {
                        retVal = v;
                        break;
                    }
                }
            }

            return retVal; 
        }

        private static void deleteCsValues(IEnumerable<CriteriaSetValue> entities, IEnumerable<string> models, ProductCriteriaSet criteriaSet)
        {
            //delete values that are missing from the list in the file
            var valuesToDelete = entities.Select(e => e.Value).Except(models).Select(s => s).ToList();
            valuesToDelete.ForEach(e =>
            {
                var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.Value == e);
                criteriaSet.CriteriaSetValues.Remove(toDelete);
            });
        }

        //private static List<string> extractCsvList(string text)
        //{
        //    //returns list of strings from input string, split on commas, each value is trimmed, and only non-empty values are returned.
        //    //return text.Split(',').Select(str => str.Trim()).Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
        //    return text.ConvertToList();
        //}

        private static void processKeywords(string text)
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
                        var newKeyword = new ProductKeyword {Value = keyword, TypeCode = "HIDD", ID = _globalUniqueId--};
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
                var sheetCategoryList = categories.Join(Lookups.CategoryList, cat => cat, lookup => lookup.Name, (cat, lookup) => lookup.Code);
                var toRemove = _currentProduct.SelectedProductCategories.Where(c => !sheetCategoryList.Contains(c.Code)).ToList();
                toRemove.ForEach(r => _currentProduct.SelectedProductCategories.Remove(r));
            }
        }

        private static void processImage(string text)
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

        private static void processDontMakeActive(string text)
        {
            if (_firstRowForProduct)
            {
                if (!string.IsNullOrWhiteSpace(text) && text.ToLower() == "y")
                    _publishCurrentProduct = false;
            }
        }

        private static void processInventoryLink(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.ProductLevelInventoryLink = BasicFieldProcessor.UpdateField(text, _currentProduct.ProductLevelInventoryLink);
        }

        private static void processSummary(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Summary = BasicFieldProcessor.UpdateField(text, _currentProduct.Summary);
        }

        private static void processDescription(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Description = BasicFieldProcessor.UpdateField(text, _currentProduct.Description);
        }

        private static void processProductNumber(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AsiProdNo = BasicFieldProcessor.UpdateField(text, _currentProduct.AsiProdNo);
        }

        private static void processProductName(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Name = BasicFieldProcessor.UpdateField(text, _currentProduct.Name);
        }        

        private static void finishProduct()
        {
            //if we've started a radar model, 
            // we "send" the product to Radar for processing. 
            if (_currentProduct != null && !_hasErrors)
            {
                var x = _currentProduct;
                if (!_publishCurrentProduct)
                {
                    //add "no pub" attribute to radar POST
                }
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
        /// <param name="inputValueList">list of strings to lookup and validate against LookupList</param>
        /// <param name="lookupList">List of known values to match against</param>
        /// <returns>list of lookup values with matching code or NULL</returns>
        private static List<GenericLookUp> validateLookupValues(IEnumerable<string> inputValueList, List<GenericLookUp> lookupList)
        {
            //var inputValueList = strInputLookups.ConvertToList();
            return (from value in inputValueList 
                    let existingLookup = lookupList.Find(l => l.CodeValue == value) 
                    select existingLookup ?? new GenericLookUp {CodeValue = value, ID = null})
                    .ToList();
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

        private static List<string> getColumnsFromSheet(OpenXmlElement row)
        {
            return row.Elements<Cell>().Select(getCellText).ToList();
        }

        private static IEnumerable<string> getColumnsByFormatVersion(string p1, string p2)
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

        private static void addValidationError(string criteriaCode, string info)
        {
            //TODO: criteria code will not always correctly map to field codes
            //TODO: where did "ILUV" error code come from? 

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
