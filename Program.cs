using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ImportPOC2.Models;
using ImportPOC2.Processors;
using Newtonsoft.Json;
using Radar.Core.Models.Batch;
using Radar.Data;
using Radar.Models;
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
using CriteriaSetCodeValue = Radar.Models.Criteria.CriteriaSetCodeValue;
using CriteriaSetValue = Radar.Models.Criteria.CriteriaSetValue;
using MediaCitation = Radar.Models.Company.MediaCitation;
using MediaCitationReference = Radar.Models.Company.MediaCitationReference;
using Path = System.IO.Path;
using SetCodeValue = Radar.Models.Criteria.SetCodeValue;

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
        private static PriceProcessor _priceProcessor;

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
            var productLevelFieldsProcessor = new ProductLevelFieldsProcessor(_curProdRow, _firstRowForProduct, _currentProduct, _publishCurrentProduct, _curBatch);
            productLevelFieldsProcessor.ProcessProductLevelFields();                       
            processSimpleLookups();
            processColorsMaterials();
            processSizes();
            processOptions();
            //NOTE: the following must be processed after the above since they depend upon configuration of the product
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
            _priceProcessor.ProcessPriceRow(_curProdRow, _currentProduct);
        }

        private static void processOptions()
        {
            //TODO: VNI-7
        }

        private static void processSizes()
        {
            //TODO: VNI-6
            //todo: will need to look at both size type and value columns to know what to do.
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
            processPersonalization(_curProdRow.Personalization);         
            processImprintSizeLocation(_curProdRow.Imprint_Size, _curProdRow.Imprint_Location);           
            processAdditionalColors(_curProdRow.Additional_Color);
            processAdditionalLocations(_curProdRow.Additional_Location);
            processProductSample(_curProdRow.Product_Sample);  
            processSpecSample(_curProdRow.Spec_Sample);           
            //production time
            processProductionTime(_curProdRow.Production_Time); 
            //rush service
            //rush time
            //same day
            processPackagingOptions(_curProdRow.Packaging);
            processShippingItems(_curProdRow.Shipping_Items);
            processShippingDimensions(_curProdRow.Shipping_Dimensions);
            processShippingWeight(_curProdRow.Shipping_Weight);           
            processComplianceCertifications(_curProdRow.Comp_Cert);           
            processSafetyWarnings(_curProdRow.Safety_Warnings);
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
                _priceProcessor = new PriceProcessor();
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
        
        private static void processShapes(string text)
        {
            lookupFieldProcessor(text, Constants.CriteriaCodes.Shape, Lookups.ShapesLookup);
        }

        private static void lookupFieldProcessor(string text, string criteriaCode, List<GenericLookUp> lookup)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                //split the values, if it's csv 
                var sheetValueList = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                var existingCsValues = criteriaSet.CriteriaSetValues.ToList();

                var sheetCodeValuesToValidate = parseByCriteriaCode(criteriaCode, sheetValueList);
                var matchedList = validateLookupValues(sheetCodeValuesToValidate, lookup);

                matchedList.ForEach(sheetValue =>
                {
                    if (!sheetValue.ID.HasValue)
                    {
                       handleValueExistenceByCode(criteriaCode, sheetValue);
                    }

                    if (sheetValue.ID.HasValue)
                    {
                        // ReSharper disable once PossibleInvalidOperationException
                        var existing = getCsValueBySetCodeValueId(sheetValue.ID.Value, existingCsValues);
                        if (existing == null)
                        {
                            //create it
                            createNewValue(criteriaCode, sheetValue.CodeValue, sheetValue.ID.Value);
                        }
                        else
                        {
                            //update existing value object? 
                            updateCsValue(criteriaCode, sheetValue.CodeValue, sheetValue.ID.Value);
                        }
                    }
                });

                deleteCsValues(existingCsValues, sheetValueList, criteriaSet);
            }
        }

        private static void lookupFieldProcessor_Tradenames(string text, string criteriaCode)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                //split the values, if it's csv 
                var sheetValueList = text.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                var existingCsValues = criteriaSet.CriteriaSetValues.ToList();                

                sheetValueList.ForEach(sheetValue =>
                {
                    var results = DataFetchers.Lookup.GetMatchingTradenames(sheetValue);
                    var tradenameFound = results.FirstOrDefault();

                    if (tradenameFound == null)
                    {
                        handleValueExistenceByCode(criteriaCode, new GenericLookUp { CodeValue = sheetValue });
                    }
                    else
                    {
                        var exists = existingCsValues.Any(v => v.BaseLookupValue.ToLower() == sheetValue.ToLower());
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            if (tradenameFound.ID != null)
                                createNewValue(criteriaCode, sheetValue, tradenameFound.ID.Value);
                        }
                    }                                                              
                });

                deleteCsValues(existingCsValues, sheetValueList, criteriaSet);
            }
        }

        /// <summary>
        /// "convert" incoming sheet values into non-encoded values stripped of aliases, special formatting, etc.
        /// the output is expected to be list that will be checked against a valid code_value list from Look_SetCodeValue
        /// </summary>
        /// <param name="criteriaCode"></param>
        /// <param name="sheetValueList"></param>
        /// <returns></returns>
        private static IEnumerable<string> parseByCriteriaCode(string criteriaCode, List<string> sheetValueList)
        {
            var retVal = sheetValueList;
            //NOTE: only make exceptions here, lave out fields that have no special processing, such as Shape, theme, etc.
            switch (criteriaCode)
            {
                case "IMMD":
                    //format is simply code_value=alias, so strip off after "="
                    retVal = sheetValueList.Select(s => s.Split('=').First().Trim()).ToList();
                    break;
            }

            return retVal;
        }

        /// <summary>
        /// this method should either log an error (such as for SHAP) or perform whatever is necessary to 
        /// handle when the user has provided a value that didn't match our lookups
        /// for imprint method, we reassign their entry to "Other"
        /// </summary>
        /// <param name="criteriaCode"></param>
        /// <param name="sheetValue"></param>
        /// <returns></returns>
        private static bool handleValueExistenceByCode(string criteriaCode, GenericLookUp sheetValue)
        {
            var retVal = true;
            switch (criteriaCode)
            {
                case "SHAP":
                case "THEM":
                case "TDNM":
                case "ORGN":
                    //for these sets if the sheet value can't be matched, it's a field-level validation error. 
                    addInvalidValueError(sheetValue.CodeValue, criteriaCode);
                    retVal = false;
                    break;

                case "IMMD":
                    //in this case, we use "Other" set code, then add it
                    var otherImprintMethodSetCodeId = Lookups.ImprintMethodsLookup.FirstOrDefault(i => i.CodeValue == "Other");
                    if (otherImprintMethodSetCodeId != null)
                        sheetValue.ID = otherImprintMethodSetCodeId.ID;
                    break;
                case "PERS":
                    //in this case, we use "Other" set code, then add it
                    //sheetValue.ID == otherPersonalizationSetCodeId;
                    break;
                case "PCKG":
                    var customPkgScv = Lookups.PackagingLookup.FirstOrDefault(p => p.Value == "Custom");
                    if (customPkgScv != null)
                    {
                        createNewValue(criteriaCode, sheetValue.CodeValue, customPkgScv.Key, "CUST");
                    }
                    break;
                case "PRCL":
                    break;
                case "MTRL":
                    break;
            }

            return retVal;
        }

        //use this method to update an existing CSVs properites, such as comments (criteriaValueDetail field)
        private static void updateCsValue(string criteriaCode, string criteriaValue, long setCodeValueId, string criteriaDetail = "")
        {
            //use criteria code to ensure we do the right type of update
            switch (criteriaCode)
            {
                case "SHAP":
                case "THEM":
                case "TDNM":
                case "ORGN":
                    //do nothing, you can't update existing values in these sets
                    break;

            }
        }

        private static void addInvalidValueError(string invalidValue, string criteriaCode)
        {
            //add to batch error log that the specified value does not match a value in lookup
        }

        private static void genericProcess(string text, string criteriaCode, IEnumerable<GenericLookUp> lookup)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

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
            if (string.IsNullOrWhiteSpace(text))
                return;

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
            if (string.IsNullOrWhiteSpace(text))
                return;

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
            if (string.IsNullOrWhiteSpace(text))
                return;

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

        //this generic method will handle the processing for imprint methods and personalization
        private static void genericProcessImprintMethods(string text, string criteriaCode, IEnumerable<SetCodeValue> lookup)
        {            
            //comma separated list of values
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var valueList = text.ConvertToList();                
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

                valueList.ForEach(value =>
                {                    
                    var splittedValue = value.SplitValue('=');
                    var existing = Lookups.ImprintMethodsLookup.FirstOrDefault(l => l.CodeValue.ToLower() == splittedValue.CodeValue.ToLower());
                    if (existing != null)
                    {
                        var exists = existingCsvalues.Any(csv => csv.Value == splittedValue.Alias);
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            createNewValue(criteriaCode, splittedValue.Alias, existing.ID);
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

        //this generic method will handle the processing for product and spec samples
        private static void genericProcessSamples(string text, string criteriaCode, IEnumerable<ImprintCriteriaLookUp> lookup, string sampleType)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            //comma separated list of values
            if (_firstRowForProduct)
            {
                //should have only one value
                var splittedValue = text.SplitValue(':');
                //should have Y/N as the first value                
                var validValues = new[] { "Y", "N" };
                if (!string.IsNullOrWhiteSpace(splittedValue.CodeValue) && !validValues.Contains(splittedValue.CodeValue))
                {
                    addValidationError("GNER", "invalid value for Samples");
                    return;
                }

                var criteriaSet = getCriteriaSetByCode("SMPL");   

                if (splittedValue.CodeValue == Constants.BooleanFlag.TRUE)
                {                                                         
                    //check if the sample value already exists
                    var csValue = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.Value == sampleType);
                    if (csValue != null)
                    {
                        csValue.CriteriaValueDetail = splittedValue.Alias;
                    }
                    else
                    {
                        var criteria = Lookups.ImprintCriteriaLookup.FirstOrDefault(c => c.Code == "SMPL");
                        if (criteria != null)
                        {
                            var group = criteria.CodeValueGroups.FirstOrDefault();
                            if (group != null)
                            {
                                //get set code value based on sampleType param; this will be "Product Sample" or "Spec Sample"
                                var setCodeValue = group.SetCodeValues.FirstOrDefault(s => s.CodeValue == sampleType);
                                if (setCodeValue != null)
                                {
                                    long smplScvId = setCodeValue.ID;
                                    createNewValue("SMPL", sampleType, smplScvId);
                                }
                            }
                        }
                    }                                       
                } 
                else
                {
                    var existingValue = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.Value == sampleType);
                    if (existingValue != null)
                    {
                        criteriaSet.CriteriaSetValues.Remove(existingValue);
                    }  
                }               
            }
        }

        private static void processThemes(string text)
        {
            //TODO: ensure we this the lookup as generic direct from Lookups object instead of converting here. 
            lookupFieldProcessor(text, Constants.CriteriaCodes.Theme, Lookups.ThemesLookup);
        }

        private static void processTradenames(string text)
        {                     
            lookupFieldProcessor_Tradenames(text, Constants.CriteriaCodes.TradeName);


            if (string.IsNullOrWhiteSpace(text))
                return;

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
                        var exists = existingCsvalues.Any(v => string.Equals(v.Value, tradename, StringComparison.InvariantCultureIgnoreCase));
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            if (tradenameFound.ID != null)
                                createNewValue(criteriaCode, tradename, tradenameFound.ID.Value);
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

        private static void processOrigins(string text)
        {
            lookupFieldProcessor(text, "ORGN", Lookups.OriginsLookup);
        }

        private static void processShippingItems(string text)
        {
            if (_firstRowForProduct)
            {
                var shippingItems = text.Split(':');
                if (shippingItems.Length == 2)
                {
                    var criteriaCode = "SHES";                
                    var criteriaSet = getCriteriaSetByCode(criteriaCode);
                    var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
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
        
        private static void processShippingDimensions(string text)
        {
            if (_firstRowForProduct)
            {
                var criteriaCode = "SDIM";
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                //var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                var dimensionTypes = new string[] {"Length", "Width", "Height"};
                var shippingDimensions = text.Split(';');

                for (var i = 0; i < shippingDimensions.Length; i++)
                {
                    processDimension(criteriaCode, dimensionTypes[i], shippingDimensions[i]);
                }              
            }
        }

        private static void processDimension(string criteriaCode, string dimentionType, string dimensionUnitValue)
        {
            if (!string.IsNullOrWhiteSpace(dimensionUnitValue))
            {
                var dimensionValues = dimensionUnitValue.Split(':');

                if (dimensionValues.Length == 2)
                {
                    var criteriaSet = getCriteriaSetByCode(criteriaCode);
                    var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

                    var dimension = dimensionValues[0];
                    var unit = dimensionValues[1];

                    var criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, dimentionType);
                    var unitFound = criteriaAttribute.UnitsOfMeasure.FirstOrDefault(u => u.Format == unit);
                    if (unitFound != null)
                    {
                        var value = new
                        {
                            CriteriaAttributeId = criteriaAttribute.ID,
                            UnitValue = dimension,
                            UnitOfMeasureCode = unit
                        };

                        if (!existingCsvalues.Any())
                        {
                            //add new value if it doesn't exist                       
                            var group = criteriaAttribute.CriteriaItem.CodeValueGroups.FirstOrDefault();
                            var setCodeValueId = 0L;
                            if (group != null)
                            {
                                var setCodeValue = group.SetCodeValues.FirstOrDefault();
                                if (setCodeValue != null)
                                    setCodeValueId = setCodeValue.ID;
                            }
                            var valueList = new List<dynamic> {value};
                            createNewValue(criteriaCode, valueList, setCodeValueId, "CUST");
                        }
                        else
                        {
                            var criteriaSetValue = existingCsvalues.FirstOrDefault();
                            if (criteriaSetValue != null)
                                criteriaSetValue.Value.Add(value);
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

        private static void processShippingWeight(string text)
        {
            if (_firstRowForProduct)
            {
                processDimension("SHWT", "Unit", text);               
            }
        }      

        private static void processPackagingOptions(string text)
        {
            //TODO: ensure we this the lookup as generic direct from Lookups object instead of converting here. 
            var packagingOptionsAsGeneric = new List<GenericLookUp>();
            packagingOptionsAsGeneric.AddRange(Lookups.PackagingLookup.Select(s => new GenericLookUp { CodeValue = s.Value, ID = s.Key }));

            lookupFieldProcessor(text, Constants.CriteriaCodes.Packaging, packagingOptionsAsGeneric);            
        }       

        private static void processImprintSizeLocation(string imprintSizeText, string imprintLocationText)
        {
            //comma delimited list of imprint size location
            if (_firstRowForProduct)
            {
                var criteriaCode = Constants.CriteriaCodes.ImprintSizeLocation;
                var imprintSizes = imprintSizeText.ConvertToList();
                var imprintLocation = imprintLocationText.ConvertToList();
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                long customImprintSizeLocationScvId = 0;

                var imsz = Lookups.ImprintSizeLocationLookup.FirstOrDefault();
                //get set code value id for imprint size location
                if (imsz != null)
                {
                    var group = imsz.CodeValueGroups.FirstOrDefault();
                    if (group != null)
                    {
                        var setCodeValue = group.SetCodeValues.FirstOrDefault();
                        if (setCodeValue != null)
                        {
                            customImprintSizeLocationScvId = setCodeValue.ID;
                        }
                    }
                }

                //TODO: handle add new values for imprint size location

                //TODO: handle delete for imprint size locations
            }           
        }

        private static void processImprintMethods(string text)
        {
            genericProcessImprintMethods(text, Constants.CriteriaCodes.ImprintMethod, Lookups.ImprintMethodsLookup);                       

            //var imprintMethodsAsGeneric = new List<GenericLookUp>();
            //imprintMethodsAsGeneric.AddRange(Lookups.ImprintMethodsLookup.Select(s => new GenericLookUp { CodeValue = s.CodeValue, ID = s.ID }));

            //lookupFieldProcessor(text, Constants.CriteriaCodes.ImprintMethod, imprintMethodsAsGeneric);
        }

        private static void processPersonalization(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            genericProcessImprintMethods(text, "PERS", Lookups.PersonalizationLookup);
            var pers = getCriteriaSetByCode("PERS");

            if (pers != null && pers.CriteriaSetValues.Any())
            {
                _currentProduct.IsPersonalizationAvailable = true;
            }
            else
            {
                _currentProduct.IsPersonalizationAvailable = false;
            }           
        }               

        private static void processImprintColors(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

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
            if (string.IsNullOrWhiteSpace(text))
                return;

            if (_firstRowForProduct)
            {
                //should be Y/N
                var soldUnimprinted = text;
                var validValues = new[] { "Y", "N" };
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
                    soldUnimprintedScvId = Convert.ToInt64(unimprinted.ID);
                }

                var criteriaCode = Constants.CriteriaCodes.ImprintMethod;
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                CriteriaSetValue unimprintedCsvalue = null;

                if (criteriaSet != null)
                {
                    unimprintedCsvalue = getCsValueBySetCodeValueId(soldUnimprintedScvId, criteriaSet.CriteriaSetValues);
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
            //comma delimited list of imprint artworks
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var valueList = text.ConvertToList();               
                var criteriaCode = Constants.CriteriaCodes.Artwork;
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();  
                var otherArtworkScValue = Lookups.ArtworkLookup.FirstOrDefault(a => a.CodeValue == "Other");
                List<FieldInfo> modelValues = new List<FieldInfo>();
                long otherArtworkScValueId = 0;

                if (otherArtworkScValue != null)
                {
                    if (otherArtworkScValue.ID != null)
                        otherArtworkScValueId = otherArtworkScValue.ID.Value;
                }

                valueList.ForEach(item =>
                {
                    var splittedValue = item.SplitValue(':');
                    modelValues.Add(splittedValue);
                    var exists = Lookups.ArtworkLookup.FirstOrDefault(l => String.Equals(l.CodeValue, splittedValue.CodeValue, StringComparison.CurrentCultureIgnoreCase));

                    if (exists != null)
                    {
                        CriteriaSetValue existingCsValue = null;
                        if (exists.ID != otherArtworkScValueId)
                        {
                            if (exists.ID != null)
                                existingCsValue = getCsValueBySetCodeValueId(exists.ID.Value, criteriaSet.CriteriaSetValues);
                        }
                        else
                        {
                            if (exists.ID != null)
                                existingCsValue = findCriteriaValue(exists.ID.Value, criteriaSet, splittedValue.Alias);
                        }

                        //add new value if it doesn't exists
                        if (existingCsValue == null)
                        {
                            var value = exists.ID != null && exists.ID.Value != otherArtworkScValueId ? splittedValue.CodeValue : splittedValue.Alias;
                            if (exists.ID != null)
                                createNewValue(criteriaCode, value, exists.ID.Value, "CUST", splittedValue.Alias);
                        }
                        else
                        {
                            //update value if it exists
                            existingCsValue.CriteriaValueDetail = splittedValue.Alias;
                        }
                    }
                });

                //delete values that are missing from the list in the file
                var csValuesToDelete = new List<CriteriaSetValue>();
                existingCsvalues.ForEach(e =>
                {
                    var exists = modelValues.FirstOrDefault(m => m.CodeValue == e.Value && m.Alias == e.CriteriaValueDetail);
                    if (exists == null)
                    {
                        criteriaSet.CriteriaSetValues.Remove(e);
                    }
                });                                              
            }
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

        private static void processProductSample(string text)
        {
            genericProcessSamples(text, "SMPL", Lookups.ImprintCriteriaLookup, "Product Sample");
        }

        private static void processSpecSample(string text)
        {
            genericProcessSamples(text, "SMPL", Lookups.ImprintCriteriaLookup, "Spec Sample");
        }

        private static void processProductionTime(string text)
        {
            if (_firstRowForProduct)
            {
                var criteriaCode = Constants.CriteriaCodes.ProductionTime;
                var productionTimes = new List<FieldInfo>();
                var productionTimesTokens = text.ConvertToList();
                long customSetCodeValueId = 0;
                var criteriaSet = getCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();                
               
                var criteriaLookUp = Lookups.ProductionTimeCriteriaLookup.FirstOrDefault(i => i.Code == criteriaCode);

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

                productionTimesTokens.ForEach(token => productionTimes.Add(token.SplitValue(':')));

                productionTimes.ForEach(productionTime =>
                {
                    var comment = string.Empty;

                    string time = productionTime.CodeValue;

                    if (!string.IsNullOrWhiteSpace(productionTime.Alias))
                    {
                        comment = productionTime.Alias;
                    }

                    var exists = existingCsvalues.FirstOrDefault(v => !(v.Value is string) && v.Value != null && v.Value.First.UnitValue == time && v.CriteriaValueDetail == comment);
                    //add new value if it doesn't exists
                    if (exists == null)
                    {
                        var value = new 
                        {
                              CriteriaAttributeId = 13,
                              UnitValue = time,
                              UnitOfMeasureCode = "BUSI"
                        };

                         createNewValue(criteriaCode, value, customSetCodeValueId, "CUST", comment);
                    }                   
                });                

                deleteCsValues(existingCsvalues, productionTimes, criteriaSet);               
            }
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
            if (string.IsNullOrWhiteSpace(text))
                return;

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

        private static void processCatalogInfo(string text)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                //should only be one
                var catalog = text.ConvertToList();
                var existingCatalog = _currentProduct.ProductMediaCitations;

                if (catalog != null && existingCatalog != null)
                {
                    var cat = catalog.FirstOrDefault();
                    if (cat != null)
                    {
                        var splittedCat = cat.Split(':');
                        var foundMediaCitation = FindMediaCitation(splittedCat);

                        if (foundMediaCitation != null)
                        {
                            //see if this catalog is already associated with the product

                            var exists = _currentProduct.ProductMediaCitations.FirstOrDefault(m => m.MediaCitationId == foundMediaCitation.ID);
                            if (exists != null)
                            {
                                if (splittedCat.Length == 3 && !string.IsNullOrWhiteSpace(splittedCat[2]))
                                {
                                    var mediaRef = exists.ProductMediaCitationReferences.FirstOrDefault();
                                    if (mediaRef != null)
                                    {
                                        mediaRef.MediaCitationReference.Number = splittedCat[2];
                                    }
                                }
                            }
                            else
                            {
                                //per VELO-2863 only one catalog can be associated with the product
                                //check if the product already has a catalog then replace with the new one
                                if (_currentProduct.ProductMediaCitations.Any())
                                    _currentProduct.ProductMediaCitations.Clear();

                                //create new media citation and add it to the current product's collection
                                var newProductMediaCitation = new ProductMediaCitation();
                                newProductMediaCitation.ProductId = _currentProduct.ID;
                                newProductMediaCitation.Description = foundMediaCitation.Name;
                                newProductMediaCitation.MediaCitationId = foundMediaCitation.ID;

                                var newProductMediaCitationReference = new ProductMediaCitationReference();
                                newProductMediaCitationReference.MediaCitationReference = new MediaCitationReference();
                                newProductMediaCitationReference.MediaCitationReference.Number = splittedCat.Length == 3 ? splittedCat[2] : string.Empty;
                                newProductMediaCitationReference.MediaCitationId = foundMediaCitation.ID;

                                newProductMediaCitation.ProductMediaCitationReferences = new List<ProductMediaCitationReference>();
                                newProductMediaCitation.ProductMediaCitationReferences.Add(newProductMediaCitationReference);
                                _currentProduct.ProductMediaCitations.Add(newProductMediaCitation);
                            }
                        }
                    }
                }
            }
        }

        private static void processMaterials(string text)
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

        //TODO: can we "detect" what value type code to use instead of passing it in? 
        //TODO: pass in criteria set, it's known from everywhere it is invoked
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
                CriteriaValueDetail = valueDetail,
                FormatValue = value.ToString() //default formatvalue to be same as value
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

        private static CriteriaSetValue getCsValueBySetCodeValueId(long scvId, IEnumerable<CriteriaSetValue> criteriaSetValues)
        {
            return (from v in criteriaSetValues 
                    let scv = v.CriteriaSetCodeValues.FirstOrDefault(s => s.SetCodeValueId == scvId) 
                    where scv != null 
                    select v)
                    .FirstOrDefault();
        }

        //private static CriteriaSetValue getCsValueBySetCodeValueId(long scvId, ProductCriteriaSet criteriaSet)
        //{           
        //    CriteriaSetValue retVal = null;

        //    if (criteriaSet != null)
        //    {                               
        //        foreach (var v in 
        //            from v in criteriaSet.CriteriaSetValues 
        //            let scv = v.CriteriaSetCodeValues.FirstOrDefault(s => s.SetCodeValueId == scvId) 
        //            where scv != null select v)
        //        {                    
        //            retVal = v;
        //            break;
        //        }
        //    }

        //    return retVal; 
        //}

        private static CriteriaSetValue findCriteriaValue(long scvId, ProductCriteriaSet criteriaSet, string criteriaValue)
        {
            return (from v in criteriaSet.CriteriaSetValues 
                    let scv = v.CriteriaSetCodeValues.FirstOrDefault(s => s.SetCodeValueId == scvId) 
                    where scv != null 
                        && v.Value == criteriaValue 
                    select v)
                    .FirstOrDefault();
        }

        private static MediaCitation FindMediaCitation(string[] catalogInfo)
        {
            MediaCitation retVal = null;

            //catalogInfo is a tokenized array with 3 elements: [0] = Catalog name; [1] = Year; [2] = Page number            
            string name = catalogInfo[0];
            string year = catalogInfo[1];
            string pageNumber = string.Empty;

            if (catalogInfo.Length == 3)
            {
                pageNumber = catalogInfo[2];
            }

            var mediaCitation = Lookups.MediaCitations.FirstOrDefault(m => m.Name == name && m.Year == year);
            if (mediaCitation != null)
            {
                if (pageNumber != string.Empty)
                {
                    var mediaCitationReference = mediaCitation.MediaCitationReferences.FirstOrDefault(r => r.Number == pageNumber);
                    if (mediaCitationReference != null)
                    {
                        retVal = mediaCitation;
                    }
                }
                else
                {
                    retVal = mediaCitation;
                }
            }

            return retVal;                
        }

        //TODO: pretty sure this can be done without passing in criteriaset parameter
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

        private static void deleteCsValues(IEnumerable<CriteriaSetValue> entities, IEnumerable<FieldInfo> models, ProductCriteriaSet criteriaSet)
        {
            //delete values that are missing from the list in the file
            var csValuesToDelete = new List<CriteriaSetValue>();
            entities.ToList().ForEach(e => {
                if (!(e.Value is string))
                {
                    var exists = models.FirstOrDefault(m => string.Equals(m.CodeValue, e.Value.First.UnitValue.ToString(), StringComparison.InvariantCultureIgnoreCase) && m.Alias == e.CriteriaValueDetail);
                    if (exists == null)
                    {
                        csValuesToDelete.Add(e);
                    }
                }
            });

            csValuesToDelete.ForEach(e =>
            {
                var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v == e);
                criteriaSet.CriteriaSetValues.Remove(toDelete);
            });
        }

        //private static List<string> extractCsvList(string text)
        //{
        //    //returns list of strings from input string, split on commas, each value is trimmed, and only non-empty values are returned.
        //    //return text.Split(',').Select(str => str.Trim()).Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
        //    return text.ConvertToList();
        //}        

        private static void finishProduct()
        {
            //if we've started a radar model, 
            // we "send" the product to Radar for processing. 
            if (_currentProduct != null && !_hasErrors)
            {
                _priceProcessor.Finalize();
                //TODO: other repeatable sets will "finalize" here as well. 

                //var x = _currentProduct;
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
