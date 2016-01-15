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
using CriteriaSetValue = Radar.Models.Criteria.CriteriaSetValue;
using MediaCitation = Radar.Models.Company.MediaCitation;
using MediaCitationReference = Radar.Models.Company.MediaCitationReference;
using Path = System.IO.Path;
using SetCodeValue = Radar.Models.Criteria.SetCodeValue;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace ImportPOC2
{
    public class Program
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
        private static bool _publishCurrentProduct = true;

        private static log4net.ILog _log;
        private static bool _hasErrors = false;
        private static ProductRow _curProdRow;
        private static PriceProcessor _priceProcessor;
        private static CriteriaProcessor _criteriaProcessor;

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
                    finishProduct();

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
            if (string.IsNullOrWhiteSpace(_curProdRow.Size_Group) || string.IsNullOrWhiteSpace(_curProdRow.Size_Values))
            {
                addValidationError("SIZE", "size group/values must be provided");
                return;
            }

            //determine size group
            var sizeType = StaticLookups.SizeTypes.FirstOrDefault(s => s.Value == _curProdRow.Size_Group);
            if (sizeType == null)
            {
                addValidationError("SIZE", "invalid size group found");
                return;
            }

            if (sizeType.Code == Constants.CriteriaCodes.SIZE_CAPS || sizeType.Code == Constants.CriteriaCodes.SIZE_SVWT || sizeType.Code == Constants.CriteriaCodes.SIZE_DIMS)
            {


            }
            else
            {                               
                var sizeCs = _criteriaProcessor.getSizeCriteriaSetByCode(sizeType.Code);
                var sizeSetCodeValues = Lookups.SizesLookup.Where(s => s.CriteriaCode == sizeType.Code).ToList();
                lookupFieldProcessor(_curProdRow.Size_Values, sizeType.Code, sizeSetCodeValues, sizeCs);
            }            
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
            processProductionTime(_curProdRow.Production_Time);
            processRushService(_curProdRow.Rush_Service);          
            processRushTime(_curProdRow.Rush_Time);
            processSameDay(_curProdRow.Same_Day_Service);           
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
                _log.DebugFormat("staring product {0}", _curXid);
                //using current XID, check if product exists, otherwise create new empty model 
                _currentProduct = getProductByXid() ?? new Product { CompanyId = _companyId };
                _firstRowForProduct = true;
                _hasErrors = false;
                _criteriaProcessor = new CriteriaProcessor(_currentProduct);
                _priceProcessor = new PriceProcessor(_criteriaProcessor);

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

        private static void lookupFieldProcessor(string text, string criteriaCode, List<GenericLookUp> lookup, ProductCriteriaSet cs = null)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                //split the values, if it's csv 
                var sheetValueList = text.ConvertToList();
                var criteriaSet = cs == null ? _criteriaProcessor.GetCriteriaSetByCode(criteriaCode) : cs;                                             
                var existingCsValues = criteriaSet.CriteriaSetValues.ToList();

                var sheetCodeValuesToValidate = parseByCriteriaCode(criteriaCode, sheetValueList);
                var matchedList = validateLookupValues(sheetCodeValuesToValidate, lookup);

                matchedList.ForEach(sheetValue =>
                {
                    if (!sheetValue.ID.HasValue)
                    {
                        handleValueExistenceByCode(criteriaCode, sheetValue, existingCsValues);
                    }

                    if (sheetValue.ID.HasValue)
                    {
                        // ReSharper disable once PossibleInvalidOperationException
                        var existing = _criteriaProcessor.getCsValueBySetCodeValueId(sheetValue.ID.Value, existingCsValues);
                        if (existing == null)
                        {
                            //create it
                            _criteriaProcessor.CreateNewValue(criteriaCode, sheetValue.CodeValue, sheetValue.ID.Value);
                        }
                        else
                        {
                            //update existing value object? 
                            updateCsValue(criteriaCode, existing, sheetValue.CodeValue, sheetValue.ID.Value);
                        }
                    }
                });

                _criteriaProcessor.DeleteCsValues(existingCsValues, sheetValueList, criteriaSet);
            }
        }

        private static void lookupFieldProcessor_Tradenames(string text, string criteriaCode)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                //split the values, if it's csv 
                var sheetValueList = text.ConvertToList();
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
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
                        var exists = existingCsValues.Any(v => string.Equals(v.BaseLookupValue.ToString(), sheetValue, StringComparison.CurrentCultureIgnoreCase));
                        //add new value if it doesn't exists
                        if (!exists)
                        {
                            if (tradenameFound.ID != null)
                                _criteriaProcessor.CreateNewValue(criteriaCode, sheetValue, tradenameFound.ID.Value);
                        }
                    }                                                              
                });

                _criteriaProcessor.DeleteCsValues(existingCsValues, sheetValueList, criteriaSet);
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
        private static bool handleValueExistenceByCode(string criteriaCode, GenericLookUp sheetValue, IEnumerable<CriteriaSetValue> csValues = null)
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
                    var otherImprintMethodSetCodeId = Lookups.ImprintMethodsLookup.FirstOrDefault(i => string.Equals(i.CodeValue, "Other", StringComparison.CurrentCultureIgnoreCase));
                    if (otherImprintMethodSetCodeId != null)
                        sheetValue.ID = otherImprintMethodSetCodeId.ID;
                    break;
                case "PERS":
                    //in this case, we use "Other" set code, then add it
                    //sheetValue.ID == otherPersonalizationSetCodeId;
                    break;
                case "PCKG":
                    var customPkgScv = Lookups.PackagingLookup.FirstOrDefault(p => string.Equals(p.Value ,"Custom", StringComparison.CurrentCultureIgnoreCase));
                    if (customPkgScv != null)
                    {
                        _criteriaProcessor.CreateNewValue(criteriaCode, sheetValue.CodeValue, customPkgScv.Key, "CUST");
                    }
                    break;
                case "PRCL":
                    break;
                case "MTRL":
                    break;

                case "SABR":
                case "SANS":
                case "SAHU":
                case "SAIT":
                case "SAWI":
                case "SSNM":
                case "SVWT":
                case "CAPS":
                case "DIMS":
                case "SOTH":                    
                    //this will be a custom value
                    //if the custom value doesn't exists already then create the new value
                    var sizeIds = Lookups.SizeIdsLookup.FirstOrDefault(s => s.CriteriaCode == criteriaCode);                   
                    if (sizeIds != null)
                    {
                        var existingCsValue = _criteriaProcessor.getCsValueByFormatValue(sizeIds.CustomSetCodeValueId, csValues, sheetValue.CodeValue);
                        if (existingCsValue == null)
                        {                            
                            var value = new
                            {
                                CriteriaAttributeId = sizeIds.CriteriaAttributeId,
                                UnitValue = sheetValue.CodeValue,
                                UnitOfMeasureCode = ""                                
                            }; 

                            _criteriaProcessor.CreateNewValue(criteriaCode, value, sizeIds.CustomSetCodeValueId, "CUST");
                        }
                    }
                    break;
            }

            return retVal;
        }

        //use this method to update an existing CSVs properites, such as comments (criteriaValueDetail field)
        private static void updateCsValue(string criteriaCode, CriteriaSetValue csValue, string criteriaValue, long setCodeValueId, string criteriaDetail = "")
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
                
                case "IMMD":
                    var cscv = csValue.CriteriaSetCodeValues.FirstOrDefault();
                    if (cscv != null)                    
                        cscv.SetCodeValueId = setCodeValueId;                    
                    break;
            }
        }

        private static void addInvalidValueError(string invalidValue, string criteriaCode)
        {
            //add to batch error log that the specified value does not match a value in lookup
        }

        private static void genericProcessImprintCriteria(string text, string criteriaCode)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var valueList = text.ConvertToList();
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                long customSetCodeValueId = 0;

                //get other set code value id
                var imprintCriteria = Lookups.ImprintCriteriaLookup.FirstOrDefault(i => string.Equals(i.Code, criteriaCode, StringComparison.CurrentCultureIgnoreCase));
                if (imprintCriteria != null)
                {                   
                    var group = imprintCriteria.CodeValueGroups.FirstOrDefault(cvg => string.Equals(cvg.Description, "Other", StringComparison.CurrentCultureIgnoreCase));
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
                    var exists = existingCsvalues.Any(csv => string.Equals(csv.Value.ToString(), value, StringComparison.CurrentCultureIgnoreCase));
                    if (!exists)
                    {
                        //add new value if it doesn't exist
                        _criteriaProcessor.CreateNewValue(criteriaCode, value, customSetCodeValueId, "CUST");
                    }
                });

                _criteriaProcessor.DeleteCsValues(existingCsvalues, valueList, criteriaSet);
            }
        }

        //this generic method will handle the processing for imprint methods and personalization
        private static void genericProcessImprintMethods(string text, string criteriaCode, IEnumerable<SetCodeValue> lookup)
        {            
            //comma separated list of values
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var valueList = text.ConvertToList();                
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                var modelValues = new List<string>();

                valueList.ForEach(value =>
                {                    
                    var splittedValue = value.SplitValue('=');
                    modelValues.Add(splittedValue.Alias);
                    var existing = Lookups.ImprintMethodsLookup.FirstOrDefault(l => String.Equals(l.CodeValue, splittedValue.CodeValue, StringComparison.CurrentCultureIgnoreCase));
                    if (existing != null)
                    {
                        var existingCsValue = existingCsvalues.FirstOrDefault(csv => string.Equals(csv.Value.ToString(), splittedValue.Alias, StringComparison.CurrentCultureIgnoreCase));
                        //add new value if it doesn't exists
                        if (existingCsValue ==  null)
                        {
                            _criteriaProcessor.CreateNewValue(criteriaCode, splittedValue.Alias, existing.ID);
                        }
                        else
                        {
                            updateCsValue(criteriaCode, existingCsValue, "", existing.ID);
                        }                       
                        //NOTE alias cannot be "updated" - an alias change triggers a delete then add of new CSV
                    }
                    else
                    {
                        //log batch error
                        addValidationError(criteriaCode, value);
                        _hasErrors = true;
                    }
                });

                _criteriaProcessor.DeleteCsValues(existingCsvalues, modelValues, criteriaSet);
            }
        }

        //this generic method will handle the processing for product and spec samples
        private static void genericProcessSamples(string text, string criteriaCode, IEnumerable<ImprintCriteriaLookUp> lookup, string sampleType)
        {
            //comma separated list of values
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
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

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode("SMPL");

                if (splittedValue.CodeValue == Constants.BooleanFlag.TRUE)
                {
                    //check if the sample value already exists
                    var csValue = criteriaSet.CriteriaSetValues.FirstOrDefault(v => string.Equals(v.Value.ToString(), sampleType, StringComparison.CurrentCultureIgnoreCase));
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
                                var setCodeValue = group.SetCodeValues.FirstOrDefault(s => string.Equals(s.CodeValue, sampleType, StringComparison.CurrentCultureIgnoreCase));
                                if (setCodeValue != null)
                                {
                                    long smplScvId = setCodeValue.ID;
                                    _criteriaProcessor.CreateNewValue("SMPL", sampleType, smplScvId);
                                }
                            }
                        }
                    }
                }
                else
                {
                    var existingValue = criteriaSet.CriteriaSetValues.FirstOrDefault(v => string.Equals(v.Value.ToString(), sampleType, StringComparison.CurrentCultureIgnoreCase));
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
        }

        private static void processOrigins(string text)
        {
            lookupFieldProcessor(text, "ORGN", Lookups.OriginsLookup);
        }

        private static void processShippingItems(string text)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var shippingItems = text.Split(':');
                if (shippingItems.Length == 2)
                {
                    var criteriaCode = "SHES";
                    var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                    var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                    var items = shippingItems[0];
                    var unit = shippingItems[1];
                    var criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Unit");
                    var unitFound = criteriaAttribute.UnitsOfMeasure.FirstOrDefault(u => string.Equals(u.DisplayName, unit, StringComparison.CurrentCultureIgnoreCase));

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

                            _criteriaProcessor.CreateNewValue(criteriaCode, value, setCodeValueId, "CUST");
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
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var criteriaCode = "SDIM";              
                var dimensionTypes = new [] { "Length", "Width", "Height" };
                var shippingDimensions = text.Split(';');

                for (var i = 0; i < shippingDimensions.Length; i++)
                {
                    genericProcessDimension(criteriaCode, dimensionTypes[i], shippingDimensions[i]);
                }
            }
        }

        private static void genericProcessDimension(string criteriaCode, string dimentionType, string dimensionUnitValue)
        {
            if (!string.IsNullOrWhiteSpace(dimensionUnitValue))
            {
                var dimensionValues = dimensionUnitValue.SplitValue(':');

                if (dimensionValues.CodeValue != dimensionValues.Alias)
                {
                    var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                    var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

                    var dimension = dimensionValues.CodeValue;
                    var unit = dimensionValues.Alias;

                    var criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, dimentionType);
                    var unitFound = criteriaAttribute.UnitsOfMeasure.FirstOrDefault(u => string.Equals(u.Format, unit, StringComparison.CurrentCultureIgnoreCase));
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
                            var valueList = new List<dynamic> { value };
                            _criteriaProcessor.CreateNewValue(criteriaCode, valueList, setCodeValueId, "CUST");
                        }
                        else
                        {
                            try
                            {                                
                                var values = (criteriaSet.CriteriaSetValues.FirstOrDefault().Value as IEnumerable<dynamic>).ToList();
                                var exists = values.FirstOrDefault(v => !(v is string) && v != null && v.CriteriaAttributeId == value.CriteriaAttributeId);

                                if (exists != null)
                                {
                                   values.Remove(exists);
                                }

                                values.Add(value);
                                existingCsvalues.FirstOrDefault().Value = values;
                            }
                            catch (Exception ex)
                            {
                            }
                                                      
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
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                genericProcessDimension("SHWT", "Unit", text);
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
            if (_firstRowForProduct && (!string.IsNullOrWhiteSpace(imprintSizeText) || !string.IsNullOrWhiteSpace(imprintLocationText)))
            {
                var criteriaCode = Constants.CriteriaCodes.ImprintSizeLocation;
                var imprintSizes = imprintSizeText.ConvertToList();
                var imprintLocations = imprintLocationText.ConvertToList();
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                var modelValues = new List<string>();

                long customImprintSizeLocationScvId = 0;
                List<FieldInfo> imprintSizeLocationSet = new List<FieldInfo>();
                var i = 0;
                
                //using FieldInfo as a container for imprint size location set
                //CodeValue will contain value for imprint size
                //Alias will contain value for imprint location
                for(; i < imprintSizes.Count; i++)
                {
                    var imszObject = new FieldInfo();
                    imszObject.CodeValue = imprintSizes[i];

                    if (imprintLocations.Count > i)
                    {
                        imszObject.Alias = imprintLocations[i];
                    }

                    imprintSizeLocationSet.Add(imszObject);
                }

                if (imprintSizes.Count < imprintLocations.Count)
                {
                    for(; i < imprintLocations.Count; i++)
                    {
                        var imszObject = new FieldInfo();
                        imszObject.Alias = imprintLocations[i];

                        imprintSizeLocationSet.Add(imszObject);
                    }
                }

                //get set code value id for imprint size location
                var imsz = Lookups.ImprintSizeLocationLookup.FirstOrDefault();                
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

                imprintSizeLocationSet.ForEach(item =>
                {
                    var size = item.CodeValue ?? string.Empty;
                    var location = item.Alias ?? string.Empty;
                    var valueToMatch = size + "|" + location;     
          
                    modelValues.Add(valueToMatch);
                    //TODO: check that this works with all forms of imprint size/location (i.e., "s1|l1", "s1", "|l1", empty, etc.)
                    var exists = existingCsvalues.Any(csv => string.Equals(csv.Value.ToString(), valueToMatch, StringComparison.CurrentCultureIgnoreCase));

                    if (!exists)
                    {
                        _criteriaProcessor.CreateNewValue(criteriaCode, valueToMatch, customImprintSizeLocationScvId, "CUST");                            
                    }
                });

                _criteriaProcessor.DeleteCsValues(existingCsvalues, modelValues, criteriaSet);                                              
            }           
        }

        private static void processImprintMethods(string text)
        {
            genericProcessImprintMethods(text, Constants.CriteriaCodes.ImprintMethod, Lookups.ImprintMethodsLookup);                                   
        }

        private static void processPersonalization(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            genericProcessImprintMethods(text, "PERS", Lookups.PersonalizationLookup);
            var pers = _criteriaProcessor.GetCriteriaSetByCode("PERS");

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
            //comma delimited list of imprint colors
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var criteriaCode = Constants.CriteriaCodes.ImprintColor;
                var imprintColors = text.ConvertToList();
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);

                long imprintColorScvId = 0;

                //get set code value id for imprint color
                var imcl = Lookups.ImprintColorLookup.FirstOrDefault(i => i.Code == criteriaCode);

                if (imcl != null)
                {
                    var group = imcl.CodeValueGroups.FirstOrDefault(cvg => string.Equals(cvg.Description, "Other", StringComparison.CurrentCultureIgnoreCase));

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
                    var exists = existingCsvalues.Any(v => string.Equals(v.Value.ToString(), color, StringComparison.CurrentCultureIgnoreCase));
                    if (!exists)
                    {
                        //add new value if it doesn't exists                        
                        _criteriaProcessor.CreateNewValue(criteriaCode, color, imprintColorScvId, "CUST");
                    }
                });

                _criteriaProcessor.DeleteCsValues(existingCsvalues, imprintColors, criteriaSet);
            }
        }

        private static void processSoldUnimprinted(string text)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
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
                var unimprinted = Lookups.ImprintMethodsLookup.FirstOrDefault(i => string.Equals(i.CodeValue, "Unimprinted", StringComparison.CurrentCultureIgnoreCase));

                if (unimprinted != null)
                {                   
                    soldUnimprintedScvId = Convert.ToInt64(unimprinted.ID);
                }

                var criteriaCode = Constants.CriteriaCodes.ImprintMethod;
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                CriteriaSetValue unimprintedCsvalue = null;

                if (criteriaSet != null)
                {
                    unimprintedCsvalue = _criteriaProcessor.getCsValueBySetCodeValueId(soldUnimprintedScvId, criteriaSet.CriteriaSetValues);
                }

                if (soldUnimprinted == "Y")
                {
                    //create new value for unimprinted if it doesn't exists
                    if (unimprintedCsvalue == null)
                    {
                        _criteriaProcessor.CreateNewValue(criteriaCode, "Unimprinted", soldUnimprintedScvId);
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
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                var otherArtworkScValue = Lookups.ArtworkLookup.FirstOrDefault(a => string.Equals(a.CodeValue, "Other", StringComparison.CurrentCultureIgnoreCase));
                var modelValues = new List<FieldInfo>();
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
                            existingCsValue = _criteriaProcessor.getCsValueBySetCodeValueId(exists.ID.Value, existingCsvalues);                                                           
                        }
                        else
                        {                           
                            existingCsValue = _criteriaProcessor.getCsValueByAlias(exists.ID.Value, existingCsvalues, splittedValue.Alias);
                        }

                        //add new value if it doesn't exists
                        if (existingCsValue == null)
                        {
                            var value = exists.ID != null && exists.ID.Value != otherArtworkScValueId ? splittedValue.CodeValue : splittedValue.Alias;
                            if (exists.ID != null)
                                _criteriaProcessor.CreateNewValue(criteriaCode, value, exists.ID.Value, "CUST", splittedValue.Alias);
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
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var criteriaCode = Constants.CriteriaCodes.ProductionTime;
                var productionTimes = new List<FieldInfo>();
                var productionTimesTokens = text.ConvertToList();
                long customSetCodeValueId = 0;
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                var criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Unit");
                var criteriaLookUp = Lookups.ProductionTimeCriteriaLookup.FirstOrDefault(i => string.Equals(i.Code, criteriaCode, StringComparison.CurrentCultureIgnoreCase));

                if (criteriaLookUp != null)
                {
                    var group = criteriaLookUp.CodeValueGroups.FirstOrDefault(cvg => string.Equals(cvg.Description, "Other", StringComparison.CurrentCultureIgnoreCase));
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
                                       
                    //TODO: also what happens when comment is specified but time is empty - in the DB? 
                    var exists = existingCsvalues.FirstOrDefault(v => !(v.Value is string) && v.Value != null && v.Value.First.UnitValue == time);// && v.CriteriaValueDetail == comment);
                    //add new value if it doesn't exists
                    if (exists == null)
                    {
                        var value = new
                        {
                            CriteriaAttributeId = criteriaAttribute.ID,
                            UnitValue = time,
                            UnitOfMeasureCode = "BUSI"
                        };

                        _criteriaProcessor.CreateNewValue(criteriaCode, value, customSetCodeValueId, "CUST", comment);
                    }
                    else
                    {
                        exists.CriteriaValueDetail = comment;
                    }
                });

                _criteriaProcessor.DeleteCsValues(existingCsvalues, productionTimes, criteriaSet, "UnitValue");
            }
        }

        private static void processRushService(string text)
        {
            //comma separated list of values
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var criteriaCode = Constants.CriteriaCodes.RushService;
                var valueField = "Rush Service";
                //should have only one value
                var splittedValue = text.SplitValue(':');
                //should have Y/N as the first value                
                var validValues = new[] { "Y", "N" };
                if (!string.IsNullOrWhiteSpace(splittedValue.CodeValue) && !validValues.Contains(splittedValue.CodeValue))
                {
                    addValidationError("GNER", "invalid value for Process Rush Service");
                    return;
                }

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);

                if (splittedValue.CodeValue == Constants.BooleanFlag.TRUE)
                {
                    //check if the rush service value already exists
                    if (criteriaSet.CriteriaSetValues.Count() == 1)
                    {
                        var csValue = criteriaSet.CriteriaSetValues.FirstOrDefault(v => (v.Value is string) && string.Equals(v.Value.ToString(), valueField, StringComparison.CurrentCultureIgnoreCase));
                        if (csValue != null)
                        {
                            csValue.CriteriaValueDetail = splittedValue.Alias;
                        }                        
                    }
                    else
                    {
                        var criteria = Lookups.ProductionTimeCriteriaLookup.FirstOrDefault(c => c.Code == criteriaCode);

                        if (criteria != null)
                        {
                            var group = criteria.CodeValueGroups.FirstOrDefault();
                            if (group != null)
                            {
                                var setCodeValue = group.SetCodeValues.FirstOrDefault(s => string.Equals(s.CodeValue, "Other", StringComparison.CurrentCultureIgnoreCase));
                                if (setCodeValue != null)
                                {
                                    long scvId = setCodeValue.ID;
                                    _criteriaProcessor.CreateNewValue(criteriaCode, valueField, scvId, "CUST", valueField);
                                }
                            }
                        }
                    }
                }
                else
                {
                    var existingValue = criteriaSet.CriteriaSetValues.FirstOrDefault(v => string.Equals(v.Value.ToString(), valueField, StringComparison.CurrentCultureIgnoreCase));
                    if (existingValue != null)
                    {
                        criteriaSet.CriteriaSetValues.Remove(existingValue);
                    }
                }
            }           
        }      

        private static void processRushTime(string text)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var criteriaCode = Constants.CriteriaCodes.RushService;
                var rushTimes = new List<FieldInfo>();
                var rushTimesTokens = text.ConvertToList();
                long customSetCodeValueId = 0;
                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                var criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Unit");
                var criteriaLookUp = Lookups.ProductionTimeCriteriaLookup.FirstOrDefault(i => i.Code == criteriaCode);

                if (criteriaLookUp != null)
                {
                    var group = criteriaLookUp.CodeValueGroups.FirstOrDefault(cvg => string.Equals(cvg.Description, "Other", StringComparison.CurrentCultureIgnoreCase));
                    if (group != null)
                    {
                        var setCodeValue = group.SetCodeValues.FirstOrDefault();
                        if (setCodeValue != null)
                            customSetCodeValueId = setCodeValue.ID;
                    }
                }

                rushTimesTokens.ForEach(token => rushTimes.Add(token.SplitValue(':')));

                rushTimes.ForEach(rushTime =>
                {
                    var comment = string.Empty;

                    string days = rushTime.CodeValue;

                    if (!string.IsNullOrWhiteSpace(rushTime.Alias))
                    {
                        comment = rushTime.Alias;
                    }

                    if (criteriaSet.CriteriaSetValues.Count() == 1 && criteriaSet.CriteriaSetValues.First().Value is string)
                    {
                        var value = new
                        {
                            CriteriaAttributeId = criteriaAttribute.ID,
                            UnitValue = days,
                            UnitOfMeasureCode = "BUSI"
                        };                      
                        criteriaSet.CriteriaSetValues.First().Value = value;                        
                    }
                    else
                    {
                        var exists = criteriaSet.CriteriaSetValues.FirstOrDefault(csv => !(csv.Value is string) && csv.Value != null && csv.Value.First.UnitValue == days);                       
                        //add new value if it doesn't exists
                        if (exists == null)
                        {
                            var value = new
                            {
                                CriteriaAttributeId = criteriaAttribute.ID,
                                UnitValue = days,
                                UnitOfMeasureCode = "BUSI"
                            };

                            _criteriaProcessor.CreateNewValue(criteriaCode, value, customSetCodeValueId, "CUST", comment);
                        }
                        else
                        {
                            exists.CriteriaValueDetail = comment;
                        }
                    }
                });
              
                _criteriaProcessor.DeleteCsValues(criteriaSet.CriteriaSetValues, rushTimes, criteriaSet, "UnitValue");
            }
        }

        private static void processSameDay(string text)
        {
            //comma separated list of values
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var criteriaCode = "SDRU";
                var valueField = "Same Day Service";
                //should have only one value
                var splittedValue = text.SplitValue(':');
                //should have Y/N as the first value                
                var validValues = new[] { "Y", "N" };
                if (!string.IsNullOrWhiteSpace(splittedValue.CodeValue) && !validValues.Contains(splittedValue.CodeValue))
                {
                    addValidationError("GNER", "invalid value for Process Same Day Service");
                    return;
                }

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);

                if (splittedValue.CodeValue == Constants.BooleanFlag.TRUE)
                {
                    //check if the sample value already exists
                    var csValue = criteriaSet.CriteriaSetValues.FirstOrDefault(v => string.Equals(v.Value, valueField, StringComparison.CurrentCultureIgnoreCase));
                    if (csValue != null)
                    {
                        csValue.CriteriaValueDetail = splittedValue.Alias;
                    }
                    else
                    {
                        var criteria = Lookups.ProductionTimeCriteriaLookup.FirstOrDefault(c => c.Code == criteriaCode);
                      
                        if (criteria != null)
                        {
                            var group = criteria.CodeValueGroups.FirstOrDefault();
                            if (group != null)
                            {
                                var setCodeValue = group.SetCodeValues.FirstOrDefault(s => string.Equals(s.CodeValue, "Other", StringComparison.CurrentCultureIgnoreCase));
                                if (setCodeValue != null)
                                {
                                    long smplScvId = setCodeValue.ID;
                                    _criteriaProcessor.CreateNewValue(criteriaCode, valueField, smplScvId);
                                }
                            }
                        }
                    }
                }
                else
                {
                    var existingValue = criteriaSet.CriteriaSetValues.FirstOrDefault(v => string.Equals(v.Value.ToString(), valueField, StringComparison.CurrentCultureIgnoreCase));
                    if (existingValue != null)
                    {
                        criteriaSet.CriteriaSetValues.Remove(existingValue);
                    }
                }
            }            
        }

        private static void processSafetyWarnings(string text)
        {
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var safetyWarnings = text.ConvertToList();

                if (_currentProduct.SelectedSafetyWarnings == null)
                    _currentProduct.SelectedSafetyWarnings = new Collection<SafetyWarning>();

                safetyWarnings.ForEach(curSafetyWarning =>
                {
                    //need to lookup safetyWarnings
                    var safetyWarning = Lookups.SafetywarningsLookup.FirstOrDefault(c => string.Equals(c.Value, curSafetyWarning, StringComparison.CurrentCultureIgnoreCase));
                    if (safetyWarning != null)
                    {
                        var existing = _currentProduct.SelectedSafetyWarnings.FirstOrDefault(c => string.Equals(c.Description, safetyWarning.Value, StringComparison.CurrentCultureIgnoreCase));
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
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var complianceCertifications = text.ConvertToList();

                if (_currentProduct.SelectedComplianceCerts == null)
                    _currentProduct.SelectedComplianceCerts = new Collection<ProductComplianceCert>();

                complianceCertifications.ForEach(curCert =>
                {
                    //need to lookup complianceCertifications
                    var complianceCert = Lookups.ComplianceLookup.FirstOrDefault(c => string.Equals(c.Value, curCert, StringComparison.CurrentCultureIgnoreCase));
                    if (complianceCert != null)
                    {
                        var existing = _currentProduct.SelectedComplianceCerts.FirstOrDefault(c => string.Equals(c.Description, complianceCert.Value, StringComparison.CurrentCultureIgnoreCase));
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
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var linenames = text.ConvertToList();
                var existingLinenames = _currentProduct.SelectedLineNames;

                linenames.ForEach(linename =>
                {
                    var linenameFound = Lookups.LinenamesLookup.FirstOrDefault(l => string.Equals(l.Name, linename, StringComparison.CurrentCultureIgnoreCase));
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
                var lineNamesToDelete = existingLinenames.Select(e => e.Name).Except(linenames).Select(s => s).ToList();

                lineNamesToDelete.ForEach(l =>
                {
                    var toDelete = _currentProduct.SelectedLineNames.FirstOrDefault(v => string.Equals(v.Name, l, StringComparison.CurrentCultureIgnoreCase));
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
                                var newProductMediaCitation = new ProductMediaCitation
                                {
                                    ProductId = _currentProduct.ID, 
                                    Description = foundMediaCitation.Name, 
                                    MediaCitationId = foundMediaCitation.ID
                                };

                                var newProductMediaCitationReference = new ProductMediaCitationReference
                                {
                                    MediaCitationId = foundMediaCitation.ID, 
                                    MediaCitationReference = new MediaCitationReference
                                    {
                                        Number = splittedCat.Length == 3 ? splittedCat[2] : string.Empty
                                    }
                                };

                                newProductMediaCitation.ProductMediaCitationReferences = new List<ProductMediaCitationReference>
                                {
                                    newProductMediaCitationReference
                                };


                                _currentProduct.ProductMediaCitations.Add(newProductMediaCitation);
                            }
                        }
                    }
                }
            }
        }

        private static MATERIAL_FORMAT GetMaterialFormat(string text)
        {
            MATERIAL_FORMAT format = MATERIAL_FORMAT.SINGLE;
            var isCombo = text.ToLower().Contains("combo");
            var isBlend = text.ToLower().Contains("blend");

            if (isCombo && !isBlend)
            {
                format = MATERIAL_FORMAT.COMBO;
            }
            else if (!isCombo && isBlend)
            {
                format = MATERIAL_FORMAT.BLEND;
            }
            else if (isCombo && isBlend)
            {
                format = MATERIAL_FORMAT.COMBO_BLEND;
            }

            return format;
        }

        private static MajorCodeValueGroup GetLookupMaterial(string materialName, string materialAlias, IEnumerable<MajorCodeValueGroup> lookup)
        {
            var validMaterial = lookup.FirstOrDefault(m => string.Equals(m.Description, materialAlias, StringComparison.InvariantCultureIgnoreCase));

            if (validMaterial == null)
            {
                if (materialName != materialAlias)
                    validMaterial = lookup.FirstOrDefault(m => string.Equals(m.Description, materialName, StringComparison.InvariantCultureIgnoreCase));

                if (validMaterial == null) // Material is not found set group to Other
                    validMaterial = lookup.FirstOrDefault(m => string.Equals(m.Description, "Other", StringComparison.InvariantCultureIgnoreCase));
            }

            return validMaterial;
        }

        private static long GetMaterialSetCodeValueId(string name, string alias, IEnumerable<MajorCodeValueGroup> lookup)
        {
            var blendMaterial = GetLookupMaterial(name, alias, lookup);
            return blendMaterial.CodeValueGroups.FirstOrDefault().SetCodeValues.FirstOrDefault().ID;
        }

        private static void genericProcessMaterial(string material, string percentage, string materialAlias, string criteriaCode, MATERIAL_FORMAT materialFormat, bool isfirstBlendMaterial, IEnumerable<MajorCodeValueGroup> lookup)
        {
            var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
            var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
            var validMaterial = GetLookupMaterial(material, materialAlias, lookup);

            if (validMaterial != null)
            {
                var materialCSV = existingCsvalues.FirstOrDefault(csv => string.Equals(csv.Value.ToString(), materialAlias, StringComparison.CurrentCultureIgnoreCase));
                var setCodeValueId = validMaterial.CodeValueGroups.FirstOrDefault().SetCodeValues.FirstOrDefault().ID;
                //add new value if it doesn't exists
                if (materialCSV == null)
                {
                    if (materialFormat == MATERIAL_FORMAT.BLEND || materialFormat == MATERIAL_FORMAT.COMBO_2_BLEND)
                    {
                        var blendMaterialcsvId = GetMaterialSetCodeValueId("Blend", string.Empty, lookup);
                        var criteriaSetCodeValueLink = new CriteriaSetCodeValueLink
                        {
                            ChildCriteriaSetCodeValue = new CriteriaSetCodeValue
                            {
                                SetCodeValueId = setCodeValueId,
                                CodeValue = percentage.ToString()
                            }
                        };
                        _criteriaProcessor.CreateNewValue(criteriaCode, materialAlias, blendMaterialcsvId, "LIST", "", "", "", criteriaSetCodeValueLink);

                    }
                    else
                    {
                        _criteriaProcessor.CreateNewValue(criteriaCode, materialAlias, setCodeValueId, "LIST");
                    }
                }
                else
                {
                    var blendMaterialcsvId = GetMaterialSetCodeValueId("Blend", string.Empty, lookup);
                    switch (materialFormat)
                    {
                        case MATERIAL_FORMAT.SINGLE:
                            materialCSV.CriteriaSetCodeValues.FirstOrDefault().SetCodeValueId = setCodeValueId;
                            break;
                        case MATERIAL_FORMAT.COMBO:
                            var cscv = materialCSV.CriteriaSetCodeValues.FirstOrDefault(cv => cv.SetCodeValueId == setCodeValueId);
                            if (cscv == null)
                            {
                                var newCscv = new CriteriaSetCodeValue
                                {
                                    CriteriaSetValueId = materialCSV.ID,
                                    SetCodeValueId = setCodeValueId,
                                    ID = Utils.IdGenerator.getNextid(),
                                    DisplaySequence = 2
                                };

                                materialCSV.CriteriaSetCodeValues.Add(newCscv);
                            }
                            break;
                        case MATERIAL_FORMAT.BLEND:
                            var blendcv = materialCSV.CriteriaSetCodeValues.FirstOrDefault(cv => cv.SetCodeValueId == blendMaterialcsvId);
                            if (blendcv != null)
                            {
                                var materialExists = blendcv.ChildCriteriaSetCodeValues.FirstOrDefault(ccv => ccv.ChildCriteriaSetCodeValue.SetCodeValueId == setCodeValueId);
                                if (materialExists == null)
                                {
                                    var criteriaSetCodeValueLink = new CriteriaSetCodeValueLink
                                    {
                                        ChildCriteriaSetCodeValue = new CriteriaSetCodeValue
                                        {
                                            SetCodeValueId = setCodeValueId,
                                            CodeValue = percentage.ToString()
                                        }
                                    };

                                    blendcv.ChildCriteriaSetCodeValues.Add(criteriaSetCodeValueLink);
                                }
                            }
                            break;
                        case MATERIAL_FORMAT.COMBO_BLEND:
                        case MATERIAL_FORMAT.COMBO_2_BLEND:
                            var blendedCVs = materialCSV.CriteriaSetCodeValues.Where(cv => cv.SetCodeValueId == blendMaterialcsvId);
                            var blendedCV = blendedCVs.FirstOrDefault();
                            if (materialFormat == MATERIAL_FORMAT.COMBO_2_BLEND)
                            {
                                if (!isfirstBlendMaterial)
                                    blendedCV = blendedCVs.Count() == 2 ? blendedCVs.FirstOrDefault(cv => cv.DisplaySequence == 2) : null;
                            }
                            if (blendedCV != null)
                            {
                                var _materialExists = blendedCV.ChildCriteriaSetCodeValues.FirstOrDefault(ccv => ccv.ChildCriteriaSetCodeValue.SetCodeValueId == setCodeValueId);
                                if (_materialExists == null)
                                {
                                    var criteriaSetCodeValueLink = new CriteriaSetCodeValueLink
                                    {
                                        ChildCriteriaSetCodeValue = new CriteriaSetCodeValue
                                        {
                                            SetCodeValueId = setCodeValueId,
                                            CodeValue = percentage.ToString()
                                        }
                                    };
                                    blendedCV.ChildCriteriaSetCodeValues.Add(criteriaSetCodeValueLink);
                                }
                            }
                            else
                            {
                                var newCriteriaSetCodeValueLink = new CriteriaSetCodeValueLink
                                {
                                    ChildCriteriaSetCodeValue = new CriteriaSetCodeValue
                                    {
                                        SetCodeValueId = setCodeValueId,
                                        CodeValue = percentage.ToString()
                                    }
                                };

                                var newCscv = new CriteriaSetCodeValue
                                {
                                    CriteriaSetValueId = materialCSV.ID,
                                    SetCodeValueId = blendMaterialcsvId,
                                    ID = Utils.IdGenerator.getNextid(),
                                    DisplaySequence = 2
                                };
                                newCscv.ChildCriteriaSetCodeValues.Add(newCriteriaSetCodeValueLink);
                                materialCSV.CriteriaSetCodeValues.Add(newCscv);
                            }
                            break;
                    }
                }
            }
            else
            {
                //log batch error
                addValidationError(criteriaCode, material);
                _hasErrors = true;
            }
        }
        
        private static void processMaterials(string text)
        {
            //comma separated list of values
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var criteriaCode = Constants.CriteriaCodes.Material;
                var materialLookup = Lookups.MaterialLookup;
                var valueList = text.ConvertToList();
                var aliasList = new List<string>();
                valueList.ForEach(value =>
                {
                    var material_format = GetMaterialFormat(value);
                    var splittedValue = value.SplitValue('=');
                    var materials = splittedValue.CodeValue;
                    var materialAlias = splittedValue.Alias;
                    aliasList.Add(materialAlias);
                    switch (material_format)
                    { 
                        case MATERIAL_FORMAT.SINGLE:
                            genericProcessMaterial(materials, string.Empty, materialAlias, criteriaCode, material_format, true, materialLookup);
                            break;
                        case MATERIAL_FORMAT.COMBO: //Format:  Group Name1:Combo:Group Name2=Alias
                            processComboMaterial(materials, materialAlias, criteriaCode, material_format, materialLookup);                 
                            break;
                        case MATERIAL_FORMAT.BLEND: 
                            processBlendMaterial(materials, materialAlias, criteriaCode, material_format, true, materialLookup);
                            break;
                        case MATERIAL_FORMAT.COMBO_BLEND: 
                            processComboBlendMaterial(materials, materialAlias, criteriaCode, material_format, materialLookup);
                            break;
                    }                    
                });

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                _criteriaProcessor.DeleteCsValues(existingCsvalues, aliasList, criteriaSet);               
            }            
        }

        private static void processComboMaterial(string materials, string materialAlias, string criteriaCode, MATERIAL_FORMAT material_format, List<MajorCodeValueGroup> materialLookup)
        {            
            var materialSeparators = new string[] { ":Combo:", "combo" };
            var comboMaterial = materials.Split(materialSeparators, StringSplitOptions.RemoveEmptyEntries);
            if (comboMaterial.Length == 2)
            {
                var firstMaterialName = comboMaterial[0];                  
                var firstValidMaterialcvid = GetMaterialSetCodeValueId(firstMaterialName, materialAlias, materialLookup);

                var secondMaterialName = comboMaterial[1];                
                var secondValidMaterialcvid = GetMaterialSetCodeValueId(secondMaterialName, materialAlias, materialLookup);

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var materialCSV = criteriaSet.CriteriaSetValues.FirstOrDefault(csv => string.Equals(csv.Value.ToString(), materialAlias, StringComparison.CurrentCultureIgnoreCase));
                if (materialCSV != null)
                {
                    _criteriaProcessor.deleteCodeValues(materialCSV, new List<long> { firstValidMaterialcvid, secondValidMaterialcvid }, criteriaSet);
                }

                genericProcessMaterial(firstMaterialName, string.Empty, materialAlias, criteriaCode, material_format, true, materialLookup);
                genericProcessMaterial(secondMaterialName, string.Empty, materialAlias, criteriaCode, material_format, false, materialLookup);
            }
            else
            {
                addValidationError(criteriaCode, string.Join(materials, "=", materialAlias));
                _hasErrors = true;
            }           
        }

        private static void processBlendMaterial(string materials, string materialAlias, string criteriaCode, MATERIAL_FORMAT material_format, bool isfirstBlendMaterial, List<MajorCodeValueGroup> materialLookup)
        {
            //Format:  Blend:Material1:%:Material2:%=Alias
            var blendMaterials = materials.Split(':');
            var isPercentageExists = blendMaterials.Length == 5;
            if (blendMaterials.Length == 3 || blendMaterials.Length == 5)
            {
                var firstMaterialName = blendMaterials[1];
                var firstValidMaterial = GetLookupMaterial(firstMaterialName, materialAlias, materialLookup);
                var firstValidMaterialcvid = firstValidMaterial != null ? firstValidMaterial.CodeValueGroups.FirstOrDefault().SetCodeValues.FirstOrDefault().ID : -1;

                var firstMaterialPercentage = string.Empty;
                if (isPercentageExists)
                {
                    firstMaterialPercentage = blendMaterials[2];
                }

                var secondMaterialName = !isPercentageExists ? blendMaterials[2] : blendMaterials[3];
                var secondValidMaterial = GetLookupMaterial(secondMaterialName, materialAlias, materialLookup);
                var secondValidMaterialcvid = secondValidMaterial != null ? secondValidMaterial.CodeValueGroups.FirstOrDefault().SetCodeValues.FirstOrDefault().ID : -1;

                var secondMaterialPercentage = string.Empty;
                if (isPercentageExists)
                {
                    secondMaterialPercentage = blendMaterials[4];
                }

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var materialCSVs = criteriaSet.CriteriaSetValues.Where(csv => string.Equals(csv.Value.ToString(), materialAlias, StringComparison.CurrentCultureIgnoreCase));
                var blendMaterialCSV = materialCSVs.FirstOrDefault();

                if (material_format == MATERIAL_FORMAT.COMBO_2_BLEND)
                {
                    if (!isfirstBlendMaterial)
                        blendMaterialCSV = materialCSVs.Count() == 2 ? materialCSVs.FirstOrDefault(cv => cv.DisplaySequence == 2) : null;                      
                }
                    
                if (blendMaterialCSV != null)
                {
                    _criteriaProcessor.deleteChildCriteriaSetCodeValues(blendMaterialCSV.CriteriaSetCodeValues.First(), new List<long> { firstValidMaterialcvid, secondValidMaterialcvid }, blendMaterialCSV);
                }
                genericProcessMaterial(firstMaterialName, firstMaterialPercentage, materialAlias, criteriaCode, material_format, isfirstBlendMaterial, materialLookup);
                genericProcessMaterial(secondMaterialName, secondMaterialPercentage, materialAlias, criteriaCode, material_format, isfirstBlendMaterial, materialLookup);
            }
            else
            {
                addValidationError(criteriaCode, materials + "=" + materialAlias);
                _hasErrors = true;
            }            
        }

        private static void processComboBlendMaterial(string materials, string materialAlias, string criteriaCode, MATERIAL_FORMAT material_format, List<MajorCodeValueGroup> materialLookup)
        {
            // Format:  Blend:Material1:%:Material2:%:Combo:Group Name=Alias OR
            // Format:  Group Name:Combo:Blend:Material1:%:Material2:%=Alias
            // Format:  Blend:Material1:%:Material2:%:Combo:Blend:Material1:%:Material2:%=Alias
            
            var materialSeparators = new string[] { ":Combo:", ":combo:" };
            var comboMaterial = materials.Split(materialSeparators, StringSplitOptions.RemoveEmptyEntries);
            if (comboMaterial.Length == 2)
            {
                var blendMaterialcvid = GetMaterialSetCodeValueId("Blend", string.Empty, materialLookup);
                var firstMaterial = comboMaterial[0];
                var firstMaterial_format = GetMaterialFormat(firstMaterial);
                var firstMaterialcvid = firstMaterial_format == MATERIAL_FORMAT.BLEND ? blendMaterialcvid : GetMaterialSetCodeValueId(firstMaterial, materialAlias, materialLookup);

                var secondMaterial = comboMaterial[1];
                var secondMaterial_format = GetMaterialFormat(secondMaterial);
                var secondMaterialcvid = secondMaterial_format == MATERIAL_FORMAT.BLEND ? blendMaterialcvid : GetMaterialSetCodeValueId(secondMaterial, materialAlias, materialLookup);

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var materialCSV = criteriaSet.CriteriaSetValues.FirstOrDefault(csv => string.Equals(csv.Value.ToString(), materialAlias, StringComparison.CurrentCultureIgnoreCase));
                if (materialCSV != null)
                {
                    _criteriaProcessor.deleteCodeValues(materialCSV, new List<long> { firstMaterialcvid, secondMaterialcvid }, criteriaSet);
                }

                if (firstMaterial_format == MATERIAL_FORMAT.BLEND && secondMaterial_format == MATERIAL_FORMAT.BLEND)
                {
                    processBlendMaterial(firstMaterial, materialAlias, criteriaCode, MATERIAL_FORMAT.COMBO_2_BLEND, true, materialLookup);
                }
                else if (firstMaterial_format == MATERIAL_FORMAT.BLEND)
                {
                    processBlendMaterial(firstMaterial, materialAlias, criteriaCode, MATERIAL_FORMAT.BLEND, true, materialLookup);
                }
                else
                {
                    genericProcessMaterial(firstMaterial, string.Empty, materialAlias, criteriaCode, MATERIAL_FORMAT.SINGLE, true, materialLookup);
                }

                if (firstMaterial_format == MATERIAL_FORMAT.BLEND && secondMaterial_format == MATERIAL_FORMAT.BLEND)
                {
                    processBlendMaterial(secondMaterial, materialAlias, criteriaCode, MATERIAL_FORMAT.COMBO_2_BLEND, false, materialLookup);
                }
                else if (secondMaterial_format == MATERIAL_FORMAT.BLEND)
                {
                    processBlendMaterial(secondMaterial, materialAlias, criteriaCode, material_format, false, materialLookup);
                }
                else
                {
                    genericProcessMaterial(secondMaterial, string.Empty, materialAlias, criteriaCode, MATERIAL_FORMAT.COMBO, false, materialLookup);
                }
                   
            }
            else
            {
                addValidationError(criteriaCode, string.Join(materials, "=", materialAlias));
                _hasErrors = true;
            }            
        }

        private static COLOR_FORMAT GetColorFormat(string text)
        {
            COLOR_FORMAT format = COLOR_FORMAT.SINGLE;
            var isCombo = text.ToLower().Contains("combo");
            var tokens = text.Split(':');

            if (isCombo && tokens.Length == 4)
            {
                format = COLOR_FORMAT.COMBO_1PRIMARY_1SECONDARY;
            }
            else if (isCombo && tokens.Length == 6)
            {
                format = COLOR_FORMAT.COMBO_1PRIMARY_2SECONDARY;
            }           

            return format;
        }

        private static ColorGroup GetLookupColor(string colorName, string colorAlias, IEnumerable<ColorGroup> lookup)
        {
            var validColor = lookup.FirstOrDefault(m => string.Equals(m.Description, colorAlias, StringComparison.InvariantCultureIgnoreCase));

            if (validColor == null)
            {
                if (colorName != colorAlias)
                    validColor = lookup.FirstOrDefault(m => string.Equals(m.Description, colorName, StringComparison.InvariantCultureIgnoreCase));

                if (validColor == null) // Color is not found set group to Other
                    validColor = lookup.FirstOrDefault(m => string.Equals(m.Description, "Unclassified/Other", StringComparison.InvariantCultureIgnoreCase));
            }

            return validColor;
        }

        private static long GetColorSetCodeValueId(string name, string alias, IEnumerable<ColorGroup> lookup)
        {
            var color = GetLookupColor(name, alias, lookup);
            return color.SetCodeValues.First().Id;
        }

        private static void processComboColor(string colors, string colorAlias, string criteriaCode, COLOR_FORMAT color_format, IEnumerable<ColorGroup> colorLookup)
        {
            // Format:  SubGroup Name:Combo:SubGroup Name:Type=Alias
            // Format:  SubGroup Name:Combo:SubGroup Name:Type:SubGroup Name:Type=Alias

            var colorSeparators = new string[] { ":Combo:", ":combo:" };
            var comboColor = colors.Split(colorSeparators, StringSplitOptions.RemoveEmptyEntries);
            if (comboColor.Length == 2)
            {
                
                var firstColor = comboColor[0];              
                var firstColorcvid = GetColorSetCodeValueId(firstColor, colorAlias, colorLookup);

                var secondColor = string.Empty;
                var secondColorType = string.Empty;
                var secondColorcvid = 0L;

                var thirdColor = string.Empty;
                var thirdColorType = string.Empty;
                var thirdColorcvid = 0L;

                var combo2ndPart = comboColor[1].Split(':');
                if (combo2ndPart.Length >= 2)
                {
                    secondColor = combo2ndPart[0];
                    secondColorType = combo2ndPart[1];
                    secondColorcvid = GetColorSetCodeValueId(firstColor, colorAlias, colorLookup);
                }

                if (color_format == COLOR_FORMAT.COMBO_1PRIMARY_2SECONDARY)
                {
                    thirdColor = combo2ndPart[2];
                    thirdColorType = combo2ndPart[3];
                    thirdColorcvid = GetColorSetCodeValueId(thirdColor, colorAlias, colorLookup);
                }

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var colorCSV = criteriaSet.CriteriaSetValues.FirstOrDefault(csv => string.Equals(csv.Value.ToString(), colorAlias, StringComparison.CurrentCultureIgnoreCase));
                if (colorCSV != null)
                {
                    _criteriaProcessor.deleteCodeValues(colorCSV, new List<long> { firstColorcvid, secondColorcvid, thirdColorcvid }, criteriaSet);
                }

                genericProcessColor(firstColor, "main", colorAlias, criteriaCode, color_format, true, colorLookup);
                genericProcessColor(secondColor, secondColorType, colorAlias, criteriaCode, color_format, false, colorLookup);

                if (color_format == COLOR_FORMAT.COMBO_1PRIMARY_2SECONDARY)
                {
                    genericProcessColor(thirdColor, thirdColorType, colorAlias, criteriaCode, color_format, false, colorLookup);
                }
            }
            else
            {
                addValidationError(criteriaCode, string.Join(colors, "=", colorAlias));
                _hasErrors = true;
            }   
        }     

        private static void genericProcessColor(string color, string colorType, string colorAlias, string criteriaCode, COLOR_FORMAT colorFormat, bool isfirstColor, IEnumerable<ColorGroup> lookup)
        {
            var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
            var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
            var validColor = GetLookupColor(color, colorAlias, lookup);

            if (validColor != null)
            {
                var colorCSV = existingCsvalues.FirstOrDefault(csv => string.Equals(csv.Value.ToString(), colorAlias, StringComparison.CurrentCultureIgnoreCase));
                var setCodeValueId = validColor.SetCodeValues.First().Id;
                //add new value if it doesn't exists
                if (colorCSV == null)
                {
                    _criteriaProcessor.CreateNewValue(criteriaCode, colorAlias, setCodeValueId, "LIST", "", "", colorType);                   
                }
                else
                {                   
                    switch (colorFormat)
                    {
                        case COLOR_FORMAT.SINGLE:
                            colorCSV.CriteriaSetCodeValues.FirstOrDefault().SetCodeValueId = setCodeValueId;
                            break;
                        case COLOR_FORMAT.COMBO_1PRIMARY_1SECONDARY:
                        case COLOR_FORMAT.COMBO_1PRIMARY_2SECONDARY:
                            var cscv = colorCSV.CriteriaSetCodeValues.FirstOrDefault(cv => cv.SetCodeValueId == setCodeValueId);
                            if (cscv == null)
                            {
                                var newCscv = new CriteriaSetCodeValue
                                {
                                    CriteriaSetValueId = colorCSV.ID,
                                    SetCodeValueId = setCodeValueId,
                                    ID = Utils.IdGenerator.getNextid(),
                                    CodeValueDetail = colorType,                                
                                };

                                colorCSV.CriteriaSetCodeValues.Add(newCscv);
                            }                                      
                            break;               
                    }
                }
            }
            else
            {
                //log batch error
                addValidationError(criteriaCode, color);
                _hasErrors = true;
            }
        }

        private static void processProductColors(string text)
        {
            //comma separated list of values
            if (_firstRowForProduct && !string.IsNullOrWhiteSpace(text))
            {
                var criteriaCode = Constants.CriteriaCodes.Color;
                var colorLookup = Lookups.ColorGroupList.SelectMany(g => g.CodeValueGroups);
                var valueList = text.ConvertToList();
                var aliasList = new List<string>();
                valueList.ForEach(value =>
                {
                    var color_format = GetColorFormat(value);
                    var splittedValue = value.SplitValue('=');
                    var colors = splittedValue.CodeValue;
                    var colorAlias = splittedValue.Alias;
                    aliasList.Add(colorAlias);
                    switch (color_format)
                    {
                        case COLOR_FORMAT.SINGLE:
                            genericProcessColor(colors, "main", colorAlias, criteriaCode, color_format, true, colorLookup);
                            break;
                        case COLOR_FORMAT.COMBO_1PRIMARY_1SECONDARY:
                        case COLOR_FORMAT.COMBO_1PRIMARY_2SECONDARY:
                            processComboColor(colors, colorAlias, criteriaCode, color_format, colorLookup);
                            break;
                    }
                });

                var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode);
                var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();
                _criteriaProcessor.DeleteCsValues(existingCsvalues, aliasList, criteriaSet);
            }           
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

            var mediaCitation = Lookups.MediaCitations.FirstOrDefault(m => string.Equals(m.Name, name, StringComparison.CurrentCultureIgnoreCase) && m.Year == year);
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

        private static void finishProduct()
        {
            //if we've started a radar model, 
            // we "send" the product to Radar for processing. 
            if (_currentProduct != null && !_hasErrors)
            {
                _priceProcessor.FinalizeProductPricing(_currentProduct);
                //TODO: other repeatable sets will "finalize" here as well. 

                //var x = _currentProduct;
                if (!_publishCurrentProduct)
                {
                    //add "no pub" attribute to radar POST
                }
                _log.DebugFormat("completed work with product {0}", _curXid);
                _currentProduct = null;
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
                    select existingLookup ?? new GenericLookUp { CodeValue = value, ID = null })
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
            var version = (p2 == "V2" ? "2.0.0" : p1 == "ASIS" ? "0.0.1" : "0.0.5");

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
            else
            {
                if (c.InlineString != null)
                {
                    text = c.InlineString.InnerText;
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
