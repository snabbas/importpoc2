using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using ASI.Sugar.Collections;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using Radar.Core.Models.Batch;
using Radar.Core.Models.Import;
using Radar.Core.Models.Product;
using Radar.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;

namespace ImportPOC2
{
    class Program
    {
        private static SharedStringTablePart _stringTable;
        private static List<string> _sheetColumnsList;
        private static IQueryable<Template> _mapping;
        private static string _curXid;
        private static int _companyId;
        private static Batch _curBatch;
        private static UowPRODTask _prodTask;
        private static readonly HttpClient RadarHttpClient = new HttpClient {BaseAddress = new Uri("http://local-espupdates.asicentral.com/api/api/")};
        private static Product _currentProduct;
        private static bool _firstRowForProduct = true;

        private static List<Category> _catlist = null;
        private static List<Category> categoryList
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

        private static log4net.ILog _log;
        static void Main(string[] args)
        {
            //onetime stuff
            RadarHttpClient.DefaultRequestHeaders.Accept.Clear();
            RadarHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            //get directory from config
            //TODO: change this to read from config

            //NOTES: 
            //file name is batch ID - use to retreive batch to assign details. 
            // also use to determine company ID of sheet. 

            var curDir = Directory.GetCurrentDirectory();
            log4net.Config.XmlConfigurator.ConfigureAndWatch(new FileInfo(curDir));

            _log = log4net.LogManager.GetLogger("ImportPOC");

            _log.DebugFormat("running in {0}", curDir);

            var xlsxFiles = Directory.GetFiles(curDir, "*.xlsx");

            if (!xlsxFiles.Any())
            {
                logit(string.Format("No xlsx files found in {0}", curDir));
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
            logit(string.Format("processing {0}", Path.GetFileName(xlsxFile)));
            //processor: 

            //initizations
            //get batch: 
            var batchId = getBatchId(xlsxFile);
            _curBatch = getBatchById(batchId);

            if (_curBatch == null)
            {
                logit(string.Format("unable to find batch {0}", batchId));
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
                    //it's a column on the current product, process it. 
                    processColumn(curColIndex, text);
                }

                //colIndex++;
                //no longer incrementing this as getColIndex function determines where we are in the sheet
                //some columns are apparently skipped if they don't have data. 
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
                _currentProduct = getProductByXid() ?? new Product { CompanyId = _companyId};
                _firstRowForProduct = true; 
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
                    assignProductName(text);
                    break;
                case "Product_Number":
                    assignProductNumber(text);
                    break;

                case "Description":
                    assignDescription(text);
                    break;

                case "Summary":
                    assignSummary(text);
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
                    assignInventoryLink(text);
                    break;

                case "Product_Color":
                    break;

                case "Material":
                    break;

                case "Size_Group":
                    break;

                case "Size_Values":
                    break;

            }
        }

        private static void processKeywords(string text)
        {
            //comma delimited list of keywords - only "visible" keywords, never ad or seo keywords
            if (_firstRowForProduct)
            {
                var keywords = text.Split(',').Select(str => str.Trim()).ToList();

                keywords.Where(k => !string.IsNullOrWhiteSpace(k)).ForEach(keyword =>
                {
                    var keyObj = _currentProduct.ProductKeywords.FirstOrDefault(p => p.Value == keyword);

                });
            }

        }

        private static void processCategory(string text)
        {
            if (_firstRowForProduct)
            {
                var categories = text.Split(',').Select(str => str.Trim()).ToList();
                
                categories.ForEach(curCat =>
                {
                    if (!string.IsNullOrWhiteSpace(curCat))
                    {
                        //need to lookup categories
                        var category = categoryList.FirstOrDefault(c => c.Name == curCat);
                        if (category != null)
                        {
                            //just in case it's totally empty/null
                            if (_currentProduct.SelectedProductCategories == null)
                                _currentProduct.SelectedProductCategories = new Collection<ProductXCategory>();

                            var existing = _currentProduct.SelectedProductCategories.FirstOrDefault(c => c.CategoryCode == category.Code);
                            if (existing == null)
                            {
                                var newCat = new ProductXCategory {CategoryCode = category.Code, AdCategoryFlg = false};
                                _currentProduct.SelectedProductCategories.Add(newCat);
                            }
                            else
                            {
                                //s/b nothing to do? 
                            }
                        }
                    }
                });
            }
        }

        private static void processImage(string text)
        {
            if (_firstRowForProduct)
            {
                //text here should be a list of comma sepearated URLs, in order of display
                var urls = text.Split(',').Select(str => str.Trim()).ToList();
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

        private static void assignInventoryLink(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.ProductInventoryLink = updateField(text);
        }

        private static void assignSummary(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Summary = updateField(text);
        }

        private static void assignDescription(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Description = updateField(text);
        }

        private static void assignProductNumber(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.AsiProdNo = updateField(text);
        }

        private static void assignProductName(string text)
        {
            if (_firstRowForProduct)
                _currentProduct.Name = updateField(text);
        }

        /// <summary>
        /// Returns updated value of string field, based upon following rules:
        /// 1) if text is empty, no update occurs
        /// 2) if text is literial "NULL", field is emptied of its value
        /// 3) otherwise, new value is returned.
        /// </summary>
        /// <param name="newValue"></param>
        /// <returns>string</returns>
        private static string updateField(string newValue)
        {
            string retVal = newValue;

            if (!string.IsNullOrWhiteSpace(newValue))
            {
                retVal = (newValue == "NULL" ? string.Empty : newValue);
            }
            return retVal;
        }

        private static void finishProduct()
        {
            //if we've started a radar model, 
            // we "send" the product to Radar for processing. 
            var x = _currentProduct;
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

            retVal = compareColumns(columnMetaDataPriceV1);
            if (!retVal)
            {
                logit("sheet is not price v1 format");
                retVal = compareColumns(columnMetaDataPriceV2);
            }
            else
            {
                logit("sheet MATCH price v1 format");
            }
            if (!retVal)
            {
                logit("sheet is not price v2 format");
                retVal = compareColumns(columnMetaDataFullV1);
            }
            else
            {
                logit("sheet MATCH price v2 format");
            }
            if (!retVal)
            {
                logit("sheet is not full v1 format");
                retVal = compareColumns(columnMetaDataFullV2);
            }
            else
            {
                logit("sheet MATCH full v1 format");
            }

            logit(!retVal ? "sheet is not full v2 format" : "sheet MATCH full v2 format");


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
            //todo: hook up log4net here 
            Console.WriteLine(message);
        }
    }
}
