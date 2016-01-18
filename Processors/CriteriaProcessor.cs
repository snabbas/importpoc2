using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ImportPOC2.Utils;
using Radar.Models.Product;
using Constants = Radar.Core.Common.Constants;
using CriteriaSetCodeValue = Radar.Models.Criteria.CriteriaSetCodeValue;
using CriteriaSetValue = Radar.Models.Criteria.CriteriaSetValue;
using Radar.Models.Criteria;

namespace ImportPOC2.Processors
{
    /// <summary>
    /// set of helper method(s) to process the criteria from a sheet
    /// and reconcile with product model configurations
    /// it is expected that the product model configuration is complete at the time these methods are invoked.
    /// </summary>
    public class CriteriaProcessor
    {
        private Product _currentProduct;

        public CriteriaProcessor(Product currentProduct)
        {
            _currentProduct = currentProduct;
        }

        public ProductCriteriaSet GetCriteriaSetByCode(string criteriaCode, string optionName = "")
        {
            ProductCriteriaSet retVal = null;
            if (!string.IsNullOrWhiteSpace(criteriaCode))
            {
                var prodConfig = getDefaultProdConfig();

                if (prodConfig != null)
                {
                    var cSets = prodConfig.ProductCriteriaSets.Where(c => c.CriteriaCode == criteriaCode).ToList();
                    retVal = !string.IsNullOrWhiteSpace(optionName) ? 
                        cSets.FirstOrDefault(c => string.Equals(c.CriteriaDetail, optionName, StringComparison.CurrentCultureIgnoreCase)) : 
                        cSets.FirstOrDefault();
                }

                retVal = retVal ?? CreateNewCriteriaSet(criteriaCode, optionName);
            }
            return retVal;
        }

        public ProductCriteriaSet CreateNewCriteriaSet(string criteriaCode, string optionName = "")
        {
            var newCs = new ProductCriteriaSet
            {
                CriteriaCode = criteriaCode,
                CriteriaSetId = IdGenerator.getNextid(),
                ProductId = _currentProduct.ID
            };

            if (!string.IsNullOrWhiteSpace(optionName))
            {
                newCs.CriteriaDetail = optionName;
            }

            var productConfiguration = getDefaultProdConfig();
            if (productConfiguration != null)
                productConfiguration.ProductCriteriaSets.Add(newCs);

            return newCs;

        }
        public IEnumerable<CriteriaSetValue> GetCSValuesByCriteriaCode (string criteriaCode, string optionName = "")
        {
            var cSet = GetCriteriaSetByCode(criteriaCode, optionName);
            var result = cSet.CriteriaSetValues.ToList();

            return result;
        }


        //TODO: can we "detect" what value type code to use instead of passing it in?         
        public void CreateNewValue(ProductCriteriaSet cSet, object value, long setCodeValueId, string valueTypeCode = "LOOK", string valueDetail = "", string optionName = "", string codeValueDetail="", CriteriaSetCodeValueLink childCriteriaSetCodeValue = null, bool setFormatValue = true)
        {            
            //create new criteria set value
            var newCsv = new CriteriaSetValue
            {
                CriteriaCode = cSet.CriteriaCode,
                CriteriaSetId = cSet.CriteriaSetId,
                Value = value,
                ID = IdGenerator.getNextid(),
                ValueTypeCode = valueTypeCode,
                CriteriaValueDetail = valueDetail,
                //for most cases this would be the same as value field except in cases where value is an object and not a string
                //for those cases Radar will set this field
                FormatValue = setFormatValue ? value.ToString() : "" 
            };

            //create new criteria set code value
            var newCscv = new CriteriaSetCodeValue
            {
                CriteriaSetValueId = newCsv.ID,
                SetCodeValueId = setCodeValueId,               
                ID = IdGenerator.getNextid(),
                CodeValueDetail = codeValueDetail,
                DisplaySequence = 1,
            };

            if (childCriteriaSetCodeValue != null)
            {
                newCscv.ChildCriteriaSetCodeValues.Add(childCriteriaSetCodeValue);
            }

            newCsv.CriteriaSetCodeValues.Add(newCscv);
            cSet.CriteriaSetValues.Add(newCsv);
        }

        //TODO: pretty sure this can be done without passing in criteriaset parameter
        public void DeleteCsValues(IEnumerable<CriteriaSetValue> entities, IEnumerable<string> models, ProductCriteriaSet criteriaSet)
        {
            //delete values that are missing from the list in the file
            var valuesToDelete = entities.Select(e => e.FormatValue).Except(models).Select(s => s).ToList();
            valuesToDelete.ForEach(e =>
            {
                var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.FormatValue == e);
                criteriaSet.CriteriaSetValues.Remove(toDelete);
            });
        }

        public void DeleteCsValues(IEnumerable<CriteriaSetValue> entities, IEnumerable<FieldInfo> models, ProductCriteriaSet criteriaSet, string fieldName)
        {
            //delete values that are missing from the list in the file
            var csValuesToDelete = new List<CriteriaSetValue>();
            entities.ToList().ForEach(e =>
            {
                {
                    var exists = false;
                    switch (fieldName)
                    {
                        case "UnitValue":
                            if (e.Value is string)
                            {
                                exists = models.Any(m => string.Equals(m.CodeValue, e.Value, StringComparison.CurrentCultureIgnoreCase));
                            }
                            else if (e.Value is IList)
                            {
                                exists = models.Any(m => string.Equals(m.CodeValue, e.Value.First.UnitValue.ToString(), StringComparison.CurrentCultureIgnoreCase) && m.Alias == e.CriteriaValueDetail);
                            }
                            else
                            {
                                exists = models.Any(m => string.Equals(m.CodeValue, e.Value.UnitValue.ToString(), StringComparison.CurrentCultureIgnoreCase) && m.Alias == e.CriteriaValueDetail);
                            }
                            break;                       
                        default:
                            exists = models.Any(m => m.Alias == e.CriteriaValueDetail);
                            break;
                    }

                    if (!exists)
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


        /// <summary>
        /// given a configuration string in import sheet format (i.e., "PRCL:Red,Blue") return CSV objects that match. 
        /// </summary>
        /// <param name="criteriaDefinition"></param>
        /// <returns></returns>
        public List<CriteriaSetValue> GetValuesBySheetSpecification(string criteriaDefinition)
        {
            var retVal = new List<CriteriaSetValue>();

            if (!string.IsNullOrWhiteSpace(criteriaDefinition))
            {
                //get criteria code
                var code = parseCodeFromPriceCriteria(criteriaDefinition);
                // get list of values
                var val = parseValuesFromPriceCriteria(criteriaDefinition);
                // match 'em up

                var allcsvs = GetCSValuesByCriteriaCode(code);
                var valueDefinitions = val.Split(',').Select(v => v.Trim()).ToList();
                valueDefinitions.ForEach(d =>
                {
                    var tmp = allcsvs.FirstOrDefault(v => string.Equals(v.FormatValue.Trim(), d, StringComparison.CurrentCultureIgnoreCase));
                    if (tmp != null)
                    {
                        retVal.Add(tmp);
                    }
                });

            }
            return retVal;
        }

        private string parseCodeFromPriceCriteria(string criteriaDefinition)
        {
            var retVal = string.Empty;
            if (!string.IsNullOrWhiteSpace(criteriaDefinition))
            {
                var tmp = criteriaDefinition.Split(':');
                if (tmp.Length > 1)
                {
                    if (tmp[0].Trim().Length == 4)
                    {
                        //it "appears" to be a valid code, send it back;
                        retVal = tmp[0].Trim();
                    }
                }
            }

            return retVal;
        }

        private string parseValuesFromPriceCriteria(string criteriaDefinition)
        {
            var retVal = string.Empty;
            if (!string.IsNullOrWhiteSpace(criteriaDefinition))
            {
                //note: cannot use Split here as values could have embedded colons
                var separatorPos = criteriaDefinition.IndexOf(':');
                var tmp = criteriaDefinition.Substring(separatorPos + 1);

                retVal = tmp.Trim();
            }

            return retVal;
        }

        public void deleteCodeValues(CriteriaSetValue entity, IEnumerable<long> models, ProductCriteriaSet criteriaSet)
        {
            if (entity != null)
            {
                var exists = false;
                var cscvToDelete = new List<CriteriaSetCodeValue>();

                entity.CriteriaSetCodeValues.ToList().ForEach(e =>
                {
                    exists = models.Any(m => m == e.SetCodeValueId);

                    if (!exists)
                    {
                        cscvToDelete.Add(e);
                    }
                });

                cscvToDelete.ForEach(e =>
                {
                    var csv = criteriaSet.CriteriaSetValues.FirstOrDefault();
                    if (csv != null)
                    {
                        var toDelete = csv.CriteriaSetCodeValues.FirstOrDefault(cv => cv.SetCodeValueId == e.SetCodeValueId);
                        if (toDelete != null)
                        {
                            criteriaSet.CriteriaSetValues.FirstOrDefault().CriteriaSetCodeValues.Remove(toDelete);
                        }
                    }
                });
            }
        }

        public void deleteChildCriteriaSetCodeValues(CriteriaSetCodeValue entity, IEnumerable<long> models, CriteriaSetValue criteriaSetValue)
        {
            if (entity != null)
            {
                var exists = false;
                var cscvToDelete = new List<CriteriaSetCodeValueLink>();

                entity.ChildCriteriaSetCodeValues.ToList().ForEach(e =>
                {
                    exists = models.Any(m => m == e.ChildCriteriaSetCodeValue.SetCodeValueId);

                    if (!exists)
                    {
                        cscvToDelete.Add(e);
                    }
                });

                cscvToDelete.ForEach(e =>
                {
                    var cv = criteriaSetValue.CriteriaSetCodeValues.FirstOrDefault();
                    if(cv != null)
                    {
                        var toDelete = cv.ChildCriteriaSetCodeValues.FirstOrDefault(scv => scv == e);
                        if (toDelete != null)
                        {
                            criteriaSetValue.CriteriaSetCodeValues.FirstOrDefault().ChildCriteriaSetCodeValues.Remove(toDelete);
                        }
                    }
                });
            }
        }

        public CriteriaSetValue getCsValueBySetCodeValueId(long scvId, IEnumerable<CriteriaSetValue> criteriaSetValues)
        {
            return (from v in criteriaSetValues
                    let scv = v.CriteriaSetCodeValues.FirstOrDefault(s => s.SetCodeValueId == scvId)
                    where scv != null
                    select v)
                    .FirstOrDefault();
        }

        public CriteriaSetValue getCsValueByAlias(long scvId, IEnumerable<CriteriaSetValue> criteriaSetValues, string alias)
        {
            return (from v in criteriaSetValues
                    let scv = v.CriteriaSetCodeValues.FirstOrDefault(s => s.SetCodeValueId == scvId)
                    where scv != null
                        && v.Value == alias
                    select v)
                    .FirstOrDefault();
        }

        public CriteriaSetValue getCsValueByFormatValue(long scvId, IEnumerable<CriteriaSetValue> criteriaSetValues, string formatValue)
        {
            return (from v in criteriaSetValues
                    let scv = v.CriteriaSetCodeValues.Where(s => s.SetCodeValueId == scvId)
                    where scv != null
                        && v.FormatValue == formatValue
                    select v)
                    .FirstOrDefault();
        }

        public ProductCriteriaSet getSizeCriteriaSetByCode(string criteriaCode)
        {
            ProductCriteriaSet retVal = null;
            var prodConfig = getDefaultProdConfig();

            if (prodConfig != null)
            {
                var cSet = prodConfig.ProductCriteriaSets.FirstOrDefault(c => c.CriteriaCode == criteriaCode);
                if (cSet == null)
                {
                    cSet = prodConfig.ProductCriteriaSets.FirstOrDefault(c => Constants.CriteriaCodes.SIZE.Contains(c.CriteriaCode));
                    //create a new size criteria set if none already exists
                    if (cSet == null)
                    {
                        retVal = CreateNewCriteriaSet(criteriaCode);
                    }
                    else
                    {
                        //if another size criteria set already exists replace it with the new size criteria set
                        removeCriteriaSet(cSet.CriteriaCode);
                        retVal = CreateNewCriteriaSet(criteriaCode);
                    }
                }
                else
                {
                    retVal = cSet;
                }
            }

            return retVal;
        }

        private ProductConfiguration getDefaultProdConfig()
        {
            var defaultProdConfig = _currentProduct.ProductConfigurations.FirstOrDefault(c => c.IsDefault) ?? 
                new ProductConfiguration
            {
                IsDefault = true, 
                ProductId = _currentProduct.ID,
                ProductCriteriaSets = new Collection<ProductCriteriaSet>()
            };

            return defaultProdConfig;
        }

        public void removeCriteriaSet(string criteriaCode)
        {
            var productConfiguration = getDefaultProdConfig();
            if (productConfiguration != null)
            {
                var cs = productConfiguration.ProductCriteriaSets.FirstOrDefault(c => c.CriteriaCode == criteriaCode);
                if (cs != null)
                    productConfiguration.ProductCriteriaSets.Remove(cs);
            }
        }

        public dynamic createSizeValueObject(string criteriaCode, string value)
        {   
            var retVal = new List<dynamic>();
            CriteriaAttribute criteriaAttribute = null;
            dynamic newVal = null;
            var format = string.Empty;

            switch (criteriaCode)
            {
                case "SAHU":
                case "SABR":
                    criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode);
                    if (criteriaAttribute != null)
                    {
                        newVal = createNewValueObject(criteriaAttribute.ID, value);
                        retVal.Add(newVal);
                    }
                    break;

                case "SAWI":               
                    var separator = 'x';                    
                    var splittedValue = value.Split(separator);

                    if (splittedValue.Length > 2)
                    {
                        BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                       
                        break;
                    }

                    criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Waist");
                    if (criteriaAttribute != null)
                    {
                        newVal = createNewValueObject(criteriaAttribute.ID, splittedValue[0]);
                        retVal.Add(newVal);
                    }

                    if (splittedValue.Length == 2)
                    {
                        criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Inseam");
                        if (criteriaAttribute != null)
                        {
                            newVal = createNewValueObject(criteriaAttribute.ID, splittedValue[1]);
                            retVal.Add(newVal);
                        }
                    }                   
                    break;

                case "SANS":
                    separator = '(';
                    splittedValue = value.Split(separator);

                    if (splittedValue.Length > 2)
                    {
                        BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                        break;
                    }

                    criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Neck");
                    if (criteriaAttribute != null)
                    {
                        newVal = createNewValueObject(criteriaAttribute.ID, splittedValue[0]);
                        retVal.Add(newVal);
                    }

                    if (splittedValue.Length == 2)
                    {
                        criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Sleeve");
                        if (criteriaAttribute != null)
                        {
                            newVal = createNewValueObject(criteriaAttribute.ID, splittedValue[1]);
                            retVal.Add(newVal);
                        }
                    }
                    break;   
                 
                case "SAIT":
                    criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode);
                    //check whether the value is an integer
                    //if they provided only an integer value then the unit is defaulted to months
                    int res = 0;
                    int.TryParse(value, out res);                    
                    if (res > 0)
                    {                            
                        value = value + " months";   
                    }
                                                                
                    if (value.Contains("months"))
                    {
                        splittedValue = value.Split(' ');

                        if (splittedValue.Length == 2 && splittedValue[1] == "months")
                        {
                            value = splittedValue[0];
                            newVal = createValueOnBasisOfUnitOfMeasure("months", criteriaAttribute, value);                                
                        }
                        else
                        {
                            BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                            break;
                        }
                    }
                    else if (value.Contains("T"))
                    {
                        splittedValue = value.Split('T');

                        if (splittedValue.Length == 2 && splittedValue[1] == string.Empty)
                        {
                            value = splittedValue[0];
                            newVal = createValueOnBasisOfUnitOfMeasure("T", criteriaAttribute, value);
                        }
                        else
                        {
                            BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                            break;
                        }
                    }
                    else
                    {
                        BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                        break;
                    }                  

                    if (newVal != null)
                        retVal.Add(newVal);

                    break;
               
                case "SVWT":
                case "CAPS":
                    separator = ':';
                    splittedValue = value.Split(separator);

                    if (splittedValue.Length != 2)
                    {
                        BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                        break;
                    }

                    if (criteriaCode == "SVWT")
                    {
                        //first look it up into unit of measures for volume
                        criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Volume");
                        if (criteriaAttribute != null)
                        {
                            var foundUom = criteriaAttribute.UnitsOfMeasure.FirstOrDefault(u => u.Format == splittedValue[1]);
                            if (foundUom != null)
                            {
                                newVal = createNewValueObject(criteriaAttribute.ID, splittedValue[0], foundUom.Code);
                            }
                            //then look it up into unit of measures for weight
                            else
                            {
                                criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, "Weight");
                                if (criteriaAttribute != null)
                                {
                                    foundUom = criteriaAttribute.UnitsOfMeasure.FirstOrDefault(u => u.Format == splittedValue[1]);
                                    if (foundUom != null)
                                    {
                                        newVal = createNewValueObject(criteriaAttribute.ID, splittedValue[0], foundUom.Code);
                                    }
                                    else
                                    {
                                        BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode);
                        if (criteriaAttribute != null)
                        {
                            var foundUom = criteriaAttribute.UnitsOfMeasure.FirstOrDefault(u => u.Format == splittedValue[1]);
                            if (foundUom != null)
                            {
                                newVal = createNewValueObject(criteriaAttribute.ID, splittedValue[0], foundUom.Code);
                            }
                        }
                    }

                    if (newVal != null)
                        retVal.Add(newVal);

                    break;

                case "DIMS":
                    separator = ';';
                    splittedValue = value.Split(separator);

                    if (splittedValue.Length > 3)
                    {
                        BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                        break;
                    }

                    foreach (var dim in splittedValue)
                    {
                        var dimSplit = dim.Split(':');
                        if (dimSplit.Length != 3)
                        {
                            BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                            break;
                        }

                        var attribute = dimSplit[0].ToString();
                        var unitValue = dimSplit[1].ToString();
                        var uom = dimSplit[2].ToString();

                        //workaround for diamter as "Dia" doesn't matches the description in lookup
                        if (attribute == "Dia")
                            attribute = "Diameter";

                        criteriaAttribute = Lookups.CriteriaAttributeLookup(criteriaCode, attribute);

                        if (criteriaAttribute != null)
                        {
                            //work around for feet and inch because they don't give us the format in the sheet                            
                            switch (uom)
                            {
                                case "ft":
                                    format = "'";
                                    break;
                                case "in":
                                    format = "\"";
                                    break;
                                default:
                                    format = uom;
                                    break;
                            }

                            var foundUom = criteriaAttribute.UnitsOfMeasure.FirstOrDefault(u => u.Format == format);
                            if (foundUom != null)
                            {
                                newVal = createNewValueObject(criteriaAttribute.ID, unitValue, foundUom.Code);
                            }
                            else
                            {
                                BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                                break;
                            }
                        }
                        else
                        {
                            BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid size value: {0}", value));                           
                            break;
                        }

                        if (newVal != null)
                            retVal.Add(newVal);
                    }                   

                    break;
            }

            return retVal;
        }

        private dynamic createNewValueObject(int criteriaAttributeId, string unitValue, string unitOfMeasureCode = "")
        {
            var newValue = new
            {
                CriteriaAttributeId = criteriaAttributeId,
                UnitValue = unitValue,
                UnitOfMeasureCode = unitOfMeasureCode
            };

            return newValue;
        }

        private dynamic createValueOnBasisOfUnitOfMeasure(string unitOfMeasureType, CriteriaAttribute criteriaAttribute, string value)
        {
            dynamic retVal = null;

            if (criteriaAttribute != null)
            {
                var uomCode = criteriaAttribute.UnitsOfMeasure.First(m => m.Format == unitOfMeasureType);
                if (uomCode != null)
                {
                    retVal = createNewValueObject(criteriaAttribute.ID, value, uomCode.Code);                    
                }
            }

            return retVal;
        }

        public string formatSizeDimensionValue(string value)
        {
            var retVal = string.Empty;

            var dimSplit = value.Split(';');
            foreach (var str in dimSplit)
            {
                var split = str.Split(':');

                if (split.Length != 3)
                    break;

                //workaround for diamter as "Dia" doesn't matches the description in lookup
                if (split[0] == "Dia")
                    split[0] = "Diameter";

                var criteriaAttribute = Lookups.CriteriaAttributeLookup("DIMS", split[0]);
                if (criteriaAttribute != null)
                {
                    //work around for feet and inch because they don't give us the format in the sheet 
                    string format;
                    switch (split[2])
                    {
                        case "ft":
                            format = "'";
                            break;
                        case "in":
                            format = "\"";
                            break;
                        default:
                            format = split[2];
                            break;
                    }
                                       
                    retVal += split[1] + " " + format + " x ";                    
                }
            }

            var lastPos = retVal.LastIndexOf(" x ", StringComparison.Ordinal);
            if (lastPos > 0)
                retVal = retVal.Remove(lastPos, 3);

            return retVal;
        }
    }
}
