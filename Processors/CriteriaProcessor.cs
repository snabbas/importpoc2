using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
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
                var prodConfig = _currentProduct.ProductConfigurations.FirstOrDefault(c => c.IsDefault);

                if (prodConfig != null)
                {
                    var cSets = prodConfig.ProductCriteriaSets.Where(c => c.CriteriaCode == criteriaCode).ToList();
                    retVal = !string.IsNullOrWhiteSpace(optionName) ? cSets.FirstOrDefault(c => string.Equals(c.CriteriaDetail, optionName, StringComparison.CurrentCultureIgnoreCase)) : cSets.FirstOrDefault();
                }

                retVal = retVal ?? addCriteriaSet(criteriaCode, optionName);
            }
            return retVal;
        }

        private ProductCriteriaSet addCriteriaSet(string criteriaCode, string optionName = "")
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
        public IEnumerable<CriteriaSetValue> GetCSValuesByCriteriaCode (string criteriaCode, string optionName = "")
        {
            var cSet = GetCriteriaSetByCode(criteriaCode, optionName);
            var result = cSet.CriteriaSetValues.ToList();

            return result;
        }


        //TODO: can we "detect" what value type code to use instead of passing it in? 
        //TODO: pass in criteria set, it's known from everywhere it is invoked
        public void CreateNewValue(string criteriaCode, object value, long setCodeValueId, string valueTypeCode = "LOOK", string valueDetail = "", string optionName = "", CriteriaSetCodeValueLink childCriteriaSetCodeValue = null)
        {
            var cSet = GetCriteriaSetByCode(criteriaCode, optionName);

            //create new criteria set value
            var newCsv = new CriteriaSetValue
            {
                CriteriaCode = criteriaCode,
                CriteriaSetId = cSet.CriteriaSetId,
                Value = value,
                ID = Utils.IdGenerator.getNextid(),
                ValueTypeCode = valueTypeCode,
                CriteriaValueDetail = valueDetail,
                FormatValue = value.ToString() //default formatvalue to be same as value
            };

            //create new criteria set code value
            var newCscv = new CriteriaSetCodeValue
            {
                CriteriaSetValueId = newCsv.ID,
                SetCodeValueId = setCodeValueId,               
                ID = Utils.IdGenerator.getNextid(),
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
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault().CriteriaSetCodeValues.FirstOrDefault(cv => cv.SetCodeValueId == e.SetCodeValueId);
                    criteriaSet.CriteriaSetValues.FirstOrDefault().CriteriaSetCodeValues.Remove(toDelete);
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
                    var toDelete = criteriaSetValue.CriteriaSetCodeValues.FirstOrDefault().ChildCriteriaSetCodeValues.FirstOrDefault(scv => scv == e);
                    criteriaSetValue.CriteriaSetCodeValues.FirstOrDefault().ChildCriteriaSetCodeValues.Remove(toDelete);
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
                    let scv = v.CriteriaSetCodeValues.FirstOrDefault(s => s.SetCodeValueId == scvId)
                    where scv != null
                        && v.FormatValue == formatValue
                    select v)
                    .FirstOrDefault();
        }

        public ProductCriteriaSet getSizeCriteriaSetByCode(string criteriaCode)
        {
            ProductCriteriaSet retVal = null;
            var prodConfig = _currentProduct.ProductConfigurations.FirstOrDefault(c => c.IsDefault);

            if (prodConfig != null)
            {
                var cSet = prodConfig.ProductCriteriaSets.FirstOrDefault(c => c.CriteriaCode == criteriaCode);
                if (cSet == null)
                {
                    cSet = prodConfig.ProductCriteriaSets.FirstOrDefault(c => Constants.CriteriaCodes.SIZE.Contains(c.CriteriaCode));
                    //create a new size criteria set if none already exists
                    if (cSet == null)
                    {
                        retVal = addCriteriaSet(criteriaCode);
                    }
                    else
                    {
                        //if another size criteria set already exists replace it with the new size criteria set
                        removeCriteriaSet(cSet.CriteriaCode);
                        retVal = addCriteriaSet(criteriaCode);
                    }
                }
                else
                {
                    retVal = cSet;
                }
            }

            return retVal;
        }

        public void removeCriteriaSet(string criteriaCode)
        {
            var productConfiguration = _currentProduct.ProductConfigurations.FirstOrDefault(cfg => cfg.IsDefault);
            if (productConfiguration != null)
            {
                var cs = productConfiguration.ProductCriteriaSets.FirstOrDefault(c => c.CriteriaCode == criteriaCode);
                if (cs != null)
                    productConfiguration.ProductCriteriaSets.Remove(cs);
            }
        }
    }
}
