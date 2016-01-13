﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Radar.Models.Criteria;
using Radar.Models.Product;

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
            var prodConfig = _currentProduct.ProductConfigurations.FirstOrDefault(c => c.IsDefault);

            if (prodConfig != null)
            {
                var cSets = prodConfig.ProductCriteriaSets.Where(c => c.CriteriaCode == criteriaCode).ToList();
                retVal = !string.IsNullOrWhiteSpace(optionName) ? cSets.FirstOrDefault(c =>  string.Equals(c.CriteriaDetail , optionName, StringComparison.CurrentCultureIgnoreCase)) : cSets.FirstOrDefault();
            }

            retVal = retVal ?? addCriteriaSet(criteriaCode, optionName);

            return retVal;
        }

        private ProductCriteriaSet addCriteriaSet(string criteriaCode, string optionName)
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
        public void CreateNewValue(string criteriaCode, object value, long setCodeValueId, string valueTypeCode = "LOOK", string valueDetail = "", string optionName = "")
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
                ID = Utils.IdGenerator.getNextid()
            };

            newCsv.CriteriaSetCodeValues.Add(newCscv);
            cSet.CriteriaSetValues.Add(newCsv);
        }

        //TODO: pretty sure this can be done without passing in criteriaset parameter
        public void DeleteCsValues(IEnumerable<CriteriaSetValue> entities, IEnumerable<string> models, ProductCriteriaSet criteriaSet)
        {
            //delete values that are missing from the list in the file
            var valuesToDelete = entities.Select(e => e.Value).Except(models).Select(s => s).ToList();
            valuesToDelete.ForEach(e =>
            {
                var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault(v => v.Value == e);
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

        public void deleteCodeValues(CriteriaSetValue entity, IEnumerable<long> models, ProductCriteriaSet criteriaSet)
        {
            if (entity != null)
            {
                var exists = false;
                var csValuesToDelete = new List<CriteriaSetCodeValue>();

                entity.CriteriaSetCodeValues.ToList().ForEach(e =>
                {
                    exists = models.Any(m => m == e.SetCodeValueId);

                    if (!exists)
                    {
                        csValuesToDelete.Add(e);
                    }
                });

                csValuesToDelete.ForEach(e =>
                {
                    var toDelete = criteriaSet.CriteriaSetValues.FirstOrDefault().CriteriaSetCodeValues.FirstOrDefault(cv => cv.SetCodeValueId == e.SetCodeValueId);
                    criteriaSet.CriteriaSetValues.FirstOrDefault().CriteriaSetCodeValues.Remove(toDelete);
                });
            }
        }
    }
}
