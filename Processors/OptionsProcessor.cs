using ImportPOC2.Models;
using ImportPOC2.Utils;
using Radar.Core.Common;
using Radar.Core.Models.Batch;
using Radar.Models;
using Radar.Models.Product;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ImportPOC2.Processors
{
    public class OptionsProcessor
    {
        private CriteriaProcessor _criteriaProcessor;              
        private List<ProductCriteriaSet> productOptionsMap;

        public OptionsProcessor(CriteriaProcessor criteriaProcessor)
        {
            _criteriaProcessor = criteriaProcessor;
            productOptionsMap = new List<ProductCriteriaSet>();
        }

        private long GetSetCodeValueIdByCriteriaOption(string criteriaCode)
        {
            var setCodeValueId = 0L;            
            var csCode = string.Empty;
            IEnumerable<ImprintCriteriaLookUp> OptionLookups = null;
            switch (criteriaCode)
            {
                case Constants.CriteriaCodes.ShippingOption:
                    csCode = "SHIP";
                    break;
                case Constants.CriteriaCodes.ProductOption:
                      csCode = "PROD";                   
                    break;
                case Constants.CriteriaCodes.ImprintOption:
                    csCode = "IMPR";
                    break;
            }
            OptionLookups = Lookups.CriteriaLookupByCode(criteriaCode);
            var criteria = OptionLookups.FirstOrDefault(l => l.Code == csCode);
            if (criteria != null)
            { 
                var group = criteria.CodeValueGroups.FirstOrDefault(cvg => string.Equals(cvg.Description, "Other", StringComparison.CurrentCultureIgnoreCase));
                if (group != null)
                {
                    var setCodeValue = group.SetCodeValues.FirstOrDefault();
                    if (setCodeValue != null)
                        setCodeValueId = setCodeValue.ID;
                }             
            }
            return setCodeValueId;
        }

        public void ProcessOptionRow(ProductRow sheetRow, Product productModel)
        {
            if (!string.IsNullOrWhiteSpace(sheetRow.Option_Type) && !string.IsNullOrWhiteSpace(sheetRow.Option_Name))
            {
                var OptionTypeLookUp = new List<CodeValueLookUp>()
                {
                    new CodeValueLookUp { Code = Constants.CriteriaCodes.ShippingOption, Value = "Shipping Option"},
                    new CodeValueLookUp { Code = Constants.CriteriaCodes.ProductOption, Value = "Product Option"},
                    new CodeValueLookUp { Code = Constants.CriteriaCodes.ImprintOption, Value = "Imprint Option"}
                };

                var criteriaCode = string.Empty;
                var lookupCriteria = OptionTypeLookUp.FirstOrDefault(o => string.Equals(o.Value, sheetRow.Option_Type, StringComparison.CurrentCulture));
                if (lookupCriteria != null)
                {
                    criteriaCode = lookupCriteria.Code;
                }

                if (!string.IsNullOrWhiteSpace(criteriaCode))
                {
                    if (!string.IsNullOrWhiteSpace(sheetRow.Option_Name))
                    {
                        //split the values, if it's csv
                        var valueList = sheetRow.Option_Values.ConvertToList();
                        if (valueList.Count() > 0)
                        {
                            var criteriaSet = _criteriaProcessor.GetCriteriaSetByCode(criteriaCode, sheetRow.Option_Name);
                            if (criteriaSet == null)
                            {
                                criteriaSet = _criteriaProcessor.CreateNewCriteriaSet(criteriaCode, sheetRow.Option_Name);
                            }

                            criteriaSet.Description = BasicFieldProcessor.UpdateField(sheetRow.Option_Additional_Info, criteriaSet.Description);
                            criteriaSet.IsRequiredForOrder = BasicFieldProcessor.UpdateField(sheetRow.Req_for_order, criteriaSet.IsRequiredForOrder);
                            criteriaSet.IsMultipleChoiceAllowed = BasicFieldProcessor.UpdateField(sheetRow.Can_order_only_one, criteriaSet.IsMultipleChoiceAllowed);

                            var existingCsvalues = criteriaSet.CriteriaSetValues.ToList();

                            valueList.ForEach(optionValue =>
                            {
                                var optionCS = existingCsvalues.FirstOrDefault(csv => string.Equals(csv.Value, optionValue, StringComparison.CurrentCultureIgnoreCase));
                                var setCodeValueId = GetSetCodeValueIdByCriteriaOption(criteriaCode);
                                //add new value if it doesn't exists
                                if (optionCS == null)
                                {
                                    _criteriaProcessor.CreateNewValue(criteriaSet, optionValue, setCodeValueId, "CUST");                                   
                                }                                
                            });

                            _criteriaProcessor.DeleteCsValues(existingCsvalues, valueList, criteriaSet);
                            productOptionsMap.Add(criteriaSet);
                        }
                        else
                        {
                            // Must have at least one value Log an error
                            BatchProcessor.AddIncorrectFormatError(criteriaCode, "Must have at least one Option_Values"); 
                        }
                    }
                    else
                    {
                        // Option Name is not provided Log an error
                        BatchProcessor.AddIncorrectFormatError(criteriaCode, "Option_Name is required"); 
                    }
                }
                else
                { 
                    // Invalid Option_Type Log an error
                    BatchProcessor.AddIncorrectFormatError(criteriaCode, string.Format("Invalid Option_Type value: {0}", sheetRow.Option_Type)); 
                }
            }
        }

        public void FinalizeOptions()
        {
            var criteriaSetShippingOption = _criteriaProcessor.GetAllCriteriaSetByCode(Constants.CriteriaCodes.ShippingOption);
            var criteriaSetProductOption = _criteriaProcessor.GetAllCriteriaSetByCode(Constants.CriteriaCodes.ProductOption);
            var criteriaSetImprintOption = _criteriaProcessor.GetAllCriteriaSetByCode(Constants.CriteriaCodes.ImprintOption);

            var optionIdsToDelete = criteriaSetShippingOption.Select(s => s.CriteriaSetId).Except(productOptionsMap.Select(m => m.CriteriaSetId)).ToList();
            optionIdsToDelete.AddRange(criteriaSetProductOption.Select(s => s.CriteriaSetId).Except(productOptionsMap.Select(m => m.CriteriaSetId)).ToList());
            optionIdsToDelete.AddRange(criteriaSetImprintOption.Select(s => s.CriteriaSetId).Except(productOptionsMap.Select(m => m.CriteriaSetId)).ToList());

            optionIdsToDelete.ForEach(csId =>
            {
                _criteriaProcessor.removeCriteriaSet(csId);
            });
        }
    }
}
