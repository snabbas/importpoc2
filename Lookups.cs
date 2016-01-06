using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Radar.Models.Company;
using Radar.Models.Criteria;

namespace ImportPOC2
{
    public class Lookups
    {
        public static HttpClient RadarHttpClient;
        public static int CurrentCompanyId;

        private static List<GenericLookUp> _imprintMethodsLookup = null;
        public static List<GenericLookUp> ImprintMethodsLookup
        {
            get
            {
                if (_imprintMethodsLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/imprint_methods").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _imprintMethodsLookup = JsonConvert.DeserializeObject<List<GenericLookUp>>(content);
                    }
                }
                return _imprintMethodsLookup;
            }
            set { _imprintMethodsLookup = value; }
        }

        private static List<Category> _catlist = null;
        public static List<Category> CategoryList
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
        public static List<ProductColorGroup> ColorGroupList
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
        public static List<KeyValueLookUp> ShapesLookup
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
        public static List<SetCodeValue> ThemesLookup
        {
            get
            {
                if (_themesLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/themes").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        var themeGroups = JsonConvert.DeserializeObject<List<ThemeLookUp>>(content);
                        _themesLookup = new List<SetCodeValue>();

                        themeGroups.ForEach(t => _themesLookup.AddRange(t.SetCodeValues));
                    }
                }
                return _themesLookup;
            }
            set { _themesLookup = value; }
        }

        private static List<GenericLookUp> _originsLookup = null;
        public static List<GenericLookUp> OriginsLookup
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
        public static List<KeyValueLookUp> PackagingLookup
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

        private static List<KeyValueLookUp> _complianceLookup = null;
        public static List<KeyValueLookUp> ComplianceLookup
        {
            get
            {
                if (_complianceLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/compliance").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _complianceLookup = JsonConvert.DeserializeObject<List<KeyValueLookUp>>(content);
                    }
                }
                return _complianceLookup;
            }
            set { _complianceLookup = value; }
        }

        private static List<SafetyWarningLookUp> _safetywarningsLookup = null;
        public static List<SafetyWarningLookUp> SafetywarningsLookup
        {
            get
            {
                if (_safetywarningsLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/safetywarnings").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _safetywarningsLookup = JsonConvert.DeserializeObject<List<SafetyWarningLookUp>>(content);
                    }
                }
                return _safetywarningsLookup;
            }
            set { _safetywarningsLookup = value; }
        }

        private static List<CurrencyLookUp> _currencyLookup = null;
        public static List<CurrencyLookUp> CurrencyLookup
        {
            get
            {
                if (_currencyLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/currency").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _currencyLookup = JsonConvert.DeserializeObject<List<CurrencyLookUp>>(content);
                    }
                }
                return _currencyLookup;
            }
            set { _currencyLookup = value; }
        }

        private static List<CostTypeLookUp> _costTypesLookup = null;
        public static List<CostTypeLookUp> CostTypesLookup
        {
            get
            {
                if (_costTypesLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/cost_types").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _costTypesLookup = JsonConvert.DeserializeObject<List<CostTypeLookUp>>(content);
                    }
                }
                return _costTypesLookup;
            }
            set { _costTypesLookup = value; }
        }

        private static List<KeyValueLookUp> _inventoryStatusesLookup = null;
        public static List<KeyValueLookUp> InventoryStatusesLookup
        {
            get
            {
                if (_inventoryStatusesLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/inventory_statuses").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _inventoryStatusesLookup = JsonConvert.DeserializeObject<List<KeyValueLookUp>>(content);
                    }
                }
                return _inventoryStatusesLookup;
            }
            set { _inventoryStatusesLookup = value; }
        }

        private static List<CriteriaAttribute> _criteriaAttributeLookup = null;
        public static CriteriaAttribute CriteriaAttributeLookup(string code, string name)
        {
            var criteriaAttribute = new CriteriaAttribute();
            if (_criteriaAttributeLookup == null)
            {
                var results = RadarHttpClient.GetAsync("lookup/criteria_attributes").Result;
                if (results.IsSuccessStatusCode)
                {
                    var content = results.Content.ReadAsStringAsync().Result;
                    _criteriaAttributeLookup = JsonConvert.DeserializeObject<List<CriteriaAttribute>>(content);
                }
            }

            if (!string.IsNullOrWhiteSpace(code) && !string.IsNullOrWhiteSpace(name))
            {
                if (_criteriaAttributeLookup != null)
                    criteriaAttribute = _criteriaAttributeLookup.FirstOrDefault(u => u.CriteriaCode == code && u.Description == name);
            }

            return criteriaAttribute;
        }

        private static List<ImprintCriteriaLookUp> _imprintCriteriaLookup = null;
        public static List<ImprintCriteriaLookUp> ImprintCriteriaLookup
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

        private static List<CriteriaItem> _imprintColorLookup = null;
        public static List<CriteriaItem> ImprintColorLookup
        {
            get
            {
                if (_imprintCriteriaLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/criteria?code=COLR").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _imprintColorLookup = JsonConvert.DeserializeObject<List<CriteriaItem>>(content);
                    }
                }
                return _imprintColorLookup;
            }
            set { _imprintColorLookup = value; }
        }

        private static List<CriteriaItem> _imprintSizeLocation = null;
        public static List<CriteriaItem> ImprintSizeLocation
        {
            get
            {
                if (_imprintSizeLocation == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/criteria?code=IMSZ").Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _imprintColorLookup = JsonConvert.DeserializeObject<List<CriteriaItem>>(content);
                    }
                }
                return _imprintSizeLocation;
            }
            set { _imprintSizeLocation = value; }
        }

        private static List<LineName> _linenamesLookup = null;
        public static List<LineName> LinenamesLookup
        {
            get
            {
                if (_linenamesLookup == null)
                {
                    var results = RadarHttpClient.GetAsync("lookup/linenames?company_id=" + CurrentCompanyId).Result;
                    if (results.IsSuccessStatusCode)
                    {
                        var content = results.Content.ReadAsStringAsync().Result;
                        _linenamesLookup = JsonConvert.DeserializeObject<List<LineName>>(content);
                    }
                }
                return _linenamesLookup;
            }
            set { _linenamesLookup = value; }
        }


    }
}
