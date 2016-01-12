using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using Newtonsoft.Json;
using System.Net.Http.Headers;

namespace ImportPOC2.DataFetchers
{
    public static class Lookup
    {
        private static HttpClient RadarHttpClient;

        static Lookup()
        {
            var baseUri = ConfigurationManager.AppSettings["radarApiLocation"] ?? string.Empty;
            RadarHttpClient = new HttpClient { BaseAddress = new Uri(baseUri) };

            RadarHttpClient.DefaultRequestHeaders.Accept.Clear();
            RadarHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public static List<GenericLookUp> GetMatchingTradenames(string q)
        {
            List<GenericLookUp> tradeNamesLookup = null;

            var results = RadarHttpClient.GetAsync("lookup/trade_names?q=" + q).Result;
            if (results.IsSuccessStatusCode)
            {
                var content = results.Content.ReadAsStringAsync().Result;
                var fromRadar = JsonConvert.DeserializeObject<List<KeyValueLookUp>>(content);

                //decouple radar lookup from public version
                tradeNamesLookup = new List<GenericLookUp>();
                tradeNamesLookup.AddRange(fromRadar.Select(s => new GenericLookUp { CodeValue = s.Value, ID = s.Key }));               
            }

            return tradeNamesLookup;                        
        }
    }
}
