using System;
using System.Collections.Generic;
using System.Configuration;
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
            var temp_list = new List<KeyValueLookUp>();
            var tradeNamesLookup = new List<GenericLookUp>();

            var results = RadarHttpClient.GetAsync("lookup/trade_names?q=" + q).Result;
            if (results.IsSuccessStatusCode)
            {
                var content = results.Content.ReadAsStringAsync().Result;
                temp_list = JsonConvert.DeserializeObject<List<KeyValueLookUp>>(content);

                if (temp_list != null)
                {
                    temp_list.ForEach(t => tradeNamesLookup.Add(new GenericLookUp { ID = t.Key, CodeValue = t.Value }));
                }
            }

            return tradeNamesLookup;                        
        }
    }
}
