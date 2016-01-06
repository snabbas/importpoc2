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

        public static List<KeyValueLookUp> GetMatchingTradenames(string q)
        {
            var tradenamesList = new List<KeyValueLookUp>();

            var results = RadarHttpClient.GetAsync("lookup/trade_names?q=" + q).Result;
            if (results.IsSuccessStatusCode)
            {
                var content = results.Content.ReadAsStringAsync().Result;
                tradenamesList = JsonConvert.DeserializeObject<List<KeyValueLookUp>>(content);
            }

            return tradenamesList;                        
        }
    }
}
