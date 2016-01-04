using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using Newtonsoft.Json;
using System.Net.Http.Headers;

namespace ImportPOC2.DataFetchers
{
    public static class Lookup
    {
        private static readonly HttpClient RadarHttpClient = new HttpClient { BaseAddress = new Uri("http://local-espupdates.asicentral.com/api/api/") };

        static Lookup()
        {
            RadarHttpClient.DefaultRequestHeaders.Accept.Clear();
            RadarHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public static List<KeyValueLookUp> GetMatchingTradenames(string q)
        {
            List<KeyValueLookUp> tradenamesList = new List<KeyValueLookUp>();

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
