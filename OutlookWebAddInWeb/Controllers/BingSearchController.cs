using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Newtonsoft.Json;
using System.Net.Http;

using OutlookWebAddInWeb.Bing;
using System.Web.Script.Serialization;
using System.Net;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;

namespace OutlookWebAddInWeb.Controllers
{
    public class WebSer
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Url { get; set; }
        public string Snippet { get; set; }
        public string DisplayUrl { get; set; }
    }
    public class ws
    {
        public WebSer WSer { get; set; }
    }

    public class BingSearchController : Controller
    {
        private const string rootUrl = "https://api.cognitive.microsoft.com/bing/v5.0/search";
        private const string AccountKey = "YOUR KEY HERE";
       
        // GET: BingSearch
        public async Task<JsonResult> Get(string term)
        {
            List<WebSer> sres = new List<WebSer>();

            var client = new HttpClient();
            client.DefaultRequestHeaders.TryAddWithoutValidation("Ocp-Apim-Subscription-Key", AccountKey);
            var result = await client.GetAsync(string.Format("{0}?q=site:docs.microsoft.com {1}&count={2}&offset={3}&mkt={4}", rootUrl, WebUtility.UrlEncode(term), 15, 0, "en-us"));

            var json = await result.Content.ReadAsStringAsync();

            dynamic data = JObject.Parse(json);

            for (int i = 0; i < 10; i++)
            {
                sres.Add(new WebSer
                {
                    Id = data.webPages.value[i].id,
                    Name = data.webPages.value[i].name,
                    Url = data.webPages.value[i].url,
                    Snippet = data.webPages.value[i].snippet,
                    DisplayUrl = data.webPages.value[i].displayUrl
                });
            }

            return Json(sres, JsonRequestBehavior.AllowGet);
           
        }
    }
}