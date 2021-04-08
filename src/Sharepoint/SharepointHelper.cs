using System;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using TimHanewich.MicrosoftGraphHelper;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace TimHanewich.MicrosoftGraphHelper.Sharepoint
{
    public static class SharepointHelper
    {
        public static async Task<SharepointSite[]> SearchSharepointSitesAsync(this MicrosoftGraphHelper mgh, string query)
        {
            HttpRequestMessage req = mgh.PrepareHttpRequestMessage();
            req.Method = HttpMethod.Get;
            req.RequestUri = new Uri("https://graph.microsoft.com/v1.0/sites?search=" + HttpUtility.UrlEncode(query));
            HttpClient hc = new HttpClient();
            HttpResponseMessage msg = await hc.SendAsync(req);
            string content = await msg.Content.ReadAsStringAsync();
            if (msg.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception("Search of Sharepoint Sites failed with code \"" + msg.StatusCode.ToString() + "\". Body: " + content);
            }
            JObject jo = JObject.Parse(content);

            //Get them
            JArray ja_value = JArray.Parse(jo.Property("value").Value.ToString());
            List<SharepointSite> Sites = new List<SharepointSite>();
            foreach (JObject jo_ss in ja_value)
            {
                SharepointSite ss = new SharepointSite(jo_ss.ToString());
                Sites.Add(ss);
            }

            return Sites.ToArray();
        }
    }
}