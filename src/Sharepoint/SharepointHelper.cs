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
    
        public static async Task<SharepointList[]> ListSharepointListsAsync(this MicrosoftGraphHelper mgh, Guid site_id)
        {
            HttpRequestMessage req = mgh.PrepareHttpRequestMessage();
            req.Method = HttpMethod.Get;
            req.RequestUri = new Uri("https://graph.microsoft.com/v1.0/sites/" + site_id.ToString() + "/lists");
            HttpClient hc = new HttpClient();
            HttpResponseMessage msg = await hc.SendAsync(req);
            string content = await msg.Content.ReadAsStringAsync();
            if (msg.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception("Listing lists from sharepoint site '" + site_id.ToString() + "' failed with code \"" + msg.StatusCode.ToString() + "\". Body: " + content);
            }
            JObject jo = JObject.Parse(content);

            //Get them
            JArray ja_value = JArray.Parse(jo.Property("value").Value.ToString());
            List<SharepointList> Lists = new List<SharepointList>();
            foreach (JObject jo_sl in ja_value)
            {
                SharepointList sl = SharepointList.ParseFromJsonPayload(jo_sl.ToString());
                Lists.Add(sl);
            }
            return Lists.ToArray();
        }
    
        public static async Task<SharepointListItem[]> GetAllItemsFromSharepointListAsync(this MicrosoftGraphHelper mgh, Guid site_id, Guid list_id)
        {
            HttpRequestMessage req = mgh.PrepareHttpRequestMessage();
            req.Method = HttpMethod.Get;
            req.RequestUri = new Uri("https://graph.microsoft.com/v1.0/sites/" + site_id.ToString() + "/lists/" + list_id.ToString() + "/items?expand=fields");
            HttpClient hc = new HttpClient();
            HttpResponseMessage msg = await hc.SendAsync(req);
            string content = await msg.Content.ReadAsStringAsync();
            if (msg.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception("Getting all items from sharepoint list '" + list_id.ToString() + "' failed with code \"" + msg.StatusCode.ToString() + "\". Body: " + content);
            }
            JObject jo = JObject.Parse(content);

            //Get them
            JArray ja_value = JArray.Parse(jo.Property("value").Value.ToString());
            List<SharepointListItem> SPitems = new List<SharepointListItem>();
            foreach (JObject jo_li in ja_value)
            {
                SharepointListItem spli = SharepointListItem.ParseFromJsonPayload(jo_li.ToString());
                SPitems.Add(spli);
            }
            return SPitems.ToArray();
        }
    
        public static async Task CreateItemAsync(this MicrosoftGraphHelper mgh, Guid site_id, Guid list_id, JObject fields)
        {
            HttpRequestMessage req = mgh.PrepareHttpRequestMessage();
            req.Method = HttpMethod.Post;
            req.RequestUri = new Uri("https://graph.microsoft.com/v1.0/sites/" + site_id.ToString() + "/lists/" + list_id.ToString() + "/items");
            
            //Create the body
            JObject jo = new JObject();
            jo.Add("fields", fields);
            req.Content = new StringContent(jo.ToString(), System.Text.Encoding.UTF8, "application/json");
            
            //Send!
            HttpClient hc = new HttpClient();
            HttpResponseMessage resp = await hc.SendAsync(req);
            if (resp.StatusCode != HttpStatusCode.Created)
            {
                string msg = await resp.Content.ReadAsStringAsync();
                throw new Exception("Creation of new record failed. Msg: " + msg);
            }
        }
    }
}