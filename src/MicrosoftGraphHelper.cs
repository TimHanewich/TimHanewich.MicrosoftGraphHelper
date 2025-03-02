using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TimHanewich.MicrosoftGraphHelper
{
    public class MicrosoftGraphHelper
    {
        //Standard inputs - Authorization Specific
        private string login_base_url;
        private string authorize_endpoint;
        
        //Standard inputs - Token Request Specific
        private string token_endpoint;

        //Custom inputs
        public string Tenant {get; set;} //Can be "common" for both Microsoft accounts and work/school accounts, "organizations" for work/school accounts only, "consumers" for Microsoft accounts only, of a tenant identifier (GUID). According to this documentation: https://learn.microsoft.com/en-us/graph/auth-v2-user
        public Guid ClientId {get; set;}
        public List<string> Scope {get; set;}
        public string RedirectUrl {get; set;}

        //Last received payload
        public MicrosoftGraphTokenPayload LastReceivedTokenPayload {get; set;}

        public MicrosoftGraphHelper()
        {
            Scope = new List<string>();

            //Set private inputs
            login_base_url = "https://login.microsoftonline.com";
            authorize_endpoint = "oauth2/v2.0/authorize";
            token_endpoint = "oauth2/v2.0/token";
        }

        public string AssembleAuthorizationUrl(bool include_offline_access = true, bool always_show_consent = true)
        {
            string ToReturn = login_base_url + "/" + Tenant + "/" + authorize_endpoint;
            ToReturn = ToReturn + "?client_id=" + ClientId.ToString();
            ToReturn = ToReturn + "&response_type=code"; //Standard
            ToReturn = ToReturn + "&redirect_uri=" + RedirectUrl;
            ToReturn = ToReturn + "&response_mode=query"; //Standard

            //Scopes
            List<string> ScopesToUse = new List<string>();
            if (Scope != null)
            {
                ScopesToUse.AddRange(Scope);
            }
            if (ScopesToUse.Contains("offline_access") == false)
            {
                ScopesToUse.Add("offline_access");
            }
            ToReturn = ToReturn + "&scope=" + UrlEncodeScopes(ScopesToUse.ToArray());

            if (always_show_consent)
            {
                ToReturn = ToReturn + "&prompt=consent";
            }

            return ToReturn;
        }

        public async Task GetAccessTokenAsync(string authorization_code)
        {
            //Assemble the URL to request from using the code
            string ReqUrl = login_base_url + "/" + Tenant + "/" + token_endpoint;
            
            List<KeyValuePair<string, string>> KVPs = new List<KeyValuePair<string, string>>();
            KVPs.Add(new KeyValuePair<string, string>("client_id",ClientId.ToString()));
            
            
            //Scopes
            // List<string> ScopesToUse = new List<string>();
            // if (Scope != null)
            // {
            //     ScopesToUse.AddRange(Scope);
            // }
            // if (ScopesToUse.Contains("offline_access") == false)
            // {
            //     ScopesToUse.Add("offline_access");
            // }
            // KVPs.Add(new KeyValuePair<string, string>("scope",UrlEncodeScopes(ScopesToUse.ToArray())));

            KVPs.Add(new KeyValuePair<string, string>("code", authorization_code));
            KVPs.Add(new KeyValuePair<string, string>("redirect_uri", RedirectUrl));
            KVPs.Add(new KeyValuePair<string, string>("grant_type", "authorization_code"));

            string asstr = await new FormUrlEncodedContent(KVPs).ReadAsStringAsync();
           

            //Make the request
            HttpClient hc = new HttpClient();
            HttpRequestMessage reqmsg = new HttpRequestMessage();
            reqmsg.RequestUri = new Uri(ReqUrl);
            reqmsg.Method = HttpMethod.Post;
            reqmsg.Content = new StringContent(asstr, Encoding.UTF8, "application/x-www-form-urlencoded");
            HttpResponseMessage hrm = await hc.SendAsync(reqmsg);
            string content = await hrm.Content.ReadAsStringAsync();
            if (hrm.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception("Error code \"" + hrm.StatusCode.ToString() + "\" returned. Content: " + content);
            }
            
            //Parse into token payload
            MicrosoftGraphTokenPayload tokenpayload = new MicrosoftGraphTokenPayload(content);

            //Store for later
            LastReceivedTokenPayload = tokenpayload;
        }

        public async Task RefreshAccessTokenAsync()
        {
            //Assemble the URL to request from using the code
            string ReqUrl = login_base_url + "/" + Tenant + "/" + token_endpoint;
            
            List<KeyValuePair<string, string>> KVPs = new List<KeyValuePair<string, string>>();
            KVPs.Add(new KeyValuePair<string, string>("client_id",ClientId.ToString()));
            
            
            //Scopes
            // List<string> ScopesToUse = new List<string>();
            // if (Scope != null)
            // {
            //     ScopesToUse.AddRange(Scope);
            // }
            // if (ScopesToUse.Contains("offline_access") == false)
            // {
            //     ScopesToUse.Add("offline_access");
            // }
            // KVPs.Add(new KeyValuePair<string, string>("scope",UrlEncodeScopes(ScopesToUse.ToArray())));

            KVPs.Add(new KeyValuePair<string, string>("refresh_token", LastReceivedTokenPayload.RefreshToken));
            KVPs.Add(new KeyValuePair<string, string>("redirect_uri", RedirectUrl));
            KVPs.Add(new KeyValuePair<string, string>("grant_type", "refresh_token"));

            string asstr = await new FormUrlEncodedContent(KVPs).ReadAsStringAsync();
           

            //Make the request
            HttpClient hc = new HttpClient();
            HttpRequestMessage reqmsg = new HttpRequestMessage();
            reqmsg.RequestUri = new Uri(ReqUrl);
            reqmsg.Method = HttpMethod.Post;
            reqmsg.Content = new StringContent(asstr, Encoding.UTF8, "application/x-www-form-urlencoded");
            HttpResponseMessage hrm = await hc.SendAsync(reqmsg);
            string content = await hrm.Content.ReadAsStringAsync();
            if (hrm.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception("Error code \"" + hrm.StatusCode.ToString() + "\" returned. Content: " + content);
            }
            
            //Parse into token payload
            MicrosoftGraphTokenPayload tokenpayload = new MicrosoftGraphTokenPayload(content);

            //Store for later
            LastReceivedTokenPayload = tokenpayload;
        }

        public async Task RefreshAccessTokenIfExpiredAsync()
        {
            if (AccessTokenHasExpired())
            {
                await RefreshAccessTokenAsync();
            }
        }

        public bool AccessTokenHasExpired()
        {
            if (LastReceivedTokenPayload == null)
            {
                throw new Exception("An access token has not been received yet.");
            }

            if (DateTime.UtcNow > LastReceivedTokenPayload.ExpiresAtUtc)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public HttpRequestMessage PrepareHttpRequestMessage()
        {
            if (LastReceivedTokenPayload == null)
            {
                throw new Exception("Unable to prepare Http Client with authorization... access token has not been received!");
            }

            HttpRequestMessage req = new HttpRequestMessage();
            req.Headers.Add("Authorization", "Bearer " + LastReceivedTokenPayload.AccessToken);
            return req;            
        }

        private string UrlEncodeScopes(string[] ToEncode)
        {
            string ToReturn = "";
            if (ToEncode != null)
            {
                foreach (string s in ToEncode)
                {
                    ToReturn = ToReturn + s + "%20";
                }
                ToReturn = ToReturn.Substring(0, ToReturn.Length - 3); //Remove the last "%20"
            }
            return ToReturn;
        }
        
    }
}