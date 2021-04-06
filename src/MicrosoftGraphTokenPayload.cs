using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TimHanewich.MicrosoftGraphHelper
{
    public class MicrosoftGraphTokenPayload
    {
        public string[] Scope {get; set;}
        public DateTime ReceivedAtUtc {get; set;}
        public DateTime ExpiresAtUtc {get; set;}
        public string AccessToken {get; set;}
        public string RefreshToken {get; set;}

        public MicrosoftGraphTokenPayload()
        {

        }

        public MicrosoftGraphTokenPayload(string payload_json)
        {
            JObject jo = JObject.Parse(payload_json);

            //Scopes
            JProperty prop_scope = jo.Property("scope");
            string prop_scope_str = prop_scope.Value.ToString();
            Scope = prop_scope_str.Split(" ");

            //Received at and expires at
            ReceivedAtUtc = DateTime.UtcNow;
            JProperty prop_ExpiresIn = jo.Property("expires_in");
            int ExpiresInSecs = Convert.ToInt32(prop_ExpiresIn.Value.ToString());
            ExpiresAtUtc = ReceivedAtUtc.AddSeconds(ExpiresInSecs);

            //Access token
            AccessToken = jo.Property("access_token").Value.ToString();

            //refresh token
            JProperty prop_refresh_token = jo.Property("refresh_token");
            if (prop_refresh_token != null)
            {
                if (prop_refresh_token.Type != JTokenType.Null)
                {
                    RefreshToken = prop_refresh_token.Value.ToString();
                }
            }
        }
    }
}