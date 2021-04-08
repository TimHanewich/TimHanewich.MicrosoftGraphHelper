using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace TimHanewich.MicrosoftGraphHelper.Sharepoint
{
    public class SharepointSite
    {
        public DateTime CreatedAt {get; set;}
        public string Description {get; set;}
        public Guid Id {get; set;}
        public DateTime LastModified {get; set;}
        public string Name {get; set;}
        public string Url {get; set;}
        public string DisplayName {get; set;}

        public SharepointSite()
        {

        }

        public SharepointSite(string json_payload)
        {
            JObject jo = JObject.Parse(json_payload);

            //Created At
            JProperty prop_createdDateTime = jo.Property("createdDateTime");
            if (prop_createdDateTime != null)
            {
                if (prop_createdDateTime.Value.Type != JTokenType.Null)
                {
                    CreatedAt = DateTime.Parse(prop_createdDateTime.Value.ToString());
                }
            }

            //description
            JProperty prop_description = jo.Property("description");
            if (prop_description != null)
            {
                if (prop_description.Value.Type != JTokenType.Null)
                {
                    Description = prop_description.Value.ToString();
                }
            }

            //Id
            JProperty prop_id = jo.Property("id");
            if (prop_id != null)
            {
                if (prop_id.Value.Type != JTokenType.Null)
                {
                    List<string> Splitter = new List<string>();
                    Splitter.Add(",");
                    string[] parts = prop_id.Value.ToString().Split(Splitter.ToArray(), StringSplitOptions.None);
                    Id = Guid.Parse(Splitter[1]);
                }
            }

            //Last Modified
            JProperty prop_lastModifiedDateTime = jo.Property("lastModifiedDateTime");
            if (prop_lastModifiedDateTime != null)
            {
                if (prop_lastModifiedDateTime.Value.Type != JTokenType.Null)
                {
                    LastModified = DateTime.Parse(prop_lastModifiedDateTime.Value.ToString());
                }
            }

            //Name
            JProperty prop_name = jo.Property("name");
            if (prop_name != null)
            {
                if (prop_name.Value.Type != JTokenType.Null)
                {
                    Name = prop_name.Value.ToString();
                }
            }

            //Web url
            JProperty prop_url = jo.Property("webUrl");
            if (prop_url != null)
            {
                if (prop_url.Value.Type != JTokenType.Null)
                {
                    Url = prop_url.Value.ToString();
                }
            }

            //Display Name
            JProperty prop_displayName = jo.Property("displayName");
            if (prop_displayName != null)
            {
                if (prop_displayName.Value.Type != JTokenType.Null)
                {
                    DisplayName = prop_displayName.Value.ToString();
                }
            }
        }
    }
}