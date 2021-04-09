using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace TimHanewich.MicrosoftGraphHelper.Sharepoint
{
    public class SharepointListItem
    {
        public DateTime CreatedAt {get; set;}
        public string Id {get; set;}
        public DateTime LastModified {get; set;}
        public string Url {get; set;}
        public AzureAdUser CreatedBy {get; set;}
        public AzureAdUser LastModifiedBy {get; set;}
        public SharepointListItemField[] Fields {get; set;}

        public static SharepointListItem ParseFromJsonPayload(string payload)
        {
            SharepointListItem ToReturn = new SharepointListItem();

            JObject jo = JObject.Parse(payload);

            //Created At
            JProperty prop_createdDateTime = jo.Property("createdDateTime");
            if (prop_createdDateTime != null)
            {
                if (prop_createdDateTime.Value.Type != JTokenType.Null)
                {
                    ToReturn.CreatedAt = DateTime.Parse(prop_createdDateTime.Value.ToString());
                }
            }

            //Id
            JProperty prop_id = jo.Property("id");
            if (prop_id != null)
            {
                if (prop_id.Value.Type != JTokenType.Null)
                {
                    ToReturn.Id = prop_id.Value.ToString();
                }
            }

            //Last Modified
            JProperty prop_lastModifiedDateTime = jo.Property("lastModifiedDateTime");
            if (prop_lastModifiedDateTime != null)
            {
                if (prop_lastModifiedDateTime.Value.Type != JTokenType.Null)
                {
                    ToReturn.LastModified = DateTime.Parse(prop_lastModifiedDateTime.Value.ToString());
                }
            }

            //Web url
            JProperty prop_url = jo.Property("webUrl");
            if (prop_url != null)
            {
                if (prop_url.Value.Type != JTokenType.Null)
                {
                    ToReturn.Url = prop_url.Value.ToString();
                }
            }

            //Created by user
            JObject jo_createdBy = JObject.Parse(jo.Property("createdBy").Value.ToString());
            if (jo_createdBy != null)
            {
                JObject jo_user = JObject.Parse(jo_createdBy.Property("user").Value.ToString());
                if (jo_user != null)
                {
                    AzureAdUser user = AzureAdUser.ParseFromJsonPayload(jo_user.ToString());
                    ToReturn.CreatedBy = user;
                }   
            }

            //Last modified by user
            JObject jo_lastModifiedBy = JObject.Parse(jo.Property("lastModifiedBy").Value.ToString());
            if (jo_createdBy != null)
            {
                JObject jo_user = JObject.Parse(jo_lastModifiedBy.Property("user").Value.ToString());
                if (jo_user != null)
                {
                    AzureAdUser user = AzureAdUser.ParseFromJsonPayload(jo_user.ToString());
                    ToReturn.LastModifiedBy = user;
                }   
            }

            //Fields
            List<SharepointListItemField> fields = new List<SharepointListItemField>();
            JProperty prop_fields = jo.Property("fields");
            if (prop_fields != null)
            {
                JObject jo_fields = JObject.Parse(jo.Property("fields").Value.ToString());
                foreach (JProperty prop in jo_fields.Properties())
                {
                    SharepointListItemField thisfield = new SharepointListItemField();
                    thisfield.Label = prop.Name;
                    if (prop.Value.Type != JTokenType.Null)
                    {
                        thisfield.Value = prop.Value.ToString();
                    }
                    else
                    {
                        thisfield.Value = null;
                    }
                    fields.Add(thisfield);
                }
            }
            ToReturn.Fields = fields.ToArray();
            

            return ToReturn;
        }
    }
}