using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace TimHanewich.MicrosoftGraphHelper.Sharepoint
{
    public class SharepointList
    {
        public DateTime CreatedAt {get; set;}
        public string Description {get; set;}
        public Guid Id {get; set;}
        public DateTime LastModified {get; set;}
        public string Name {get; set;}
        public string Url {get; set;}
        public string DisplayName {get; set;}
        public AzureAdUser CreatedBy {get; set;}
        public AzureAdUser LastModifiedBy {get; set;}
        public SharepointListType ListType {get; set;}

        public static SharepointList ParseFromJsonPayload(string payload)
        {
            SharepointList ToReturn = new SharepointList();

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

            //description
            JProperty prop_description = jo.Property("description");
            if (prop_description != null)
            {
                if (prop_description.Value.Type != JTokenType.Null)
                {
                    ToReturn.Description = prop_description.Value.ToString();
                }
            }

            //Id
            JProperty prop_id = jo.Property("id");
            if (prop_id != null)
            {
                if (prop_id.Value.Type != JTokenType.Null)
                {
                    ToReturn.Id = Guid.Parse(prop_id.Value.ToString());
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

            //Name
            JProperty prop_name = jo.Property("name");
            if (prop_name != null)
            {
                if (prop_name.Value.Type != JTokenType.Null)
                {
                    ToReturn.Name = prop_name.Value.ToString();
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

            //Display Name
            JProperty prop_displayName = jo.Property("displayName");
            if (prop_displayName != null)
            {
                if (prop_displayName.Value.Type != JTokenType.Null)
                {
                    ToReturn.DisplayName = prop_displayName.Value.ToString();
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

            //List type
            JObject jo_list = JObject.Parse(jo.Property("list").Value.ToString());
            if (jo_list != null)
            {
                JProperty prop_template = jo_list.Property("template");
                if (prop_template != null)
                {
                    if (prop_template.Value.Type != JTokenType.Null)
                    {
                        string tt = prop_template.Value.ToString();
                        if (tt == "documentLibrary")
                        {
                            ToReturn.ListType = SharepointListType.DocumentLibrary;
                        }
                        else if (tt == "genericList")
                        {
                            ToReturn.ListType = SharepointListType.GenericList;
                        }
                        else
                        {
                            ToReturn.ListType = SharepointListType.Other;
                        }
                    }
                }
            }

            return ToReturn;
        }
    }
}