using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TimHanewich.MicrosoftGraphHelper
{
    public class AzureAdUser
    {
        public string Displayname {get; set;}
        public Guid? Id {get; set;}
        public string Email {get; set;}

        public AzureAdUser()
        {
            Id = null;
        }

        public static AzureAdUser ParseFromJsonPayload(string payload)
        {
            AzureAdUser ToReturn = new AzureAdUser();

            JObject jo = JObject.Parse(payload);

            //Display Name
            JProperty prop_displayName = jo.Property("displayName");
            if (prop_displayName != null)
            {
                if (prop_displayName.Value.Type != JTokenType.Null)
                {
                    ToReturn.Displayname = prop_displayName.Value.ToString();
                }
            }

            //email
            JProperty prop_email = jo.Property("email");
            if (prop_email != null)
            {
                if (prop_email.Value.Type != JTokenType.Null)
                {
                    ToReturn.Email = prop_email.Value.ToString();
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

            return ToReturn;
        }
    }
}