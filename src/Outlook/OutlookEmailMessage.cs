using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace TimHanewich.MicrosoftGraphHelper.Outlook
{
    public class OutlookEmailMessage
    {
        //Example body to post to the endpoint:
        // {
        //     "message": {
        //         "subject": "Meet for lunch?",
        //         "body": {
        //         "contentType": "Text",
        //         "content": "The new cafeteria is open."
        //         },
        //         "toRecipients": [
        //         {
        //             "emailAddress": {
        //             "address": "meganb@contoso.onmicrosoft.com"
        //             }
        //         }
        //         ],
        //         "attachments": [
        //         {
        //             "@odata.type": "#microsoft.graph.fileAttachment",
        //             "name": "attachment.txt",
        //             "contentType": "text/plain",
        //             "contentBytes": "SGVsbG8gV29ybGQh"
        //         }
        //         ]
        //     }
        // }

        public string Subject {get; set;}
        public OutlookEmailMessageContentType ContentType {get; set;}
        public string Content {get; set;}
        public List<string> ToRecipients {get; set;}

        public OutlookEmailMessage()
        {
            ToRecipients = new List<string>();
        }

        public string ToPayload()
        {
            JObject ToReturn = new JObject();

            //Inside message payload
            JObject jo_message = new JObject();
            ToReturn.Add("message", jo_message);
            if (Subject != null)
            {
                jo_message.Add("subject", Subject);
            }

            //body
            JObject jo_body = new JObject();
            jo_message.Add("body", jo_body);
            string contentTypeString = "";
            if (ContentType == OutlookEmailMessageContentType.Text)
            {
                contentTypeString = "Text";
            }
            else if (ContentType == OutlookEmailMessageContentType.HTML)
            {
                contentTypeString = "HTML";
            }
            jo_body.Add("contentType", contentTypeString);
            jo_body.Add("content", Content);

            //toRecipients
            List<JObject> ToRecipientsObjects = new List<JObject>();
            foreach (string s in ToRecipients)
            {
                JObject thisToRecip = new JObject();
                
                JObject thisToRecip_emailAddress = new JObject();
                thisToRecip.Add("emailAddress", thisToRecip_emailAddress);
                thisToRecip_emailAddress.Add("address", s);

                ToRecipientsObjects.Add(thisToRecip);
            }
            JArray arrayOfToRecips = JArray.Parse(JsonConvert.SerializeObject(ToRecipientsObjects.ToArray()));
            jo_message.Add("toRecipients", arrayOfToRecips);

            return ToReturn.ToString();
        }
    }
}