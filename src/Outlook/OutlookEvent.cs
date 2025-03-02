using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TimHanewich.MicrosoftGraphHelper.Outlook
{
    public class OutlookEvent
    {
        public string Subject {get; set;}
        public string Body {get; set;}
        public DateTime StartUTC {get; set;} // the start time of the event, in UTC time
        public DateTime EndUTC {get; set;} // the end time of the event, in UTC time

        public JObject ToPayload()
        {
            JObject ToReturn = new JObject();

            ToReturn.Add("subject", Subject);
            
            //Body
            JObject body = new JObject();
            ToReturn.Add("body", body);
            body.Add("contentType", "text");
            body.Add("content", Body);

            //start
            JObject start = new JObject();
            ToReturn.Add("start", start);
            start.Add("dateTime", StartUTC.ToString());
            start.Add("timeZone", "UTC");

            //end
            JObject end = new JObject();
            ToReturn.Add("end", end);
            end.Add("dateTime", EndUTC.ToString());
            end.Add("timeZone", "UTC");

            return ToReturn;
        }
    }
}