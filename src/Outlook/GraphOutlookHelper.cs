using System;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.Text;

namespace TimHanewich.MicrosoftGraphHelper.Outlook
{
    public static class GraphOutlookHelper
    {
        public static async Task SendOutlookEmailMessageAsync(this MicrosoftGraphHelper mgh, OutlookEmailMessage msg)
        {
            await mgh.RefreshAccessTokenIfExpiredAsync();

            //Make the request
            HttpRequestMessage reqmsg = new HttpRequestMessage();
            reqmsg.Method = HttpMethod.Post;
            reqmsg.RequestUri = new Uri("https://graph.microsoft.com/v1.0/me/sendMail");
            reqmsg.Headers.Add("Authorization", "Bearer " + mgh.LastReceivedTokenPayload.AccessToken);
            reqmsg.Content = new StringContent(msg.ToPayload(), Encoding.UTF8, "application/json");
            
            //Make the call
            HttpClient hc = new HttpClient();
            HttpResponseMessage hrm = await hc.SendAsync(reqmsg);

            if (hrm.StatusCode != HttpStatusCode.Accepted)
            {
                string errcontent = await hrm.Content.ReadAsStringAsync();
                throw new Exception("Response from graph server was \"" + hrm.StatusCode.ToString() + "\". Response body: " + errcontent);
            }
        }

        //Creates a new outlook event (appointment) on the user's default calendar
        public static async Task CreateOutlookEventAsync(this MicrosoftGraphHelper mgh, OutlookEvent ev)
        {
            await mgh.RefreshAccessTokenIfExpiredAsync();

            //Make the request
            HttpRequestMessage req = mgh.PrepareHttpRequestMessage(); //Adds the bearer token
            req.RequestUri = new Uri("https://graph.microsoft.com/v1.0/me/calendar/events");
            req.Method = HttpMethod.Post;
            req.Content = new StringContent(ev.ToPayload().ToString(), Encoding.UTF8, "application/json");

            //Make the call
            HttpClient hc = new HttpClient();
            HttpResponseMessage hrm = await hc.SendAsync(req);
            if (hrm.StatusCode != HttpStatusCode.Created)
            {
                string errcontent = await hrm.Content.ReadAsStringAsync();
                throw new Exception("Response from graph API when trying to create outlook event was \"" + hrm.StatusCode.ToString() + "\". Response body: " + errcontent);
            }
        }

    }
}