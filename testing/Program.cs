using System;
using System.Threading.Tasks;
using Newtonsoft.Json;
using TimHanewich.MicrosoftGraphHelper;
using TimHanewich.MicrosoftGraphHelper.Outlook;

namespace testing
{
    class Program
    {
        static void Main(string[] args)
        {
            //Authenticate().Wait();
            DoSomething().Wait();
        }

        public static async Task Authenticate()
        {
            MicrosoftGraphHelper mgh = new MicrosoftGraphHelper();
            mgh.Tenant = "consumers";
            mgh.ClientId = Guid.Parse("e32b77a3-67df-411b-927b-f05cc6fe8d5d");
            mgh.RedirectUrl = "https://www.google.com/";
            mgh.Scope.Add("User.Read");
            mgh.Scope.Add("Calendars.ReadWrite");
            mgh.Scope.Add("Mail.Read");

            string url = mgh.AssembleAuthorizationUrl();
            Console.WriteLine(url);

            Console.WriteLine("Go to the above URL, sign in, and then give me the code.");
            Console.Write("> ");
            string code = Console.ReadLine();

            Console.WriteLine("Redeeming code...");
            await mgh.GetAccessTokenAsync(code);
            Console.WriteLine("Redeemed!");

            System.IO.File.WriteAllText(@"C:\Users\timh\Downloads\tah\TimHanewich.MicrosoftGraphHelper\payload.json", JsonConvert.SerializeObject(mgh.LastReceivedTokenPayload, Formatting.Indented));
            Console.WriteLine("Wrote");
        }

        public static async Task DoSomething()
        {
            MicrosoftGraphTokenPayload tokens = JsonConvert.DeserializeObject<MicrosoftGraphTokenPayload>(System.IO.File.ReadAllText(@"C:\Users\timh\Downloads\tah\TimHanewich.MicrosoftGraphHelper\payload.json"));
            MicrosoftGraphHelper mgh = new MicrosoftGraphHelper();
            mgh.LastReceivedTokenPayload = tokens;
            
            if (mgh.AccessTokenHasExpired())
            {
                Console.Write("Tokens are expired! Refreshing... ");
                await mgh.RefreshAccessTokenAsync(); 
                Console.WriteLine("Refreshed!");  
            }
            else
            {
                Console.WriteLine("Tokens are still active! No need to refresh.");
            }

            //Construct email
            OutlookEmailMessage email = new OutlookEmailMessage();
            email.ToRecipients.Add("timhanewich@gmail.com");
            email.Subject = "My favorite songs";
            email.Content = "1. Chris Brown - Yeah 3X\n2. Chris Brown - Forever\n3. Chris Brown - Turn Up the Music";
            email.ContentType = OutlookEmailMessageContentType.Text;
            
            //Send email
            Console.Write("Sending email... ");
            await mgh.SendOutlookEmailMessageAsync(email);
            Console.WriteLine("Sent!");

        }

    }
}
