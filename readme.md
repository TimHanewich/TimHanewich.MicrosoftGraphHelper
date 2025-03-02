# Microsoft Graph Helper
This is a .NET class library designed to assist with the Microsoft Graph Authentication process as well as with transacting with several graph modules. This library is available on NuGet as `TimHanewich.MicrosoftGraphHelper`. 

This library was built around the Microsoft documentation specified [here](https://learn.microsoft.com/en-us/graph/auth-v2-user).

### To Install
```
dotnet add package TimHanewich.MicrosoftGraphHelper
```

## Step 1: User Provides Consent
The core class in this library is the `MicrosoftGraphHelper` class. After creating a new instance of `MicrosoftGraphHelper`, there are several inputs you make that will be used in the authentication process.
- Client ID
- Scope
- Redirect URL
After inputting these you can then assemble the URL that users will need to navigate to in order to provide consent. Use the `AssembleAuthorizationUrl` method to assemble this consent URL.

## Step 2: Use Authorization code to gain Access
After the user provides consent, they will be redirected to the redirect URL you specified in the request (and registered app in Azure). You will find the `code` parameter attached the URL. This `code` parameter is what is used to gain an access token. And the access token is what you use to transact with the graph API service. 
To convert this `code` variable to an access token, provide this to the `GetAccessTokenAsync` method. Your access token will be stored in the `MicrosoftGraphTokenPayload` class as the `LastReceivedTokenPayload` property.

## Example: Authenticating with Microsoft Graph
The following example demonstrates using this library to authenticate with the Microsoft Graph API. 

*Note: The `Tenant` property, per [Microsoft documentation](https://learn.microsoft.com/en-us/graph/auth-v2-user), can be "common" for both Microsoft accounts and work/school accounts, "organizations" for work/school accounts only, "consumers" for Microsoft accounts only, of a tenant identifier (GUID).*

```
MicrosoftGraphHelper mgh = new MicrosoftGraphHelper();
mgh.Tenant = "consumers";
mgh.ClientId = Guid.Parse("d9571adf-0c99-4285-bd6c-85d1ad9df015");
mgh.RedirectUrl = "https://www.google.com/";
mgh.Scope.Add("User.Read");
mgh.Scope.Add("Calendars.ReadWrite");
mgh.Scope.Add("Mail.Read");


//authorization via the web browser. Redirect the user to visit the url and provide consent.
//they will redirected to the redirect URL (must be a registered redirect URL in the application in Azure AD) with a "code" parameter.
string url = mgh.AssembleAuthorizationUrl();
Console.WriteLine("Please go to the following URL and sign in. After you sign in, give me the "code" parameter out of the URL it redirects you to");
Console.WriteLine(url);
Console.Write("Give me the code: ");
string code = Console.ReadLine();
mgh.GetAccessTokenAsync(code).Wait();
```

## Example: Resuming Access After a Period of Inactivity
Normally the bearer token you will be given will expire within 60 minutes. This means your access will also stop. However, if the `offline_access` scope was added to the original authorization flow (it is by default in the `MicrosoftGraphHelper` class), you can refresh your token by using the refresh token that was originally provided in the authorization flow. 

For example:

```
// The token payload was saved to JSON previously. Here, we are retrieving it and adding it back
MicrosoftGraphTokenPayload tokens = JsonConvert.DeserializeObject<MicrosoftGraphTokenPayload>(System.IO.File.ReadAllText(@"C:\Users\timh\Downloads\tah\TimHanewich.MicrosoftGraphHelper\payload.json"));
MicrosoftGraphHelper mgh = new MicrosoftGraphHelper();
mgh.LastReceivedTokenPayload = tokens;

//Refresh if the retrieved token is expired
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
```

## Example: Create Outlook Calendar Event (Appointment)
The following demonstrates how you can schedule a new event in your user's default outlook calendar. It requires the `Calendars.ReadWrite` scope.

```
//Create Outlook Event
OutlookEvent ev = new OutlookEvent();
ev.Subject = "Let's do something";
ev.Body = "Go shopping maybe?";
ev.StartUTC = new DateTime(2025, 03, 02, 12, 0, 0);
ev.EndUTC = ev.StartUTC.AddMinutes(15);

//Schedule
Console.Write("Scheduling... ");
await mgh.CreateOutlookEventAsync(ev);
Console.WriteLine("done!");
```

## Example: Send an Email via Outlook
The following requires the `Mail.Send` scope.

```
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
```

## Example: Sharepoint List Manipulation
```
//Get the sites that are available
SharepointSite[] sites = mgh.SearchSharepointSitesAsync("").Result;
Console.WriteLine(JArray.Parse(JsonConvert.SerializeObject(sites)).ToString());

//Get the lists in that site
SharepointList[] lists = mgh.ListSharepointListsAsync(Guid.Parse("2e069086-c6f2-4735-a728-eb33b8347842")).Result;
Console.WriteLine(JArray.Parse(JsonConvert.SerializeObject(lists)).ToString());

//Get the content of a list
SharepointListItem[] items = mgh.GetAllItemsFromSharepointListAsync(Guid.Parse("2e069086-c6f2-4735-a728-eb33b8347842"), Guid.Parse("771b32f1-859c-4570-8bf2-7c86d140dc5c")).Result;
Console.WriteLine(JArray.Parse(JsonConvert.SerializeObject(items)).ToString());

//Creating a new item (record) in a list
JObject jo = new JObject();
jo.Add("Title", "Harry the Hippo");
mgh.CreateItemAsync(Guid.Parse("2e069086-c6f2-4735-a728-eb33b8347842"), Guid.Parse("771b32f1-859c-4570-8bf2-7c86d140dc5c"), jo).Wait();
```