# Microsoft Graph Helper
This is a .NET class library designed to assist with the Microsoft Graph Authentication process as well as with transacting with several graph modules. This library is available on NuGet as `TimHanewich.MicrosoftGraphHelper`.

### To Install
```
dotnet add package TimHanewich.MicrosoftGraphHelper
```

## Step 1: User Provides Consent
The core class in this library is the `MicrosoftGraphHelper` class. After creating a new instance of `MicrosoftGraphHelper`, there are several inputs you make that will be used in the authentication process.
- Tenant ID
- Client ID
- Scope
- Redirect URL
After inputting these you can then assemble the URL that users will need to navigate to in order to provide consent. Use the `AssembleAuthorizationUrl` method to assemble this consent URL.

## Step 2: Use Authorization code to gain Access
After the user provides consent, they will be redirected to the redirect URL you specified in the request (and registered app in Azure). You will find the `code` parameter attached the URL. This `code` parameter is what is used to gain an access token. And the access token is what you use to transact with the graph API service. 
To convert this `code` variable to an access token, provide this to the `GetAccessTokenAsync` method. Your access token will be stored in the `MicrosoftGraphTokenPayload` class as the `LastReceivedTokenPayload` property.

## Example Code
```
MicrosoftGraphHelper mgh = new MicrosoftGraphHelper();
mgh.TenantId = Guid.Parse("1e85f23f-c0af-4bce-bb96-92014d3c1359");
mgh.ClientId = Guid.Parse("d9571adf-0c99-4285-bd6c-85d1ad9df015");
mgh.RedirectUrl = "https://www.google.com/";
mgh.Scope.Add("Sites.ReadWrite.All");

//authorization via the web browser. Redirect the user to visit the url and provide consent.
//they will redirected to the redirect URL (must be a registered redirect URL in the application in Azure AD) with a "code" parameter.
string url = mgh.AssembleAuthorizationUrl();
Console.WriteLine(url);
Console.Write("Give me the code: ");
string code = Console.ReadLine();
mgh.GetAccessTokenAsync(code).Wait();

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