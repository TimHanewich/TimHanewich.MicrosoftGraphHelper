# Microsoft Graph Helper
This is a .NET class library designed to assist with the Microsoft Graph Authentication process as well as with transacting with several graph modules. This library is available on NuGet as `TimHanewich.MicrosoftGraphHelper`.

## Step 1: User Provides Consent
The core class in this library is the `MicrosoftGraphHelper` class. After creating a new instance of `MicrosoftGraphHelper`, there are several inputs you make that will be used in the authentication process.
- Tenant ID
- Client ID
- Scope
- Redirect URL
After inputting these you can then assemble the URL that users will need to navigate to in order to provide consent. Use the `AssembleAuthorizationUrl` method to assemble this consent URL.

# Step 2: Use Authorization code to gain Access
After the user provides consent, they will be redirected to the redirect URL you specified in the request (and registered app in Azure). You will find the `code` parameter attached the URL. This `code` parameter is what is used to gain an access token. And the access token is what you use to transact with the graph API service. 
To convert this `code` variable to an access token, provide this to the `GetAccessTokenAsync` method. Your access token will be stored in the `MicrosoftGraphTokenPayload` class as the `LastReceivedTokenPayload` property.