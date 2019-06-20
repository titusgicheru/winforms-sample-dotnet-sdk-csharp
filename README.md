# winforms-sample-dotnet-sdk-csharp
Winforms Sample to demonstrate how to use the Microsoft Graph SDK, and the Microsoft Authentication Library (MSAL).

## Register the application 
 
1. Navigate to the [the Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) to register your app. Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**. 
 
2. Select **New registration**. On the **Register an application** page, set the values as follows.
  - Set **Name** to **winforms-sample-dotnet-sdk-csharp**. 
  - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**. 
  - Leave **Redirect URI** empty. 
  - Choose **Register**. 
 
3. On the **winforms-sample-dotnet-sdk-csharp** page, copy the value of **Application (client) ID**, since you will need it later. 
 
4. Select the **Add a Redirect URI** link. On the **Redirect URIs** page, locate the **Suggested Redirect URIs for public clients (mobile, desktop)** section. Select the URI that begins with `msal` **and** the **urn:ietf:wg:oauth:2.0:oob** URI. 
 
5. Open the sample solution in Visual Studio and then open the **App.config** file. Change the **clientId** string to the **Application (client) ID** value. 

## Resources

[https://developer.microsoft.com/en-us/graph/get-started](https://developer.microsoft.com/en-us/graph/get-started)

[https://docs.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=CS](https://docs.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=CS)

[https://docs.microsoft.com/en-us/graph/query-parameters](https://docs.microsoft.com/en-us/graph/query-parameters)

[https://docs.microsoft.com/en-us/graph/sdks/sdk-installation](https://docs.microsoft.com/en-us/graph/sdks/sdk-installation)

