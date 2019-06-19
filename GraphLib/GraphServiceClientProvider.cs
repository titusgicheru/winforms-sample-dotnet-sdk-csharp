using System;
using System.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace GraphLib
{
    public class GraphServiceClientProvider
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        static string clientId = ConfigurationManager.AppSettings["clientId"].ToString();
        private static string[] scopes = {
            "User.Read",
            "Calendars.ReadWrite",
            "Files.Read.All",
            "Mail.Read",
            "Mail.Send"
        };

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient()
        {
            if (graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    IPublicClientApplication clientApplication = InteractiveAuthenticationProvider.CreateClientApplication(clientId);
                    InteractiveAuthenticationProvider authProvider = new InteractiveAuthenticationProvider(clientApplication, scopes);

                    graphClient = new GraphServiceClient(authProvider);
                    return graphClient;
                }

                catch (Exception ex)
                {
                    throw new Exception("Could not create a graph client: " + ex.Message);
                }
            }

            return graphClient;
        }
    }
}
