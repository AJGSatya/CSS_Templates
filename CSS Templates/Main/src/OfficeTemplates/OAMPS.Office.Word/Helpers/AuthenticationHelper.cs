using System;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace OAMPS.Office.Word.Helpers
{
    internal class AuthenticationHelper
    {
        public const string TenantName = "obscloud365.onMicrosoft.com";
        public const string TenantId = "4fd2b2f2-ea27-4fe5-a8f3-7b1a7c975f34";
        public const string ClientId = "5de95ea1-74a7-452f-915e-4000a0940be9";
        public const string ClientSecret = "hOrJ0r0TZ4GQ3obp+vk3FZ7JBVP+TX353kNo6QwNq7Q=";
        public const string ClientIdForUserAuthn = "66133929-66a4-4edc-aaee-13b04b03207d";
        public const string AuthString = "https://login.microsoftonline.com/" + TenantName;
        public const string ResourceUrl = "https://graph.windows.net";
        public static string TokenForUser;


        /// <summary>
        ///     Async task to acquire token for Application.
        /// </summary>
        /// <returns>Async Token for application.</returns>
        public static async Task<string> AcquireTokenAsyncForApplication()
        {
            return GetTokenForApplication();
        }

        /// <summary>
        ///     Get Token for Application.
        /// </summary>
        /// <returns>Token for application.</returns>
        public static string GetTokenForApplication()
        {
            var authenticationContext = new AuthenticationContext(AuthString, false);
            // Config for OAuth client credentials 
            var clientCred = new ClientCredential(ClientId, ClientSecret);
            var authenticationResult = authenticationContext.AcquireTokenAsync(ResourceUrl,
                clientCred).Result;
            var token = authenticationResult.AccessToken;
            return token;
        }

        /// <summary>
        ///     Get Active Directory Client for Application.
        /// </summary>
        /// <returns>ActiveDirectoryClient for Application.</returns>
        public static ActiveDirectoryClient GetActiveDirectoryClientAsApplication()
        {
            var servicePointUri = new Uri(ResourceUrl);
            var serviceRoot = new Uri(servicePointUri, TenantId);
            var activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                async () => await AcquireTokenAsyncForApplication());
            return activeDirectoryClient;
        }

        /// <summary>
        ///     Async task to acquire token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> AcquireTokenAsyncForUser()
        {
            return GetTokenForUser();
        }

        /// <summary>
        ///     Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static string GetTokenForUser()
        {
            if (TokenForUser == null)
            {
                var redirectUri = new Uri("https://localhost");
                var authenticationContext = new AuthenticationContext(AuthString, false);
                var userAuthnResult = authenticationContext.AcquireToken(ResourceUrl,
                    ClientIdForUserAuthn, redirectUri, PromptBehavior.Auto);
                TokenForUser = userAuthnResult.AccessToken;
            }
            return TokenForUser;
        }

        /// <summary>
        ///     Get Active Directory Client for User.
        /// </summary>
        /// <returns>ActiveDirectoryClient for User.</returns>
        public static ActiveDirectoryClient GetActiveDirectoryClientAsUser()
        {
            var servicePointUri = new Uri(ResourceUrl);
            var serviceRoot = new Uri(servicePointUri, TenantId);
            var activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                async () => await AcquireTokenAsyncForUser());
            return activeDirectoryClient;
        }
    }
}