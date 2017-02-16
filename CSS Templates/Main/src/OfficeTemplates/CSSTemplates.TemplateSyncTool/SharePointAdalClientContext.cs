//using System;
//using System.Configuration;
//using System.DirectoryServices.AccountManagement;
//using System.Globalization;
//using CSSTemplates.TemplateSyncTool.Properties;
//using Microsoft.IdentityModel.Clients.ActiveDirectory;
//using Microsoft.SharePoint.Client;

//namespace CSSTemplates.TemplateSyncTool
//{
//    public class SharePointAdalClientContext : ClientContext
//    {
//        private const string AadInstance = "https://login.microsoftonline.com/{0}";
//        private static readonly string Tenant = ConfigurationManager.AppSettings["tenant"];
//        private static readonly string ClientId = ConfigurationManager.AppSettings["clientid"];
//        private static readonly string ClientSecret = ConfigurationManager.AppSettings["clientsecret"];
//        private static readonly string Authority = string.Format(CultureInfo.InvariantCulture, AadInstance, Tenant);
//        private static readonly Uri RedirectUri = new Uri(ConfigurationManager.AppSettings["redirecturi"]);
//        private static AuthenticationContext _authContext;
//        private static AuthenticationResult _result;
//        private static readonly FileCache FileCache = new FileCache();
//        private static readonly object Lock = new object();

//        public SharePointAdalClientContext(string webFullUrl) : base(webFullUrl)
//        {
//            ExecutingWebRequest += SharePointClientContext_ExecutingWebRequest;
//        }

//        public SharePointAdalClientContext(Uri webFullUrl) : base(webFullUrl)
//        {
//            ExecutingWebRequest += SharePointClientContext_ExecutingWebRequest;
//        }
//        private static string GetADALAttribute()
//        {
//            if (UserPrincipal.Current.EmailAddress.IndexOf("@ajg.com.au") > 0)
//                return UserPrincipal.Current.EmailAddress;
//            else
//                return UserPrincipal.Current.SamAccountName + "@ajg.com.au";
//        }
//        internal AuthenticationResult GetAuthorizatoinBearerTokenFromAdfs()
//        {
//            try
//            {
//                _authContext = new AuthenticationContext(Authority, FileCache);
//                var r = _authContext.AcquireToken(ConfigurationManager.AppSettings["TokenResourceUrl"], ClientId, RedirectUri, PromptBehavior.Auto,
//                    new UserIdentifier(GetADALAttribute(), UserIdentifierType.OptionalDisplayableId));
//                return r;
//            }
//            catch (Exception ex)
//            {
//                return null;
//            }
//        }

//        private AuthenticationResult GetAuthorizationBearerTokenFromPrompt()
//        {
//            //var activeDirectoryClient = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
//            var ac = new AuthenticationContext(ConfigurationManager.AppSettings["TenantMetadataUrl"]);
//            var arr = ac.AcquireToken(ConfigurationManager.AppSettings["TokenResourceUrl"], ClientId, RedirectUri, PromptBehavior.RefreshSession);
//            return arr;
//        }

//        private AuthenticationResult GetAuthorizationBearerTokenFromSecret()
//        {
//            var ac = new AuthenticationContext(ConfigurationManager.AppSettings["TenantMetadataUrl"]);
//            var cred = new ClientCredential(ClientId, ClientSecret);
//            var arr = ac.AcquireToken(ConfigurationManager.AppSettings["TokenResourceUrl"], cred);
//            return arr;
//        }

//        private void SharePointClientContext_ExecutingWebRequest(object sender, WebRequestEventArgs e)
//        {
//            lock (Lock)
//            {
//                if (_result == null)
//                {
//                    if (ConfigurationManager.AppSettings.AsBoolean("PromptForAuth"))
//                    {
//                        _result = GetAuthorizatoinBearerTokenFromAdfs() ?? GetAuthorizationBearerTokenFromPrompt();
//                    }
//                    else
//                    {
//                        _result = GetAuthorizationBearerTokenFromSecret();
//                    }
//                }
                    
//                if (_result == null)
//                    return;

//                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + _result.AccessToken;
//            }
//        }
        
//    }
//}