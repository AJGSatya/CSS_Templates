using System;
using System.DirectoryServices.AccountManagement;
using System.Globalization;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Properties;
using OAMPS.Office.BusinessLogic.Loggers;

namespace OAMPS.Office.Word.Models.SharePoint
{
    public class SharePointAdalClientContext : ClientContext
    {
        private const string AadInstance = "https://login.microsoftonline.com/{0}";
        private static readonly string Tenant = Settings.Default.tenant;
        private static readonly string ClientId = Settings.Default.clientid;
        private static readonly string Authority = string.Format(CultureInfo.InvariantCulture, AadInstance, Tenant);
        private static readonly Uri RedirectUri = new Uri(Settings.Default.redirecturi);
        private static AuthenticationContext _authContext;
        private static AuthenticationResult _result;
        private static readonly FileCache FileCache = new FileCache();
        private static readonly object Lock = new object();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public SharePointAdalClientContext(string webFullUrl) : base(webFullUrl)
        {
            ExecutingWebRequest += SharePointClientContext_ExecutingWebRequest;
        }

        public SharePointAdalClientContext(Uri webFullUrl) : base(webFullUrl)
        {
            ExecutingWebRequest += SharePointClientContext_ExecutingWebRequest;
        }

        private static string GetADALAttribute()
        {
            var logger = new EventViewerLogger();
            if (UserPrincipal.Current.EmailAddress.IndexOf("@ajg.com.au") > 0) {
                logger.Log(string.Format("Current user's e-mail {0} address includes @ajg.com.au, no need to append the domain for authenticating to O365", UserPrincipal.Current.EmailAddress), BusinessLogic.Interfaces.Logging.Type.Debug);
                return UserPrincipal.Current.EmailAddress;
            }
            logger.Log(string.Format("Current user's e-mail {0} address doesn't include @ajg.com.au, append the domain to the account for authenticating to O365", UserPrincipal.Current.EmailAddress), BusinessLogic.Interfaces.Logging.Type.Debug);
            return UserPrincipal.Current.SamAccountName + "@ajg.com.au";
        }

        internal AuthenticationResult GetAuthorizationBearerTokenFromAdfs()
        {
            try
            {
                var logger = new EventViewerLogger();
                logger.Log(string.Format("Getting Authorization Bearer Token from ADFS"), BusinessLogic.Interfaces.Logging.Type.Debug);
                _authContext = new AuthenticationContext(Authority, FileCache);
                var r =
                    _authContext.AcquireToken(Settings.Default.TokenResourceUrl, ClientId, RedirectUri,
                        PromptBehavior.Auto, new UserIdentifier(GetADALAttribute(), UserIdentifierType.OptionalDisplayableId));
                return r;
            }
            catch (Exception ex)
            {
                var logger = new EventViewerLogger();
                logger.Log(string.Format("Exception occured Getting Authorization Bearer Token from ADFS: {0}", ex.Message), BusinessLogic.Interfaces.Logging.Type.Error);
                return null;
            }
        }

        private AuthenticationResult GetAuthorizationBearerTokenFromPrompt()
        {
            var logger = new EventViewerLogger();
            logger.Log(string.Format("Getting Authorization Bearer Token from prompt"), BusinessLogic.Interfaces.Logging.Type.Debug);

            //var activeDirectoryClient = AuthenticationHelper.GetActiveDirectoryClientAsApplication();
            var ac = new AuthenticationContext(Settings.Default.TenantMetadataUrl);
            var arr = ac.AcquireToken(Settings.Default.TokenResourceUrl, ClientId, RedirectUri,PromptBehavior.Auto);
            return arr;
        }

        private void SharePointClientContext_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            lock (Lock)
            {
                var logger = new EventViewerLogger();
                logger.Log(string.Format("Executing Web Request to obtain SharePoint Client Context: {0}", e.WebRequestExecutor.WebRequest.RequestUri), BusinessLogic.Interfaces.Logging.Type.Debug);
                if (_result == null)
                    _result = GetAuthorizationBearerTokenFromAdfs() ?? GetAuthorizationBearerTokenFromPrompt();
                if (_result == null) return;
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + _result.AccessToken;
            }
        }
    }
}