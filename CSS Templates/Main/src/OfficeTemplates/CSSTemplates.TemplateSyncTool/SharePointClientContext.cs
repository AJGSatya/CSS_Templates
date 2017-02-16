using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Net;
using System.Security;

namespace CSSTemplates.TemplateSyncTool
{
    public class SharePointClientContext : ClientContext
    {
        public SharePointClientContext(string webFullUrl) : base(webFullUrl)
        {
            var pw = new SecureString();
            foreach (char c in ConfigurationManager.AppSettings["Credentials:Password"].ToCharArray()) pw.AppendChar(c);

            //this.AuthenticationMode = ClientAuthenticationMode.Default;
            this.Credentials = new SharePointOnlineCredentials(ConfigurationManager.AppSettings["Credentials:Username"], pw);
        }

    }
}