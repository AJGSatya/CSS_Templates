using System;
using System.Security;
using Microsoft.SharePoint.Client;

namespace OAMPS.Office.Word.Models.SharePoint
{
    public class SharePointUserDirectClientContext : ClientContext
    {
        public SharePointUserDirectClientContext(string webFullUrl) : base(webFullUrl)
        {
            Credentials = GetCredentials();
        }

        public SharePointUserDirectClientContext(Uri webFullUrl) : base(webFullUrl)
        {
            Credentials = GetCredentials();
        }

        private SharePointOnlineCredentials GetCredentials()
        {
            var password = new SecureString();
            foreach (var c in "C$$Templates") password.AppendChar(c);
            return new SharePointOnlineCredentials("CSSTemplatesSVC@ajgau.onmicrosoft.com", password);
        }
    }
}