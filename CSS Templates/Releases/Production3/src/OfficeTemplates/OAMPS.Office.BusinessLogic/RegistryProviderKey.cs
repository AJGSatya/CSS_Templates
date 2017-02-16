using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OAMPS.Office.BusinessLogic
{
    public class RegistryProviderKey
    {
        public string KeyName { get; set; }
        public string SiteUrl { get; set; }
        public string ServiceUrl { get; set; }

        public RegistryProviderKey(string keyName, string serviceUrl, string siteUrl)
        {
            KeyName = keyName;
            ServiceUrl = serviceUrl;
            SiteUrl = siteUrl;
        }
    }
}
