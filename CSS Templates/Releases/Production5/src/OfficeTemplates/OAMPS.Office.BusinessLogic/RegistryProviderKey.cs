namespace OAMPS.Office.BusinessLogic
{
    public class RegistryProviderKey
    {
        public RegistryProviderKey(string keyName, string serviceUrl, string siteUrl)
        {
            KeyName = keyName;
            ServiceUrl = serviceUrl;
            SiteUrl = siteUrl;
        }

        public string KeyName { get; set; }
        public string SiteUrl { get; set; }
        public string ServiceUrl { get; set; }
    }
}