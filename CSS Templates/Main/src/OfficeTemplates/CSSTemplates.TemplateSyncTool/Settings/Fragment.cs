using System.Configuration;

namespace CSSTemplates.TemplateSyncTool.Settings
{
    public class Fragment : ConfigurationElement
    {
        [ConfigurationProperty("source", IsRequired = true)]
        public string Source => this["source"] as string;
        [ConfigurationProperty("destination", IsRequired = true)]
        public string Destination => this["destination"] as string;
        [ConfigurationProperty("type", IsRequired = true)]
        public string Type => this["type"] as string;
    }
}