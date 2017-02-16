using System.Configuration;

namespace CSSTemplates.TemplateSyncTool.Settings
{
    public class FragmentSettings : ConfigurationSection
    {

        public static FragmentSettings GetConfig()
        {
            return (FragmentSettings)ConfigurationManager.GetSection("fragmentSettings") ?? new FragmentSettings();
        }

        [ConfigurationProperty("fragments")]
        [ConfigurationCollection(typeof(Fragments), AddItemName = "fragment")]
        public Fragments Fragments => this["fragments"] as Fragments;
    }
}