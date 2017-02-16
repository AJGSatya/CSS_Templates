using System.Configuration;

namespace CSSTemplates.TemplateSyncTool.Settings
{
    public class Fragments : ConfigurationElementCollection
    {
        public Fragment this[int index]
        {
            get
            {
                return BaseGet(index) as Fragment;
            }
            set
            {
                if (BaseGet(index) != null)
                {
                    BaseRemoveAt(index);
                }
                BaseAdd(index, value);
            }
        }

        public new Fragment this[string responseString]
        {
            get
            {
                return (Fragment)BaseGet(responseString);
            }
            set
            {
                if (BaseGet(responseString) != null)
                {
                    BaseRemoveAt(BaseIndexOf(BaseGet(responseString)));
                }
                BaseAdd(value);
            }
        }

        protected override ConfigurationElement CreateNewElement()
        {
            return new Fragment();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((Fragment)element).Source;
        }
    }
}