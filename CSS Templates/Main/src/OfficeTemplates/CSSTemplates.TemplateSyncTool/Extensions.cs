using System.Collections.Specialized;

namespace CSSTemplates.TemplateSyncTool
{
    public static class Extensions
    {
        public static bool AsBoolean(this NameValueCollection settings, string key)
        {
            var val = settings[key];
            if (string.IsNullOrEmpty(val))
                return false;

            bool res;
            var success = bool.TryParse(val, out res);
            return (success && res);
        }
    }
}