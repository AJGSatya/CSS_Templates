using System;
using System.Linq;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

namespace OAMPS.Office.BusinessLogic.Helpers
{
    public static class ListQueries
    {
        private static bool MatchLookupField(this ISharePointListItem item, string field, string value) => item.GetLookupFieldValue(field).Equals(value, StringComparison.InvariantCultureIgnoreCase);
        private static bool MatchField(this ISharePointListItem item, string field, string value) => item.GetFieldValue(field).Equals(value, StringComparison.InvariantCultureIgnoreCase);

        public static Func<ISharePointListItem, bool> HelpGetItemQuery(string template, string title)
        {
            return item => item.MatchField("Template", template) && item.MatchField("Title", title);
        }

        public static Func<ISharePointListItem, bool> GetItemByTitleQuery(string title)
        {
            return item => item.MatchField("Title", title);
        }

        public static Func<ISharePointListItem, bool> GetItemByQuestionaresLookupId(string value)
        {
            return item => item.MatchLookupField("Questionares", value);
        }

        public static Func<ISharePointListItem, bool> GetItemByPolicyTypeQuery(string value)
        {
            return item => item.MatchLookupField("Policy_x0020_Type", value);
        }

        public static Func<ISharePointListItem, bool> FactFinderFragmentsByKey()
        {
            return item => item.MatchField("Key", "General Advice Warning") || item.MatchField("Key", "Uninsured Risks Checklist") || item.MatchField("Key", "Privacy Statement") ||
                           item.MatchField("Key", "Important Notices") || item.MatchField("Key", "Financial Services Guide");
        }

        public static Func<ISharePointListItem, bool> RenewalLetterFragmentsByKey()
        {
            return item => item.MatchField("Key", "General Advice Warning") || item.MatchField("Key", "Uninsured Risks Checklist") || item.MatchField("Key", "Privacy Statement") ||
                           item.MatchField("Key", "Important Notices") || item.MatchField("Key", "Financial Services Guide letter");
        }

        public static Func<ISharePointListItem, bool> FragmentsByKeys(params string[] keys)
        {
            return item => keys.Any(key => item.MatchField("Key", key));
        } 
    }
}
