using System.DirectoryServices.AccountManagement;

namespace OAMPS.Office.Word.Models.ActiveDirectory
{
    public class UserPrincipalExSearchFilter : AdvancedFilters
    {
        public UserPrincipalExSearchFilter(Principal p) : base(p)
        {
        }

        public void LogonCount(int value, MatchType mt)
        {
            AdvancedFilterSet("LogonCount", value, typeof (int), mt);
        }
    }
}