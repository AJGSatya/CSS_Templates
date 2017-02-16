using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;

namespace OAMPS.Office.Word.Models.ActiveDirectory
{
    public class UserPrincipalExSearchFilter : AdvancedFilters
    {
        public UserPrincipalExSearchFilter(Principal p) : base(p) { }

        public void LogonCount(int value, MatchType mt)
        {
            this.AdvancedFilterSet("LogonCount", value, typeof(int), mt);
        }
    }
}
