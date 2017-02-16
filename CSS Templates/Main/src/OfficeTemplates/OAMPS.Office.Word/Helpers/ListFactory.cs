using System;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.Word.Models.Local;
using OAMPS.Office.Word.Models.SharePoint;

namespace OAMPS.Office.Word.Helpers
{
    public class ListFactory
    {
        //public static ISharePointList Create(string contextUrl, string listTitle, string itemsCamlQuery = "<View/>")
        //{
        //    return new SharePointList(contextUrl, listTitle, itemsCamlQuery);
        //    //return new LocalList(contextUrl, listTitle, itemsCamlQuery);
        //}

        public static ISharePointList Create(string listTitle, bool sortBySortOrder = false)
        {
            return new LocalList(listTitle, (item => true), sortBySortOrder);
        }

        public static ISharePointList Create(string listTitle, Func<ISharePointListItem, bool> predicate, bool sortBySortOrder = false)
        {
            return new LocalList(listTitle, predicate);
        }
    }
}
