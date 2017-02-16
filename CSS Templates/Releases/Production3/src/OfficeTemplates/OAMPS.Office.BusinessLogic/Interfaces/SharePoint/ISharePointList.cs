
using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface ISharePointList : IBaseView
    {
        string Title { get; }
        string FullUrl { get; }
        ISharePointListItem GetListItemByTitle(string title);
        List<ISharePointListItem> ListItems { get; }        
        List<ISharePointListItem> HelpListItems { get; }//test this
        void AddItem(Dictionary<string, string> fieldValues);
    }
}
