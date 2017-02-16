using System.Collections.Generic;

namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface ISharePointList : IBaseView
    {
        string Title { get; }
        string FullUrl { get; }
        List<ISharePointListItem> ListItems { get; }
        List<ISharePointListItem> HelpListItems { get; }
        ISharePointListItem GetListItemByTitle(string title);
        ISharePointListItem GetListItemById(string id);
        ISharePointListItem GetListItemByLookupField(string fieldName, string value);
        ISharePointListItem GetListItemByTextField(string fieldName, string value);
        //test this
        void AddItem(Dictionary<string, string> fieldValues);
        void UpdateField(string itemId, string fieldName, string value);
    }
}