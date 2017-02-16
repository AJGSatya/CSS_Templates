
using System.Collections.Generic;

namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface ISharePointList : IBaseView
    {
        string Title { get; }
        string FullUrl { get; }
        List<ISharePointListItem> ListItems { get; }
        List<ISharePointListItem> HelpListItems { get; }//test this
    }
}
