using System.Collections.Generic;

namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface ISharePointPictureLibrary : IBaseView
    {
        string Title { get; }
        string FullUrl { get; }
        List<ISharePointPictureLibraryItem> ListItems { get; }
    }
}