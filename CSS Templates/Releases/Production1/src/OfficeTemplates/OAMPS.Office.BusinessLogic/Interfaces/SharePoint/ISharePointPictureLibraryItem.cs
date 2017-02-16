
using System.IO;

namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface ISharePointPictureLibraryItem : ISharePointListItem
    {
        Stream ThumbnailStream { get; }
        string ThumbnailUrl { get;  }
        int Order { get; }
//        string HeaderType { get; }

     
    }
}
