namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface ISharePointPictureLibraryItem : ISharePointListItem
    {
        byte[] ThumbnailStream { get; }
        string ThumbnailUrl { get; }
        int Order { get; }
//        string HeaderType { get; }
    }
}