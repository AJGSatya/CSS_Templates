
namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface ISharePointListItem : IBaseView
    {
        string Title { get; }
        string FileRef { get; }
        string FullUrl { get; }
        
        string GetFieldValue(string fieldName);
        string GetLookupFieldValue(string fieldName);
    }
}
