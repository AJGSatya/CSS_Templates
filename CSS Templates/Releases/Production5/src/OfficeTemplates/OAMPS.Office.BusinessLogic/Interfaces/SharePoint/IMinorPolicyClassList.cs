using System.Collections.Generic;

namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface IMinorPolicyClassList : IBaseView
    {
        string Title { get; }
        string FullUrl { get; }
        List<IMinorPolicyClassListItem> ListItems { get; }
    }
}