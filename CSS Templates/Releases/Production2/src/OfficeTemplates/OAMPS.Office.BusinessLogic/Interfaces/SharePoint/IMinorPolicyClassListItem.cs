
using System.Collections.Generic;

namespace OAMPS.Office.BusinessLogic.Interfaces.SharePoint
{
    public interface IMinorPolicyClassListItem : ISharePointListItem
    {
        string MajorClassTitle { get; }
    }
}
