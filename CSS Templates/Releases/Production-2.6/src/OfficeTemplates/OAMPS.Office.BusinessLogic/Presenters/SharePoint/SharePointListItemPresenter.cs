using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

namespace OAMPS.Office.BusinessLogic.Presenters.SharePoint
{
    public class SharePointListItemPresenter
    {
        protected ISharePointListItem ListItem;

        public SharePointListItemPresenter(ISharePointListItem listItem)
        {
            ListItem = listItem;
        }
    }
}