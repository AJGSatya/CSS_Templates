using System.Collections.Generic;
using System.Linq;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters.SharePoint
{
    public class SharePointPictureLibraryPresenter
    {
        private readonly ISharePointPictureLibrary _list;
        private IBaseView _view;

        public SharePointPictureLibraryPresenter(ISharePointPictureLibrary list, IBaseView view)
        {
            _list = list;
            _view = view; //this is used to communicate back to the view.
        }

        public List<IThumbnail> GetPictureLibraryCoverPageThumnails()
        {
            List<ISharePointPictureLibraryItem> items = _list.ListItems;
            return items.Select(item => new Thumbnail
                {
                    ImageStream = item.ThumbnailStream,
                    ImageTitle = item.Title,
                    Order = item.Order,
                    FullImageUrl = item.FullUrl,
                    HeaderType = item.GetFieldValue(Constants.SharePointFields.HeaderType),
                    LongDescription = item.GetFieldValue(Constants.SharePointFields.LongDescription),
                    ShortDescription = item.GetFieldValue(Constants.SharePointFields.ShortDescription),
                    RelativeUrl = item.GetFieldValue(Constants.SharePointFields.FileRef)
                }).Cast<IThumbnail>().ToList();
        }

        public List<IThumbnail> GetPictureLibraryLogoThumbnails()
        {
            List<ISharePointPictureLibraryItem> items = _list.ListItems;
            return items.Select(item => new Thumbnail
                {
                    ImageStream = item.ThumbnailStream,
                    ImageTitle = item.Title,
                    Order = item.Order,
                    FullImageUrl = item.FullUrl,
                    ABN = item.GetFieldValue(Constants.SharePointFields.ABN),
                    AFSL = item.GetFieldValue(Constants.SharePointFields.AFSL),
                    WebSite = item.GetFieldValue(Constants.SharePointFields.Website)
                }).Cast<IThumbnail>().ToList();
        }
    }
}