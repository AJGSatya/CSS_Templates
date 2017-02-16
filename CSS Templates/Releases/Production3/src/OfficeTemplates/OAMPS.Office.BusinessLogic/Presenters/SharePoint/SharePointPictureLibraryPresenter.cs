
using System.Collections.Generic;
using System.Linq;
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
            var items = _list.ListItems;
            return items.Select(item => new Thumbnail
                {
                    ImageStream = item.ThumbnailStream,
                    ImageTitle = item.Title,
                    Order = item.Order,
                    FullImageUrl = item.FullUrl,
                    HeaderType = item.GetFieldValue(Helpers.Constants.SharePointFields.HeaderType),
                    LongDescription = item.GetFieldValue(Helpers.Constants.SharePointFields.LongDescription),
                    ShortDescription = item.GetFieldValue(Helpers.Constants.SharePointFields.ShortDescription)
                   
                }).Cast<IThumbnail>().ToList();
        }

        public List<IThumbnail> GetPictureLibraryLogoThumbnails()
        {
            var items = _list.ListItems;
            return items.Select(item => new Thumbnail
                {
                    ImageStream = item.ThumbnailStream,
                    ImageTitle = item.Title,
                    Order = item.Order,
                    FullImageUrl = item.FullUrl,
                    ABN = item.GetFieldValue(Helpers.Constants.SharePointFields.ABN),
                    AFSL = item.GetFieldValue(Helpers.Constants.SharePointFields.AFSL),
                    WebSite = item.GetFieldValue(Helpers.Constants.SharePointFields.Website)
                }).Cast<IThumbnail>().ToList();
        }
    }
}
