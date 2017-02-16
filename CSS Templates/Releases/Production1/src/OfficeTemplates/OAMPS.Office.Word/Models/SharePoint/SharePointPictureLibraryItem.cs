using System;
using System.IO;
using Microsoft.SharePoint.Client;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;


namespace OAMPS.Office.Word.Models.SharePoint
{
    public class SharePointPictureLibraryItem : ISharePointPictureLibraryItem
    {
        private readonly ListItem _listItem;
        private readonly ClientContext _context;

        public SharePointPictureLibraryItem(ListItem listItem, ClientContext context)
        {
            _listItem = listItem;
            _context = context;
        }

        public string Title
        {
            get 
            { 
                var title =  _listItem["Title"];
                return (title != null) ? title.ToString() : String.Empty; 
            }
        }

        public string FileRef
        {

            get { return _listItem[BusinessLogic.Helpers.Constants.SharePointFields.FileRef].ToString(); }

        }

        public object File
        {
            get { return _listItem.File; }

        }

        public void DisplayMessage(string text, string caption)
        {
            throw new NotImplementedException();
        }


        public string FullUrl
        {
            get { return new Uri(new Uri(_listItem.Context.Url), _listItem[BusinessLogic.Helpers.Constants.SharePointFields.FileRef].ToString()).ToString(); }
        }

        public System.IO.Stream ThumbnailStream
        {
            get
            {
                return Microsoft.SharePoint.Client.File.OpenBinaryDirect(_context, ThumbnailUrl).Stream;
            }
        }

        public string ThumbnailUrl
        {
            get
            {
                var imageUrl = _listItem[BusinessLogic.Helpers.Constants.SharePointFields.FileRef].ToString();
                var imageName = imageUrl.Substring(imageUrl.LastIndexOf("/", StringComparison.Ordinal)).Replace("/", String.Empty);
                var imageExtension = imageUrl.Substring(imageUrl.LastIndexOf(".", StringComparison.Ordinal) + 1);
                var thumnailName = imageName.Substring(0, imageName.LastIndexOf(".", StringComparison.Ordinal)) + "_" + imageExtension + ".jpg";
                var listName = _listItem[BusinessLogic.Helpers.Constants.SharePointFields.FileRef].ToString().Substring(0, _listItem[BusinessLogic.Helpers.Constants.SharePointFields.FileRef].ToString().IndexOf(imageName, StringComparison.Ordinal));
                return (listName + "_t/" + thumnailName);
            }
        }

        public int Order
        {
            get
            {
                object output;
                _listItem.FieldValues.TryGetValue("SortOrder", out output);

                return output == null ? 0 : Int16.Parse(output.ToString());
            }
        }


        public string GetFieldValue(string fieldName)
        {
            object output;
            _listItem.FieldValues.TryGetValue(fieldName, out output);
            return output == null ? String.Empty : output.ToString();
            //return _listItem[fieldName].ToString();
        }

        public string GetLookupFieldValue(string fieldName)
        {
            object output;
            _listItem.FieldValues.TryGetValue(fieldName, out output);
            return output == null ? String.Empty : ((FieldLookupValue)output).LookupValue;
        }
    }
}
