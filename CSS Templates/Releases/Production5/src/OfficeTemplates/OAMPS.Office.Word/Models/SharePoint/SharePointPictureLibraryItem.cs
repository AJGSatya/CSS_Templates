using System;
using System.Globalization;
using System.IO;
using Microsoft.SharePoint.Client;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

namespace OAMPS.Office.Word.Models.SharePoint
{
    public class SharePointPictureLibraryItem : ISharePointPictureLibraryItem
    {
        private readonly ClientContext _context;
        private readonly ListItem _listItem;

        public SharePointPictureLibraryItem(ListItem listItem, ClientContext context)
        {
            _listItem = listItem;
            _context = context;
        }

        public object File
        {
            get { return _listItem.File; }
        }

        public string Title
        {
            get
            {
                object title = _listItem["Title"];
                return (title != null) ? title.ToString() : String.Empty;
            }
        }

        public string FileRef
        {
            get { return _listItem[Constants.SharePointFields.FileRef].ToString(); }
        }

        public void DisplayMessage(string text, string caption)
        {
            throw new NotImplementedException();
        }


        public string FullUrl
        {
            get { return new Uri(new Uri(_listItem.Context.Url), _listItem[Constants.SharePointFields.FileRef].ToString()).ToString(); }
        }

        public byte[] ThumbnailStream
        {
            get
            {
                FileInformation fl = Microsoft.SharePoint.Client.File.OpenBinaryDirect(_context, ThumbnailUrl);
                return ReadFully(fl.Stream);
            }
        }

        public string ThumbnailUrl
        {
            get
            {
                string imageUrl = _listItem[Constants.SharePointFields.FileRef].ToString();
                string imageName = imageUrl.Substring(imageUrl.LastIndexOf("/", StringComparison.Ordinal)).Replace("/", String.Empty);
                string imageExtension = imageUrl.Substring(imageUrl.LastIndexOf(".", StringComparison.Ordinal) + 1);
                string thumnailName = imageName.Substring(0, imageName.LastIndexOf(".", StringComparison.Ordinal)) + "_" + imageExtension + ".jpg";
                string listName = _listItem[Constants.SharePointFields.FileRef].ToString().Substring(0, _listItem[Constants.SharePointFields.FileRef].ToString().IndexOf(imageName, StringComparison.Ordinal));
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

        public string[] GetLookupFieldValueArray(string fieldName)
        {
            object output;
            _listItem.FieldValues.TryGetValue(fieldName, out output);
            if (output == null)
            {
                return new string[0];
            }
            {
                var f = ((FieldLookupValue[]) output);
                var s = new string[f.Length];

                for (int i = 0; i < f.Length; i++)
                {
                    FieldLookupValue t = f[i];
                    s[i] = t.LookupValue;
                }

                return s;
            }
        }

        public string[] GetLookupFieldValueIdArray(string fieldName)
        {
            object output;
            _listItem.FieldValues.TryGetValue(fieldName, out output);
            if (output == null)
            {
                return new string[0];
            }
            {
                var f = ((FieldLookupValue[])output);
                var s = new string[f.Length];

                for (int i = 0; i < f.Length; i++)
                {
                    FieldLookupValue t = f[i];
                    s[i] = t.LookupId.ToString();
                }

                return s;
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
            return output == null ? String.Empty : ((FieldLookupValue) output).LookupValue;
        }

        public string GetLookupFieldValueId(string fieldName)
        {
            object output;
            _listItem.FieldValues.TryGetValue(fieldName, out output);
            return output == null ? String.Empty : ((FieldLookupValue)output).LookupId.ToString(CultureInfo.CurrentUICulture);
        }

        private byte[] ReadFully(Stream input)
        {
            var buffer = new byte[16*1024];
            using (var ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
    }
}