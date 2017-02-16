//using System;
//using System.Globalization;
//using System.IO;
//using Microsoft.SharePoint.Client;
//using OAMPS.Office.BusinessLogic.Helpers;
//using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

//namespace OAMPS.Office.Word.Models.SharePoint
//{
//    public class SharePointPictureLibraryItem : ISharePointPictureLibraryItem
//    {
//        private readonly SharePointAdalClientContext _context;
//        private readonly ListItem _listItem;

//        public SharePointPictureLibraryItem(ListItem listItem, SharePointAdalClientContext context)
//        {
//            _listItem = listItem;
//            _context = context;
//        }

//        public object File
//        {
//            get { return _listItem.File; }
//        }

//        public string Title
//        {
//            get
//            {
//                var title = _listItem["Title"];
//                return (title != null) ? title.ToString() : string.Empty;
//            }
//        }

//        public string FileRef
//        {
//            get { return _listItem[Constants.SharePointFields.FileRef].ToString(); }
//        }

//        public void DisplayMessage(string text, string caption)
//        {
//            throw new NotImplementedException();
//        }


//        public string FullUrl
//        {
//            get
//            {
//                return
//                    new Uri(new Uri(_listItem.Context.Url), _listItem[Constants.SharePointFields.FileRef].ToString())
//                        .ToString();
//            }
//        }


//        public byte[] ThumbnailStream
//        {
//            get
//            {
//                var tempContext = new SharePointUserDirectClientContext(_context.Url);
//                var fl = Microsoft.SharePoint.Client.File.OpenBinaryDirect(tempContext, ThumbnailUrl);
//                return ReadFully(fl.Stream);
//            }
//        }


//        public string ThumbnailUrl
//        {
//            get
//            {
//                var imageUrl = _listItem[Constants.SharePointFields.FileRef].ToString();
//                var imageName = imageUrl.Substring(imageUrl.LastIndexOf("/", StringComparison.Ordinal))
//                    .Replace("/", string.Empty);
//                var imageExtension = imageUrl.Substring(imageUrl.LastIndexOf(".", StringComparison.Ordinal) + 1);
//                var thumnailName = imageName.Substring(0, imageName.LastIndexOf(".", StringComparison.Ordinal)) + "_" +
//                                   imageExtension + ".jpg";
//                var listName = _listItem[Constants.SharePointFields.FileRef].ToString()
//                    .Substring(0,
//                        _listItem[Constants.SharePointFields.FileRef].ToString()
//                            .IndexOf(imageName, StringComparison.Ordinal));
//                return (listName + "_t/" + thumnailName);
//            }
//        }

//        public int Order
//        {
//            get
//            {
//                object output;
//                _listItem.FieldValues.TryGetValue("SortOrder", out output);

//                return output == null ? 0 : short.Parse(output.ToString());
//            }
//        }

//        public string[] GetLookupFieldValueArray(string fieldName)
//        {
//            object output;
//            _listItem.FieldValues.TryGetValue(fieldName, out output);
//            if (output == null)
//            {
//                return new string[0];
//            }
//            {
//                var f = ((FieldLookupValue[]) output);
//                var s = new string[f.Length];

//                for (var i = 0; i < f.Length; i++)
//                {
//                    var t = f[i];
//                    s[i] = t.LookupValue;
//                }

//                return s;
//            }
//        }

//        public string[] GetLookupFieldValueIdArray(string fieldName)
//        {
//            object output;
//            _listItem.FieldValues.TryGetValue(fieldName, out output);
//            if (output == null)
//            {
//                return new string[0];
//            }
//            {
//                var f = ((FieldLookupValue[]) output);
//                var s = new string[f.Length];

//                for (var i = 0; i < f.Length; i++)
//                {
//                    var t = f[i];
//                    s[i] = t.LookupId.ToString();
//                }

//                return s;
//            }
//        }

//        public string GetFieldValue(string fieldName)
//        {
//            object output;
//            _listItem.FieldValues.TryGetValue(fieldName, out output);
//            return output == null ? string.Empty : output.ToString();
//            //return _listItem[fieldName].ToString();
//        }

//        public string GetLookupFieldValue(string fieldName)
//        {
//            object output;
//            _listItem.FieldValues.TryGetValue(fieldName, out output);
//            return output == null ? string.Empty : ((FieldLookupValue) output).LookupValue;
//        }

//        public string GetLookupFieldValueId(string fieldName)
//        {
//            object output;
//            _listItem.FieldValues.TryGetValue(fieldName, out output);
//            return output == null
//                ? string.Empty
//                : ((FieldLookupValue) output).LookupId.ToString(CultureInfo.CurrentUICulture);
//        }

//        private byte[] ReadFully(Stream input)
//        {
//            var buffer = new byte[16*1024];
//            using (var ms = new MemoryStream())
//            {
//                int read;
//                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
//                {
//                    ms.Write(buffer, 0, read);
//                }
//                return ms.ToArray();
//            }
//        }
//    }
//}