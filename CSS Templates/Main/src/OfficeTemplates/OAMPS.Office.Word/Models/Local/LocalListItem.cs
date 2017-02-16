using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.Word.Helpers;
using OAMPS.Office.Word.Helpers.LocalSharePoint;
using OAMPS.Office.Word.Properties;

namespace OAMPS.Office.Word.Models.Local
{
    public class LocalListItem : ISharePointListItem
    {
        private readonly ListItem _listItem;

        public LocalListItem(ListItem listItem)
        {
            _listItem = listItem;
        }

        public string Title => _listItem.GetFieldValue("Title");

        public string FileRef
        {
            get
            {
                var url = _listItem.GetFieldValue("FileRef").ToLower();
                var filePath = url.Replace(_listItem.GetFieldValue("ParentList.ParentWebUrl").ToLower(), "");
                filePath = filePath.Replace("/", "\\");
                return $"file://{Settings.Default.LocalFullPathRoot}\\{filePath}";
            }
        }
    
        public string FullUrl => FileRef;

        public string GetFieldValue(string fieldName) => _listItem.GetFieldValue(fieldName);

        public string GetLookupFieldValue(string fieldName) => _listItem.GetJObjectPropertyValue(fieldName, "LookupValue");

        public string[] GetLookupFieldValueArray(string fieldName) => _listItem.GetJObjectPropertyValues(fieldName, "LookupValue");

        public string GetLookupFieldValueId(string fieldName) => _listItem.GetJObjectPropertyValue(fieldName, "LookupId");

        public string[] GetLookupFieldValueIdArray(string fieldName) => _listItem.GetJObjectPropertyValues(fieldName, "LookupId");

        public void DisplayMessage(string text, string caption)
        {
            // unused
            throw new NotImplementedException();
        }

    }
}
