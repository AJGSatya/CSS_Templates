using System;

using Microsoft.SharePoint.Client;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;

namespace OAMPS.Office.Word.Models.SharePoint
{
    public class SharePointListItem : ISharePointListItem
    {
        private readonly ListItem _listItem;

        public SharePointListItem(ListItem listItem)
        {
            _listItem = listItem;
        }

        public string Title
        {
            get { return _listItem["Title"].ToString(); }
        }

        public string FileRef
        {
            get { return _listItem[BusinessLogic.Helpers.Constants.SharePointFields.FileRef].ToString(); }
        }

        public string GetFieldValue(string fieldName)
        {
            object output;
            _listItem.FieldValues.TryGetValue(fieldName, out output);
            return output == null ? String.Empty : output.ToString();
        }

        public string GetLookupFieldValue(string fieldName)
        {
            object output;
            _listItem.FieldValues.TryGetValue(fieldName, out output);
            return output == null ? String.Empty : ((FieldLookupValue)output).LookupValue;
            //return ((FieldLookupValue)_listItem[fieldName]).LookupValue;
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
                    var t = f[i];
                    s[i] = t.LookupValue;
                }

                return s;
            }
        }
        
        public string GetChoiceFieldValue(string fieldName)
        {
            object output;
            _listItem.FieldValues.TryGetValue(fieldName, out output);
            return output == null ? String.Empty : output.ToString(); //GetLookupFieldValue()
            //return ((FieldLookupValue)_listItem[fieldName]).LookupValue;
        }


        public void DisplayMessage(string text, string caption)
        {
            throw new NotImplementedException();
        }
    
        public string FullUrl
        {
            get { return _listItem.Context.Url + _listItem[BusinessLogic.Helpers.Constants.SharePointFields.FileRef]; }
        }
    }
}
