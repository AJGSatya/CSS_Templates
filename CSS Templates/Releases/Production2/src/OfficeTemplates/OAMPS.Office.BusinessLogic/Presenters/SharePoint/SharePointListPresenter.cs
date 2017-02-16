using System;
using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Interfaces;
using System.Linq;
using OAMPS.Office.BusinessLogic.Interfaces.SharePoint;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters.SharePoint
{
    public class SharePointListPresenter
    {
       protected ISharePointList List;
       protected IBaseView View;
        public SharePointListPresenter(ISharePointList list, IBaseView view)
        {
            List = list;
            View = view;
        }

        public List<ISharePointListItem> GetItems()
        {
            return List.ListItems;
        }

        public ISharePointListItem GetItemByTitle(string title)
        {
            return List.GetListItemByTitle(title);
        }

        public List<ISharePointListItem> GetHelpItems()
        {
            return List.HelpListItems;
        }

        public List<IPolicyClass> GetMinorPolicyItems()
        {
            var items = List.ListItems;
            return items.Select(item => new PolicyClass
            {
                Title = item.Title,
                TitleNoWhiteSpace =  item.Title.Replace(" ", string.Empty),
                Url = item.FullUrl,
                Id = item.GetFieldValue(Helpers.Constants.SharePointFields.FieldID),
                MajorClass =  item.GetLookupFieldValue(Helpers.Constants.SharePointFields.MajorPolicyClass)
            }).Cast<IPolicyClass>().ToList();
        }

        public List<IQuestionClass> GetMinorQuestionItems()
        {
            var items = List.ListItems;
            return items.Select(item => new QuestionClass()
            {
                Title = item.Title,
                //TitleNoWhiteSpace = item.Title.Replace(" ", string.Empty),
                Url = item.FullUrl,
                TopCategory = item.GetLookupFieldValue(Helpers.Constants.SharePointFields.TopQuestionClass),
                SubCategory = item.GetFieldValue(Helpers.Constants.SharePointFields.SubQuestionClass),
                Id = item.GetFieldValue(Helpers.Constants.SharePointFields.FieldID)
            }).Cast<IQuestionClass>().ToList();
        }

        public List<IInsurer> GetInsurers()
        {
            var items = List.ListItems;
            return items.Select(item => new Insurer
            {
                Title = item.Title,
                Id = item.GetFieldValue(Helpers.Constants.SharePointFields.FieldID),
                Category = item.GetFieldValue(Helpers.Constants.SharePointFields.Category)
            }).Cast<IInsurer>().ToList();
        }


        public void LogUsage(string trackingType,string report, string accountExec, string userDepartment, string userOffice, string clientName,string segment, string wholesaleOrRetail, string  captureDate, string captureTime)
        {
            var values = new Dictionary<string, string>();
            values.Add("Title",report);
            values.Add("Account_x0020_Exec_x0020_in_x002", accountExec);
            values.Add("Department_x0020_of_x0020_user_x", userDepartment);
            values.Add("Office_x002f_Branch_x0020_of_x00", userOffice);
            values.Add("Client_x0020_Name_x0020__x0028_F", clientName);
            values.Add("Segment_x0020__x0028_2_x002c_3_x", segment);
            values.Add("Retail_x002f_Wholesale_x0020__x0", wholesaleOrRetail);
            values.Add("Date_x0020__x0028_captured_x0020", captureDate);
            values.Add("Time_x0020__x0028_captured_x0020", captureTime);
            values.Add("Tracking_x0020_Type", trackingType);
            List.AddItem(values);
        }
    }
}
