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
    }
}
