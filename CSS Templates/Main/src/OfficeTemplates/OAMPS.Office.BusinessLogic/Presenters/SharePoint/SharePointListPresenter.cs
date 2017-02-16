using System.Collections.Generic;
using System.Linq;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Logging;
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
                TitleNoWhiteSpace = item.Title.Replace(" ", string.Empty),
                Url = item.FullUrl,
                Id = item.GetFieldValue(Constants.SharePointFields.FieldId),
                MajorClass = item.GetLookupFieldValue(Constants.SharePointFields.MajorPolicyClass)
            }).Cast<IPolicyClass>().ToList();
        }

        //public ISharePointListItem GetQuoteSlipSchedule(string policyId)
        //{

        //    return string.IsNullOrEmpty(policyId) ? null : //List.GetListItemByLookupField("Questionares", policyId);
        //}

        public List<IQuestionClass> GetPreRenewalQuestionaireQuestions()
        {
            var items = List.ListItems;
            return items.Select(item => new QuestionClass
            {
                Title = item.Title,
                Url = item.FileRef,
                Id = item.GetFieldValue(Constants.SharePointFields.FieldId)
            }).Cast<IQuestionClass>().ToList();
        }

        public List<IManualClaimsProcedure> GetManualClaimsProcedure()
        {
            var items = List.ListItems;
            return items.Select(item => new ManualClaimsProcedure
            {
                Title = item.Title,
                Url = item.FileRef,
                Id = item.GetFieldValue(Constants.SharePointFields.FieldId),
                PolicyClass = item.GetLookupFieldValue(Constants.SharePointFields.MinorClass)
            }).Cast<IManualClaimsProcedure>().ToList();
        }

        public List<IQuoteSlipSchedules> GetScheduleItems()
        {
            var items = List.ListItems;
            return items.Select(item => new QuoteSlipSchedules
            {
                Title = item.Title,
                Url = item.FileRef,
                Id = item.GetFieldValue(Constants.SharePointFields.FieldId),
                LinkedQuestionUrl = item.GetLookupFieldValueArray(Constants.SharePointFields.FactFinder)
            }).Cast<IQuoteSlipSchedules>().ToList();
        }

        public List<IQuestionClass> GetMinorQuestionItems()
        {
            var items = List.ListItems;
            return items.Select(item => new QuestionClass
            {
                Title = item.Title,
                //TitleNoWhiteSpace = item.Title.Replace(" ", string.Empty),
                Url = item.FullUrl,
                TopCategory = item.GetLookupFieldValue(Constants.SharePointFields.TopQuestionClass),
                SubCategory = item.GetFieldValue(Constants.SharePointFields.SubQuestionClass),
                Id = item.GetFieldValue(Constants.SharePointFields.FieldId)
            }).Cast<IQuestionClass>().ToList();
        }

        public List<IPreRenewalQuestionareMappings> GetPreRenewalQuestionareMappings()
        {
            var items = List.ListItems;
            return items.Select(x => new PreRenewalQuestionareMappings
            {
                FragmentUrl = x.FileRef,
                Title = x.Title,
                PolicyType =
                    x.GetLookupFieldValueArray(Constants.SharePointFields.PreRenewalQuestionareMappingsPolicyClasses)
            }).Cast<IPreRenewalQuestionareMappings>().ToList();
        }

        public List<IInsurer> GetInsurers()
        {
            var items = List.ListItems;
            return items.Select(item => new Insurer
            {
                Title = item.Title,
                Id = item.GetFieldValue(Constants.SharePointFields.FieldId),
                Category = item.GetFieldValue(Constants.SharePointFields.Category)
            }).Cast<IInsurer>().ToList();
        }

        //public void LogUsage(IUsageLogger usageLogger)
        //{
        //    usageLogger.LogUsage();
        //}
    }
}