using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Template;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters.Word
{
    public class ConvertFactFinderToQuoteSlipPresenter
    {
        private IDocument _document;
        private IBaseView _view;

        public ConvertFactFinderToQuoteSlipPresenter(IDocument document, IBaseView view)
        {
            _document = document;
            _view = view;
        }

        public string GetIncludedPolicies(IDocument quoteSlipDoc)
        {
           return quoteSlipDoc.GetPropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes);
        }

        public void ConvertToQuoteSlip(string quoteSlipServerRelativeUrl, IDocument quoteSlipDoc, IQuoteSlipSchedules schedule, IBaseTemplate sourceTemplateData)
        {
            var quoteSlipWizardPresenter = new QuoteSlipWizardPresenter(quoteSlipDoc, _view);

            var question = new QuestionClass()
                {
                    Id = schedule.Id,
                    Title = schedule.Title,
                    Url = schedule.Url
                };
            var questions = new List<IQuestionClass> { question };

            quoteSlipWizardPresenter.InsertPolicySchedule(questions, true);

            var startRange = _document.GetBookmarkStartRange(Constants.WordBookmarks.FactFinderStart + schedule.LinkedQuestionId);
            var endRange = _document.GetBookmarkEndRange(Constants.WordBookmarks.FactFinderEnd + schedule.LinkedQuestionId);

            if(startRange < 0 || endRange < 0) //todo: ???
                return;
            
            _document.CopyRange(startRange, endRange);
            quoteSlipDoc.MoveToEndOfDocument();
            quoteSlipDoc.PasteClipboard();


            quoteSlipWizardPresenter.PopulateData(sourceTemplateData);

            quoteSlipDoc.MoveToStartOfDocument();
            quoteSlipDoc.CloseInformationPanel(true);

            
        }
    }
}
