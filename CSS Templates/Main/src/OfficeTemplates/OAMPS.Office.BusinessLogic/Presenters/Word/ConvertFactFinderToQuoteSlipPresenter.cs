using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
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
        private readonly IDocument _document;
        private readonly IBaseView _view;

        public ConvertFactFinderToQuoteSlipPresenter(IDocument document, IBaseView view)
        {
            _document = document;
            _view = view;
        }

        public string GetIncludedPolicies(IDocument quoteSlipDoc)
        {
            return quoteSlipDoc.GetPropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes);
        }

        public void CloseInfoPanel()
        {
            var uiScheduler = TaskScheduler.Default;
            Task.Factory.StartNew(() =>
            {
                var endAt = DateTime.Now.AddSeconds(30);
                while (DateTime.Now < endAt)
                {
                    _document.CloseInformationPanelOnAllDocuments();
                    Thread.Sleep(1000);
                }
            }, CancellationToken.None, TaskCreationOptions.None, uiScheduler);
        }

        public void ConvertToQuoteSlip(string quoteSlipServerRelativeUrl, IDocument quoteSlipDoc,
            IQuoteSlipSchedules schedule, IBaseTemplate sourceTemplateData)
        {
            var quoteSlipWizardPresenter = new QuoteSlipWizardPresenter(quoteSlipDoc, _view);

            var question = new QuestionClass
            {
                Id = schedule.Id,
                Title = schedule.Title,
                Url = schedule.Url
            };
            var questions = new List<IQuestionClass> {question};
            quoteSlipWizardPresenter.InsertPolicySchedule(questions, true);
            InsertUnderwritingInformations(quoteSlipDoc, schedule);
            CopyAllContentControlValuesBetweenTemplates(quoteSlipDoc);
            CopyAllTablesBetweenTemplates(quoteSlipDoc);
            FinaliseConvert(quoteSlipDoc, sourceTemplateData, quoteSlipWizardPresenter);
        }

        private void CopyAllContentControlValuesBetweenTemplates(IDocument quoteSlipDoc)
        {
            var allFields = _document.GetContentControlsByName(string.Empty);
            foreach (var controlName in allFields)
            {
                var value = _document.ReadContentControlValue(controlName);
                quoteSlipDoc.PopulateControl(controlName, value);
            }
        }

        private void CopyAllTablesBetweenTemplates(IDocument quoteSlipDoc)
        {
            var allTables = _document.GetTablesByNamePrefix(string.Empty);
            foreach (var tableName in allTables)
            {
                if (!quoteSlipDoc.TableExists(tableName)) continue;
                if (tableName.StartsWith("convert:")) continue;

                _document.CopyTable(tableName);
                quoteSlipDoc.Activate();
                quoteSlipDoc.SelectTable(tableName);
                quoteSlipDoc.DeleteTable(tableName);
                quoteSlipDoc.InsertParagraphBreak();
                quoteSlipDoc.TypeText(RemovePrefix(tableName).Trim(), Constants.WordStyles.Bold);
                quoteSlipDoc.InsertParagraphBreak();
                quoteSlipDoc.PasteClipboardOriginalFormatting();
                quoteSlipDoc.InsertParagraphBreak();
            }
        }

        private string RemovePrefix(string val)
        {
            val = val.Replace("HMF", string.Empty);
            val = val.Replace("CMF", string.Empty);
            val = val.Replace("CT", string.Empty);
            val = val.Replace("ISR", string.Empty);
            val = val.Replace("AC", string.Empty);
            val = val.Replace("PAI", string.Empty);
            val = val.Replace("PPL", string.Empty);
            val = val.Replace("RVRAL", string.Empty);
            val = val.Replace("WC", string.Empty);

            return val;
        }

        private static void FinaliseConvert(IDocument quoteSlipDoc, IBaseTemplate sourceTemplateData,
            QuoteSlipWizardPresenter quoteSlipWizardPresenter)
        {
            sourceTemplateData.DocumentTitle = "Quotation Slip";
            quoteSlipWizardPresenter.PopulateData(sourceTemplateData);
            quoteSlipDoc.MoveToStartOfDocument();
            quoteSlipDoc.CloseInformationPanel(true);
        }

        private void InsertUnderwritingInformations(IDocument quoteSlipDoc, IQuoteSlipSchedules schedule)
        {
            var s = quoteSlipDoc.GetBookmarkStartRange(Constants.WordBookmarks.UnderwritingStart);
            var e = quoteSlipDoc.GetBookmarkEndRange(Constants.WordBookmarks.UnderwritingEnd);
            quoteSlipDoc.DeleteRange(s, e);

            quoteSlipDoc.MoveCursorToStartOfBookmark(Constants.WordBookmarks.UnderwritingStart);

            var prefix = "convert:" + schedule.Id;
            var tableNamesToCopy = _document.GetTablesByNamePrefix(prefix);

            var lastMode = "port";
            foreach (var tableName in tableNamesToCopy)
            {
                _document.CopyTable(tableName);

                quoteSlipDoc.Activate();

                if (tableName.Contains("landscape"))
                {
                    if (lastMode.Equals("port", StringComparison.OrdinalIgnoreCase))
                    {
                        lastMode = "landscape";
                        quoteSlipDoc.InsertLandscapePage("nLand");
                        quoteSlipDoc.MoveCursorToStartOfBookmark("nLand");
                    }
                }
                else
                {
                    if (lastMode.Equals("landscape", StringComparison.OrdinalIgnoreCase))
                    {
                        quoteSlipDoc.InsertPortraitPage("nPort");
                        quoteSlipDoc.MoveCursorToStartOfBookmark("nPort");
                    }
                    else
                    {
                        quoteSlipDoc.InsertParagraphBreak();
                    }

                    lastMode = "port"; //oh yeah ill have some port.  some tawny port
                }

                quoteSlipDoc.InsertParagraphBreak();
                var indexOfFirstPrefix = tableName.IndexOf('-') + 1;
                quoteSlipDoc.TypeText(tableName.Remove(0, indexOfFirstPrefix), Constants.WordStyles.Bold);
                    //remove prefix
                quoteSlipDoc.PasteClipboardOriginalFormatting();
                quoteSlipDoc.InsertParagraphBreak();
            }
        }
    }
}