using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;

namespace OAMPS.Office.BusinessLogic.Presenters.Wizards
{
    public class QuoteSlipWizardPresenter : BaseWizardPresenter
    {
        public QuoteSlipWizardPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {
        }

        public void PopulateClaimMadeWarningFragment(bool populateFragement, string url)
        {
            if (!populateFragement) return;

            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ClaimsMadeWarning))
            {
                Document.InsertPageBreak();
                Document.InsertFile(url);
                // Document.InsertPageBreak();
                Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ClaimsMadeWarning, "true");
            }
        }

        public void PopulateApprovalFormFragment(bool populateFragement, string url)
        {
            if (!populateFragement) return;
            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ApprovalForm))
            {
                Document.InsertFile(url);
                Document.InsertPageBreak();

                Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.ApprovalForm, "true");
            }
        }


        //public void InsertQuestionnaireFragement(List<IPolicyClass> list)
        //{
        //    Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.BasisOfCoverPrevious);
        //    var ids = new StringBuilder();

        //    foreach (var policyClass in list)
        //    {
        //        if (policyClass.FragmentPolicyUrl != null)
        //        {
        //            Document.InsertFile(policyClass.FragmentPolicyUrl);
        //            ids.Append(policyClass.Id);
        //            ids.Append(";");
        //            Document.InsertPageBreak();
        //        }
        //    }

        //    Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes, ids.ToString());
        //}

        public void InsertPolicySchedule(List<IQuestionClass> list, bool atEndOfDoc)
        {
            //if (atEndOfDoc)
            //    Document.MoveToEndOfDocument();
            //else
            //    Document.MoveToStartOfDocument();

            Document.MoveCursorToStartOfBookmark("ScheduleBookmark");

            var ids = new StringBuilder();
            var count = 1;
            foreach (var q in list.Where(q => !string.IsNullOrEmpty(q?.Url)))
            {
                Document.InsertFile(q.Url);
                Document.PopulateControl(Constants.WordContentControls.QuoteSlipTitle, q.Title);
                Document.PopulateControl(Constants.WordContentControls.DocumentTitle, "Quote Slip");
                ids.Append(q.Id);
                ids.Append(";");

                if (count < list.Count)
                    Document.InsertRealPageBreak();

                count++;
            }

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes, ids.ToString());
        }

        public void CreateQuoteSlip(IQuestionClass q, IDocument quoteSlipDoc)
        {
            quoteSlipDoc.MoveCursorToStartOfBookmark("ScheduleBookmark");
            var ids = new StringBuilder();
            if (string.IsNullOrEmpty(q?.Url)) return;
            quoteSlipDoc.InsertFile(q.Url);
            quoteSlipDoc.PopulateControl(Constants.WordContentControls.QuoteSlipTitle, q.Title);
            quoteSlipDoc.PopulateControl(Constants.WordContentControls.DocumentTitle, "Quote Slip");
            ids.Append(q.Id);
            ids.Append(";");
            quoteSlipDoc.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes,
                ids.ToString());
        }

        public void InsertQuestionnaireFragement(List<IQuestionClass> list)
        {
            Document.MoveToEndOfDocument();
            var ids = new StringBuilder();

            foreach (var q in list)
            {
                if (q.Url == null) continue;
                Document.InsertFile(q.Url);
                ids.Append(q.Id);
                ids.Append(";");
                Document.InsertRealPageBreak();

                Document.DeleteControl("SectionTitle");
            }

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes, ids.ToString());
        }
    }
}