using System.Collections.Generic;
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
            foreach (var q in list)
            {
                if (q == null) continue;
                if (string.IsNullOrEmpty(q.Url)) continue;
                Document.InsertFile(q.Url);
                //Document.RenameControl(Constants.WordContentControls.QuoteSlipTitle, Constants.WordContentControls.QuoteSlipTitle + q.Id);
                //Document.PopulateControl(Constants.WordContentControls.QuoteSlipTitle + q.Id, q.Title);
                Document.PopulateControl(Constants.WordContentControls.QuoteSlipTitle, q.Title);

                //Document.RenameControl(Constants.WordContentControls.DocumentTitle, Constants.WordContentControls.DocumentTitle + q.Id);
                //Document.PopulateControl(Constants.WordContentControls.DocumentTitle + q.Id, "Quote Slip");
                Document.PopulateControl(Constants.WordContentControls.DocumentTitle, "Quote Slip");

                ids.Append(q.Id);
                ids.Append(";");

                if (count < list.Count)
                    Document.InsertRealPageBreak();

                count++;
            }

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes, ids.ToString());
        }

        public void InsertQuestionnaireFragement(List<IQuestionClass> list)
        {
            Document.MoveToEndOfDocument();
            var ids = new StringBuilder();

            foreach (IQuestionClass q in list)
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