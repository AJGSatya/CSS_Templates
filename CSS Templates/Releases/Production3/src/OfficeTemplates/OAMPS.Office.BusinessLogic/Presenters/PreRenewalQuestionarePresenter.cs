using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;

namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class PreRenewalQuestionarePresenter : BasePresenter
    {
        public PreRenewalQuestionarePresenter(IDocument document, IBaseView view) : base(document, view)
        {

        }

        public void PopulateClaimMadeWarningFragment(bool populateFragement, string url)
        {
            if (!populateFragement) return;

            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ClaimsMadeWarning))
            {
                Document.InsertFile(url);
                Document.InsertPageBreak();
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

        public void InsertQuestionnaireFragement(List<IQuestionClass> list)
        {
            Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.BasisOfCoverPrevious);
            var ids = new StringBuilder();

            foreach (var q in list)
            {
                if (q.Url != null)
                {
                    Document.InsertFile(q.Url);
                    ids.Append(q.Id);
                    ids.Append(";");
                    Document.InsertRealPageBreak();
                }
            }

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes, ids.ToString());
        }
    }
}
