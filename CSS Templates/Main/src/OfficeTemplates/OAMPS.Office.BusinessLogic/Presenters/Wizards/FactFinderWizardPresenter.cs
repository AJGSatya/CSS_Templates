using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters.Wizards
{
    public class FactFinderWizardPresenter : BaseWizardPresenter
    {
        public FactFinderWizardPresenter(IDocument document, IBaseView view) : base(document, view)
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

        public void GenerateQuestionsOnly(List<IQuestionClass> list, string fragUrl)
        {
            var d = Document.OpenFile(fragUrl);
            var protection = d.TurnOffProtection(string.Empty); //dont turn protection back on
            var ids = new StringBuilder();
            foreach (var q in list)
            {
                if (q.Url != null)
                {
                    //d.MoveCursorToStartOfBookmark("Unlocked");
                    d.InsertFile(q.Url);
                    if (d.HasBookmark(Constants.WordBookmarks.FactFinderStart))
                        d.RenameBookmark(Constants.WordBookmarks.FactFinderStart,
                            Constants.WordBookmarks.FactFinderStart + q.Id);

                    if (d.HasBookmark(Constants.WordBookmarks.FactFinderEnd))
                        d.RenameBookmark(Constants.WordBookmarks.FactFinderEnd,
                            Constants.WordBookmarks.FactFinderEnd + q.Id);

                    d.InsertRealPageBreak();

                    ids.Append(q.Id);
                    ids.Append(";");
                }
            }

            d.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes, ids.ToString());

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.BuiltInTitle,
                Constants.TemplateNames.PreRenewalQuestionnaire);

            d.CloseInformationPanel(true);
        }

        public void PopulateImportantNotices(List<DocumentFragment> fragments, string importantNoticesUrl,
            string privacyStatementUrl, string fsgUrl, string termsOfEngagementUrl)
        {
            if (!Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.ImportantNotes)) return;

            //Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.StatutoryInformation, statutory.ToString());

            ////insert stat notices
            //Document.InsertFile(importantNoticesUrl); //in settings
            //Document.InsertPageBreak();

            foreach (var b in fragments)
            {
                Document.InsertFile(b.Url);
                if (b != fragments.LastOrDefault())
                {
                    Document.InsertPageBreak();
                }
            }

            //switch (statutory)
            //{
            //    case Enums.Statutory.Retail:
            //        {
            //            Document.InsertFile(fsgUrl); //in settings
            //            break;
            //        }
            //    case Enums.Statutory.Wholesale:
            //        {
            //            //insert pivacy statement
            //            //Document.InsertFile(privacyStatementUrl); //compliance rule change, privacy not required ever.
            //            //Document.InsertPageBreak();

            //            Document.InsertFile(termsOfEngagementUrl); //in settings
            //            break;
            //        }
            //    case Enums.Statutory.WholesaleWithRetail:
            //        {
            //            //insert pivacy statement
            //            //Document.InsertFile(privacyStatementUrl); //compliance rule change, privacy not required ever.
            //            //Document.InsertPageBreak();

            //            Document.InsertFile(fsgUrl); //in settings

            //            Document.InsertPageBreak();
            //            Document.InsertFile(termsOfEngagementUrl); //in settings
            //            break;
            //        }
            //}
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

        public void InsertQuestionnaireFragement(List<IQuestionClass> list, string infoFragUrl)
        {
            Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.BasisOfCoverPrevious);
            var ids = new StringBuilder();
            //if (list.Count > 0)
            //// Document.InsertFile(infoFragURL);

            foreach (var q in list)
            {
                if (q.Url != null)
                {
                    Document.InsertFile(q.Url);
                    if (Document.HasBookmark(Constants.WordBookmarks.FactFinderStart))
                        Document.RenameBookmark(Constants.WordBookmarks.FactFinderStart,
                            Constants.WordBookmarks.FactFinderStart + q.Id);

                    if (Document.HasBookmark(Constants.WordBookmarks.FactFinderEnd))
                        Document.RenameBookmark(Constants.WordBookmarks.FactFinderEnd,
                            Constants.WordBookmarks.FactFinderEnd + q.Id);

                    ids.Append(q.Id);
                    ids.Append(";");
                    Document.InsertRealPageBreak();
                }
            }

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.IncludedPolicyTypes, ids.ToString());
        }
    }
}