using System;
using System.Collections.Generic;
using System.Linq;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Template;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class RenwalLetterWizardPresenter : BaseWizardPresenter
    {
        public RenwalLetterWizardPresenter(IDocument document, IBaseView view) : base(document, view)
        {
        }

        public void InsertSubFragments(string attachmentBookMark, string mainContentBookMark, List<DocumentFragment> bookmarks, string fragUrl)
        {
            string frgKeys = null;


            if (Document.MoveCursorToStartOfBookmark(attachmentBookMark))
            {
                bookmarks.ForEach(
                    x =>
                        {
                            Document.TypeText(x.Title);
                            Document.TypeText("; ");
                            frgKeys += x.Key + ";";
                        }
                    );
            }

            Document.MoveCursorToStartOfBookmark(mainContentBookMark);
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlSubFragments, frgKeys);

            var d = Document.OpenFile(fragUrl);
            var protection = d.TurnOffProtection(string.Empty);

            foreach (var b in bookmarks)
            {
                d.Activate();
                d.InsertFile(b.Url);
                if (b != bookmarks.LastOrDefault())
                {
                    d.InsertPageBreak();
                }
            }
            d.TurnOnProtection(protection, string.Empty);
            d.CloseInformationPanel(true);
        }

        public void DeleteDocumentHeaderAndFooter()
        {
            Document.RemoveHeader();
            Document.RemoveFooter();
            Document.DeleteImage(Constants.ImageProperties.CompanyLogo);
        }

        public void InsertMainFragment(string bookmark, string fragUrl, RenewalLetter template)
        {
            if (Document.MoveCursorToStartOfBookmark(bookmark))
            {
                Document.InsertFile(fragUrl);
            }
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkContacted, template.IsContactSelected.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkNewClient, template.IsNewClientSelected.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkFunding, template.IsFundingSelected.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkGaw, template.IsGawSelected.ToString());
        }

        //public void InsertMainFragment(string bookmark, List<string> bookmarks, RenewalLetter template) // bool chkContacted, bool chkNewClient, bool chkFunding)
        //{


        //    Document.MoveCursorToStartOfBookmark(bookmark);

        //    bookmarks.ForEach(
        //        (x) => Document.InsertFile(x));

        //    Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkContacted, template.IsContactSelected.ToString());
        //    Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkNewClient, template.IsNewClientSelected.ToString());
        //    Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkFunding, template.IsFundingSelected.ToString());

        //}

        public void PopulatePolicy(List<IPolicyClass> policies, string renewalDate, bool isNewInsurer)
        {
            if (policies == null || policies.Count <= 0) return;

            int r = 2; //start at second row, vsto arrays start at 1 + the 1 row for the table header.

            if (isNewInsurer)
            {
                Document.InsertTableCell(2, "Current Insurer", Constants.WordTables.RenewalLetterPolicies);
                Document.InsertTableCell(3, "Recommended Insurer", Constants.WordTables.RenewalLetterPolicies);
            }

            foreach (IPolicyClass policy in policies)
            {
                int rowCount = Document.TableRowOrColumnCount(false, Constants.WordTables.RenewalLetterPolicies);

                if (r > rowCount)
                {
                    Document.InsertTableRowNonCellLoop(rowCount, Constants.WordTables.RenewalLetterPolicies);
                }

                DateTime outDate;

                if (DateTime.TryParse(renewalDate, out outDate))
                    renewalDate = outDate.ToShortDateString();

                Document.PopulateTableCell(r, 1, policy.Title, Constants.WordTables.RenewalLetterPolicies);
                Document.PopulateTableCell(r, 2, renewalDate, Constants.WordTables.RenewalLetterPolicies);

                r++;
            }
        }

        public List<IPolicyClass> ReadPoliciesInDocument()
        {
            var items = new List<IPolicyClass>();

            int rowcount = Document.TableRowOrColumnCount(false, Constants.WordTables.RenewalLetterPolicies);

            for (int i = 1; i <= rowcount; i++)
            {
                var item = new PolicyClass
                    {
                        Title = Document.ReadTableCell(i, 1, Constants.WordTables.RenewalLetterPolicies)
                    };
                items.Add(item);
            }
            return items;
        }
    }
}