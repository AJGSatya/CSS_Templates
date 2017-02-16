using System;
using System.Collections.Generic;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters.Wizards
{
    public class GenericLetterWizardPresenter : BaseWizardPresenter
    {
        public GenericLetterWizardPresenter(IDocument document, IBaseView view) : base(document, view)
        {
        }

        public void DeleteDocumentHeaderAndFooter()
        {
            Document.RemoveHeader();
            Document.RemoveFooter();
            Document.DeleteImage(Constants.ImageProperties.CompanyLogo);
        }

        public void InsertSubFragments(string attachmentBookMark, string mainContentBookMark,
            List<DocumentFragment> fragements, string fragUrl)
        {
            string frgKeys = null;
            if (Document.MoveCursorToStartOfBookmark(attachmentBookMark))
            {
                fragements.ForEach(
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
            for (var i = 0; i < fragements.Count; i++)
            {
                var b = fragements[i];
                d.MoveCursorToStartOfBookmark(b.Locked.Equals("true", StringComparison.OrdinalIgnoreCase)
                    ? "Locked"
                    : "Unlocked");

                d.InsertFile(b.Url);

                if (i != (fragements.Count - 1))
                    d.InsertPageBreak();
            }

            d.TurnOnProtection(protection, string.Empty);
            d.CloseInformationPanel(true);
        }
    }
}