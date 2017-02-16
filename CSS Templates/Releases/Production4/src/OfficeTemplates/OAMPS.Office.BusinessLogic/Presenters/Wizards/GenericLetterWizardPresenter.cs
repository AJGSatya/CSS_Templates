using System.Collections.Generic;
using System.Linq;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Presenters.Wizards;

namespace OAMPS.Office.BusinessLogic.Presenters
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
            for (var i = 0; i < bookmarks.Count; i++)
            {
                var b = bookmarks[i];
                d.InsertFile(b.Url);

                if (i != (bookmarks.Count - 1))
                    d.InsertPageBreak();
            }

            d.TurnOnProtection(protection, string.Empty);
            d.CloseInformationPanel(true);
        }
    }
}