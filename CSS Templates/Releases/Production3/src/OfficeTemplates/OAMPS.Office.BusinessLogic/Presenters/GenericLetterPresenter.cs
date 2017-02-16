using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;
using OAMPS.Office.BusinessLogic.Models.Wizards;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Models.Template;

namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class GenericLetterPresenter : BasePresenter
    {
        public GenericLetterPresenter(IDocument document, IBaseView view) : base(document, view)
        {
                        
        }

        public void DeleteDocumentHeaderAndFooter()
        {
            Document.RemoveHeader();
            Document.RemoveFooter();
            Document.DeleteImage(BusinessLogic.Helpers.Constants.ImageProperties.CompanyLogo);

        }

        public void InsertSubFragments(string attachmentBookMark, string mainContentBookMark, List<DocumentFragment> bookmarks, string fragUrl)
        {
            string frgKeys = null;


            if (Document.MoveCursorToStartOfBookmark(attachmentBookMark))
            {
                bookmarks.ForEach(
                    (x) =>
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

            bookmarks.ForEach((x) =>
            {
                Document.InsertFile(x.Url);
                if (x != bookmarks.LastOrDefault())
                    Document.InsertPageBreak();
            });
            d.TurnOnProtection(protection, string.Empty);
            d.CloseInformationPanel(true);

        }
    }
}
