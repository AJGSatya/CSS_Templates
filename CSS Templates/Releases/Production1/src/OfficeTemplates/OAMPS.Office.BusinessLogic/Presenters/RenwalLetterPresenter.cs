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
    public class RenwalLetterPresenter : BasePresenter
    {
        public RenwalLetterPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {                        
        }

        public void InsertSubFragments(string attachmentBookMark, string mainContentBookMark, List<DocumentFragment> bookmarks)
        {
            Document.MoveCursorToStartOfBookmark(attachmentBookMark);
            string frgKeys = null;
            bookmarks.ForEach(
                (x) =>
                    {
                        
                        Document.TypeText(x.Title);
                        Document.TypeText("; ");
                        frgKeys += x.Key + ";";                        
                    }
                );
            Document.MoveCursorToStartOfBookmark(mainContentBookMark);
            
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlSubFragments, frgKeys);

            bookmarks.ForEach((x) => Document.OpenFile(x.Url));
        }

        public void DeleteDocumentHeaderAndFooter()
        {
            Document.RemoveHeader();
            Document.RemoveFooter();
        }

        public void InsertMainFragment(string bookmark, string fragUrl, RenewalLetter template)
        {
            Document.MoveCursorToStartOfBookmark(bookmark);

            Document.InsertFile(fragUrl);

            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkContacted,
                                                 template.IsContactSelected.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkNewClient,
                                                 template.IsNewClientSelected.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkFunding,
                                                 template.IsFundingSelected.ToString());
            Document.UpdateOrCreatePropertyValue(Constants.WordDocumentProperties.RlChkGAW,
                                                 template.IsWarningSelected.ToString());
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

        public void PopulatePolicy(IPolicyClass lPolicy)
        {
            if(lPolicy!=null)
            Document.PopulateControl(Constants.WordDocumentProperties.RenewalPolicy, lPolicy.Title);
            
        }      

        public string ReadPolicyReference()
        {
            return Document.ReadContentControlValue(Constants.WordDocumentProperties.RenewalPolicy);
        }
    }
}
