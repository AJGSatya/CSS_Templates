using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;
using OAMPS.Office.BusinessLogic.Presenters.SharePoint;


namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class HelpContentPresenter : BasePresenter
    {
        public HelpContentPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {                        
        }

        public string FindHeadingTextForCurrentDocument()
        {
            return Document.FindTextByStyleForCurrentDocument(BusinessLogic.Helpers.Constants.WordStyles.Heading1);
        }

       
    }
}
