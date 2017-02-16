using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Helpers;
using OAMPS.Office.BusinessLogic.Interfaces;
using OAMPS.Office.BusinessLogic.Interfaces.Word;

namespace OAMPS.Office.BusinessLogic.Presenters
{
    public class AgendaPresenter : BasePresenter
    {
        public AgendaPresenter(IDocument document, IBaseView view)
            : base(document, view)
        {                        
        }

        public void InsertMinutesFragement(string path)
        {
            if (Document.MoveCursorToStartOfBookmark(Constants.WordBookmarks.AgendaMinutes))
            {
                Document.InsertFile(path);    
            }
            
        }
    }
}
